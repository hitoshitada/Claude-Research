"""
Weekly Intelligence Report 生成モジュール

パイプライン:
  1. Article リストから articles_data.json を構築
  2. Gemini API で記事分析 → front_report.json を生成
  3. reportlab で フロントレポートPDF を生成
  4. pypdf で フロントPDF + 既存記事PDF を連結 → 最終レポート
"""

import json
import os
import re
from pathlib import Path
from typing import Optional, Callable

from google import genai
from google.genai import types as genai_types

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus.flowables import Flowable

from pypdf import PdfWriter, PdfReader

# ---------------------------------------------------------------------------
# パス定数
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent.parent
PROMPTS_DIR = BASE_DIR / "prompts"

GEMINI_MODEL = "gemini-2.5-flash"  # fallback handled in generate_front_report_json

# ---------------------------------------------------------------------------
# フォント登録 (Windows: メイリオ優先)
# ---------------------------------------------------------------------------
_FONT_REGISTERED = False

def _ensure_fonts():
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return
    candidates = [
        ("C:/Windows/Fonts/meiryo.ttc", 0),
        ("C:/Windows/Fonts/msgothic.ttc", 0),
    ]
    for fpath, idx in candidates:
        if os.path.exists(fpath):
            pdfmetrics.registerFont(TTFont("JP", fpath, subfontIndex=idx))
            # プロポーショナル版 (太字代用)
            bold_path = fpath.replace("meiryo.ttc", "meiryob.ttc")
            if os.path.exists(bold_path):
                pdfmetrics.registerFont(TTFont("JPB", bold_path, subfontIndex=idx))
            else:
                pdfmetrics.registerFont(TTFont("JPB", fpath, subfontIndex=idx))
            _FONT_REGISTERED = True
            return
    # フォールバック
    pdfmetrics.registerFont(TTFont("JP", "Helvetica"))
    pdfmetrics.registerFont(TTFont("JPB", "Helvetica"))
    _FONT_REGISTERED = True


# ===================================================================
# Step 1: Article → articles_data.json
# ===================================================================

def build_articles_data(articles, topic_name: str, collection_date: str) -> dict:
    """Article リストを中間 JSON 形式に変換する。

    Parameters
    ----------
    articles : list[Article]  (article_parser.Article)
    topic_name : str  例: "接着・封止材"
    collection_date : str  例: "2026-04-11" or "2026年04月11日"
    """
    # 日付を YYYY-MM-DD に正規化
    date_str = collection_date
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", collection_date)
    if m:
        date_str = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"

    countries = sorted(set(a.country for a in articles if a.country))

    arts = []
    for i, a in enumerate(articles, 1):
        arts.append({
            "id": f"#{i:02d}",
            "title": a.title_ja,
            "source": a.source_name,
            "country": a.country,
            "date": a.publish_date,
            "url": a.url,
            "summary": a.summary_ja,
            "details": a.content_ja,
        })

    return {
        "category": topic_name,
        "collection_date": date_str,
        "total_articles": len(articles),
        "countries": countries,
        "articles": arts,
    }


# ===================================================================
# Step 2: Gemini 分析 → front_report.json
# ===================================================================

def _load_prompt(filename: str) -> str:
    path = PROMPTS_DIR / filename
    return path.read_text(encoding="utf-8").strip()


def generate_front_report_json(
    client: genai.Client,
    articles_data: dict,
    progress_callback: Optional[Callable] = None,
) -> dict:
    """Gemini API を呼び出してフロントレポート用 JSON を生成する。"""

    system_prompt = _load_prompt("weekly_report_system.md")
    user_template = _load_prompt("weekly_report_user.md")

    # テンプレートに値を埋め込み
    articles_json_str = json.dumps(articles_data["articles"], ensure_ascii=False, indent=2)
    user_prompt = (
        user_template
        .replace("{category}", articles_data["category"])
        .replace("{collection_date}", articles_data["collection_date"])
        .replace("{total_articles}", str(articles_data["total_articles"]))
        .replace("{countries}", ", ".join(articles_data["countries"]))
        .replace("{articles_json}", articles_json_str)
    )

    if progress_callback:
        progress_callback("Gemini APIで記事分析中...")

    # gemini-1.5-flash は v1beta API で廃止済み（404 NOT_FOUND）のため除外。
    # 429/503 は一時的エラー → 同モデルでリトライ。404 は恒久エラー → 次モデルへ。
    _FALLBACK_MODELS = [
        GEMINI_MODEL,          # gemini-2.5-flash
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
    ]

    import time as _time

    def _is_transient(exc: Exception) -> bool:
        """429 / 503 / RESOURCE_EXHAUSTED など一時的エラーかどうか判定"""
        s = str(exc)
        return any(k in s for k in ("429", "503", "RESOURCE_EXHAUSTED",
                                    "overloaded", "rate", "quota", "too many"))

    def _is_not_found(exc: Exception) -> bool:
        """404 / NOT_FOUND = モデル廃止・未対応 → リトライ無意味"""
        s = str(exc)
        return any(k in s for k in ("404", "NOT_FOUND", "not found",
                                    "not supported"))

    last_err = None
    for _model in _FALLBACK_MODELS:
        # 一時的エラーに対して最大 3 回リトライ（指数バックオフ）
        for _attempt in range(3):
            try:
                if progress_callback and _attempt > 0:
                    progress_callback(
                        f"Gemini API リトライ中... ({_model} 試行 {_attempt + 1}/3)")
                response = client.models.generate_content(
                    model=_model,
                    contents=user_prompt,
                    config=genai_types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        temperature=0.3,
                    ),
                )
                text = response.text.strip()
                # JSON部分を抽出 (```json ... ``` で囲まれている場合のフォールバック)
                if text.startswith("```"):
                    text = re.sub(r"^```(?:json)?\s*", "", text)
                    text = re.sub(r"\s*```$", "", text)
                return json.loads(text)
            except Exception as e:
                last_err = e
                if _is_not_found(e):
                    # モデル廃止など恒久エラー → 次モデルへ
                    if progress_callback:
                        progress_callback(
                            f"  {_model} は利用不可（404）。次のモデルを試します...")
                    break
                if _is_transient(e) and _attempt < 2:
                    wait = 10 * (2 ** _attempt)  # 10s → 20s
                    if progress_callback:
                        progress_callback(
                            f"  一時的エラー ({_model})。{wait}秒後にリトライします...")
                    _time.sleep(wait)
                    continue
                # その他エラーまたはリトライ上限 → 次モデルへ
                break

    raise RuntimeError(f"Gemini API呼び出し失敗（全モデル試行済み）: {last_err}")


# ===================================================================
# Step 3: reportlab で フロントレポート PDF 生成
# ===================================================================

# --- Colors ---
NAVY = HexColor("#1B2A4A")
DBLUE = HexColor("#2C3E6B")
LG = HexColor("#F5F6F8")
MG = HexColor("#8A8FA0")
TXT = HexColor("#3A3A3A")
W = HexColor("#FFFFFF")
RED = HexColor("#C0392B")
GRN = HexColor("#27AE60")
TEAL = HexColor("#1ABC9C")

PW, PH = A4
ML = 18 * mm
MR = 18 * mm
MT = 20 * mm
MB = 18 * mm
CW = PW - ML - MR

# --- Styles (lazy init) ---
_S = None

def _get_styles(acc_color):
    """アクセントカラーを反映したスタイル辞書を生成。"""
    ACC = acc_color
    ACC_LT = HexColor(_lighten_hex(ACC.hexval(), 0.85))

    s = {}
    s['body'] = ParagraphStyle('body', fontName='JP', fontSize=9, textColor=TXT,
                                leading=16, alignment=TA_JUSTIFY, spaceAfter=2 * mm)
    s['bs'] = ParagraphStyle('bs', fontName='JP', fontSize=8, textColor=MG, leading=13, spaceAfter=1 * mm)
    s['exp'] = ParagraphStyle('exp', fontName='JP', fontSize=9, textColor=DBLUE,
                               leading=15, leftIndent=8 * mm, rightIndent=4 * mm,
                               spaceBefore=2 * mm, spaceAfter=3 * mm)
    s['qt'] = ParagraphStyle('qt', fontName='JP', fontSize=10, textColor=NAVY,
                              leading=16, spaceBefore=3 * mm, spaceAfter=1 * mm, leftIndent=6 * mm)
    s['qb'] = ParagraphStyle('qb', fontName='JP', fontSize=8.5, textColor=TXT,
                              leading=14, spaceAfter=3 * mm, leftIndent=6 * mm)
    s['sub'] = ParagraphStyle('sub', fontName='JP', fontSize=11, textColor=DBLUE,
                               leading=16, spaceBefore=5 * mm, spaceAfter=2 * mm)
    s['act'] = ParagraphStyle('act', fontName='JP', fontSize=9, textColor=TXT,
                               leading=15, spaceAfter=2 * mm, leftIndent=8 * mm)
    s['foot'] = ParagraphStyle('foot', fontName='JP', fontSize=7, textColor=MG, alignment=TA_CENTER)
    s['tc'] = ParagraphStyle('tc', fontName='JP', fontSize=6.5, textColor=TXT, leading=10)
    s['tcb'] = ParagraphStyle('tcb', fontName='JP', fontSize=6.5, textColor=NAVY, leading=10)
    s['ACC'] = ACC
    s['ACC_LT'] = ACC_LT
    return s


def _lighten_hex(hex_str: str, factor: float = 0.85) -> str:
    """Hex カラーを明るくする。factor=1.0 で白。"""
    hex_str = hex_str.lstrip("#").lstrip("0x").lstrip("0X")
    # HexColor.hexval() が '0xRRGGBB' 形式を返す場合の対応
    if len(hex_str) > 6:
        hex_str = hex_str[-6:]
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02X}{g:02X}{b:02X}"


# --- Flowables ---

class _Header(Flowable):
    def __init__(self, w, data, acc):
        Flowable.__init__(self)
        self.width = w
        self.height = 38 * mm
        self.data = data
        self.acc = acc

    def draw(self):
        c = self.canv
        d = self.data
        c.setFillColor(NAVY)
        c.rect(-ML, -4 * mm, PW, self.height + 4 * mm, fill=1, stroke=0)
        c.setFillColor(self.acc)
        c.rect(-ML, -4 * mm, 4 * mm, self.height + 4 * mm, fill=1, stroke=0)

        c.setFont("JP", 22)
        c.setFillColor(W)
        c.drawString(4 * mm, self.height - 11 * mm, d.get("category", ""))

        c.setFont("JP", 13)
        c.drawString(4 * mm, self.height - 19 * mm, "Weekly Intelligence Report")

        c.setFont("JP", 9)
        c.setFillColor(HexColor("#B0BEC5"))
        meta_line = f'{d.get("date", "")} | {d.get("total_articles", "")}件 | {d.get("countries_count", "")}カ国'
        c.drawString(4 * mm, self.height - 27 * mm, meta_line)

        c.setFont("JP", 7)
        c.drawString(4 * mm, self.height - 32 * mm, "troy-technical.jp")

        # 右側: キーワード
        rx = self.width - 58 * mm
        c.setFont("JP", 10)
        c.setFillColor(self.acc)
        c.drawString(rx, self.height - 11 * mm, "今週のキーワード")
        c.setFont("JP", 14)
        c.setFillColor(W)
        kw = d.get("headline_keyword", "")
        c.drawString(rx, self.height - 20 * mm, kw)
        c.setFont("JP", 7)
        c.setFillColor(HexColor("#B0BEC5"))
        c.drawString(rx, self.height - 26 * mm, d.get("headline_sub", ""))


class _Div(Flowable):
    def __init__(self, w, text, color=NAVY):
        Flowable.__init__(self)
        self.width = w
        self.height = 8 * mm
        self.text = text
        self.color = color

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.rect(0, 0, 3 * mm, self.height, fill=1, stroke=0)
        c.setFillColor(LG)
        c.rect(4 * mm, 0, self.width - 4 * mm, self.height, fill=1, stroke=0)
        c.setFont("JP", 11)
        c.setFillColor(self.color)
        c.drawString(7 * mm, 2.2 * mm, self.text)


class _ExpBox(Flowable):
    def __init__(self, w, text, label, acc, acc_lt, styles):
        Flowable.__init__(self)
        self.width = w
        self.text = text
        self.label = label
        self.acc = acc
        self.acc_lt = acc_lt
        self._p = Paragraph(text, styles['exp'])
        _, self._ph = self._p.wrap(w - 14 * mm, 500)
        self.height = self._ph + 12 * mm

    def draw(self):
        c = self.canv
        c.setFillColor(self.acc_lt)
        c.roundRect(0, 0, self.width, self.height, 2 * mm, fill=1, stroke=0)
        c.setFillColor(self.acc)
        c.rect(0, 0, 3 * mm, self.height, fill=1, stroke=0)
        c.setFont("JP", 7.5)
        c.setFillColor(self.acc)
        c.drawString(6 * mm, self.height - 5 * mm, f"▶ {self.label}")
        self._p.drawOn(c, 5 * mm, 2 * mm)


class _Metrics(Flowable):
    def __init__(self, w, ms, acc):
        Flowable.__init__(self)
        self.width = w
        self.ms = ms
        self.acc = acc
        self.height = 18 * mm

    def draw(self):
        c = self.canv
        n = len(self.ms)
        if n == 0:
            return
        cw = (self.width - (n - 1) * 2.5 * mm) / n
        for i, m in enumerate(self.ms):
            x = i * (cw + 2.5 * mm)
            c.setFillColor(LG)
            c.roundRect(x, 0, cw, self.height, 2 * mm, fill=1, stroke=0)
            c.setFont("JP", 14)
            c.setFillColor(self.acc)
            c.drawCentredString(x + cw / 2, 9 * mm, str(m.get("value", "")))
            c.setFont("JP", 6)
            c.setFillColor(MG)
            c.drawCentredString(x + cw / 2, 5.5 * mm, str(m.get("unit", "")))
            c.setFont("JP", 6)
            c.setFillColor(TXT)
            c.drawCentredString(x + cw / 2, 1.5 * mm, str(m.get("label", "")))


class _OTMatrix(Flowable):
    """機会 vs 脅威マトリクス。
    マトリクス上は「色付き円＋短いラベル」のみ描画。
    機会/脅威の詳細テキストは下の凡例テーブルに委ねることで重なりを防ぐ。
    """
    def __init__(self, w, items, acc):
        Flowable.__init__(self)
        self.width = w
        self.height = 100 * mm   # 少し高くして余裕を持たせる
        self.items = items
        self.acc = acc

    def draw(self):
        c = self.canv
        c.setFillColor(LG)
        c.roundRect(0, 0, self.width, self.height, 2 * mm, fill=1, stroke=0)
        c.setFont("JP", 9)
        c.setFillColor(NAVY)
        c.drawString(4 * mm, self.height - 6 * mm, "日本企業にとっての「機会 vs 脅威」マトリクス")

        mx, my = 12 * mm, 10 * mm
        mw = self.width - 24 * mm
        mh = self.height - 24 * mm
        hw, hh = mw / 2, mh / 2

        # 四象限の背景
        c.setFillColor(HexColor("#E8F5E9"))
        c.rect(mx, my + hh, hw, hh, fill=1, stroke=0)
        c.setFillColor(HexColor("#FFF8E1"))
        c.rect(mx + hw, my + hh, hw, hh, fill=1, stroke=0)
        c.setFillColor(W)
        c.rect(mx, my, hw, hh, fill=1, stroke=0)
        c.setFillColor(HexColor("#FFEBEE"))
        c.rect(mx + hw, my, hw, hh, fill=1, stroke=0)

        # グリッド線
        c.setStrokeColor(HexColor("#B0BEC5"))
        c.setLineWidth(1)
        c.line(mx, my + hh, mx + mw, my + hh)
        c.line(mx + hw, my, mx + hw, my + mh)
        c.setLineWidth(0.5)
        c.rect(mx, my, mw, mh, fill=0, stroke=1)

        # 象限ラベル（各象限の上部に半透明の帯で描画し、円と重ならないよう視認性を確保）
        c.setFont("JP", 6.5)
        label_band_h = 5.2 * mm  # 象限ラベル用のヘッダー帯の高さ
        # 各象限の上部に白い半透明帯
        c.setFillColor(HexColor("#FFFFFF"))
        c.setFillAlpha(0.75)
        c.rect(mx,        my + mh - label_band_h, hw, label_band_h, fill=1, stroke=0)
        c.rect(mx + hw,   my + mh - label_band_h, hw, label_band_h, fill=1, stroke=0)
        c.rect(mx,        my + hh - label_band_h, hw, label_band_h, fill=1, stroke=0)
        c.rect(mx + hw,   my + hh - label_band_h, hw, label_band_h, fill=1, stroke=0)
        c.setFillAlpha(1.0)
        # ラベル本体（帯の中に描画）
        label_y_top = my + mh - label_band_h + 1.4 * mm
        label_y_mid = my + hh - label_band_h + 1.4 * mm
        c.setFillColor(HexColor("#4CAF50"))
        c.drawString(mx + 2 * mm, label_y_top, "機会大・脅威小")
        c.setFillColor(HexColor("#E67E22"))
        c.drawString(mx + hw + 2 * mm, label_y_top, "機会大・脅威大")
        c.setFillColor(MG)
        c.drawString(mx + 2 * mm, label_y_mid, "影響小（参考）")
        c.setFillColor(HexColor("#E53935"))
        c.drawString(mx + hw + 2 * mm, label_y_mid, "脅威大・機会小")

        # 軸ラベル
        c.setFont("JP", 7)
        c.setFillColor(MG)
        c.saveState()
        c.translate(mx - 6 * mm, my + mh / 2)
        c.rotate(90)
        c.drawCentredString(0, 0, "← 機会 →")
        c.restoreState()
        c.drawCentredString(mx + mw / 2, my - 5 * mm, "← 脅威 →")

        # ── アイテムのプロット（2パス）──
        # Pass 1: 全円を描画
        # Pass 2: 全ラベルを全円の上に描画
        # → 後から描いた円が先のラベルを隠す問題を防ぐ
        COLOR_MAP = {
            "red": RED, "orange": self.acc, "teal": TEAL,
            "green": GRN, "gray": MG,
        }
        R_DOT = 4.5 * mm   # 全アイテム同サイズの円
        LABEL_H = 5.0 * mm  # ラベルボックス高さ
        LABEL_FONT = 7.5    # ラベルフォントサイズ

        # ラベルは円の上に配置するため、上端余白を追加
        # 下端は円半径のみ（ラベルは下に出ない）
        y_top_pad = (label_band_h + R_DOT + LABEL_H + 2.0 * mm) / hh
        y_bot_pad = (R_DOT + 1.5 * mm) / hh
        x_pad     = (R_DOT + 1.5 * mm) / hw
        yf_min_clamp = max(0.08, y_bot_pad)
        yf_max_clamp = min(0.92, 1.0 - y_top_pad)
        xf_min_clamp = max(0.05, x_pad)
        xf_max_clamp = min(0.95, 1.0 - x_pad)

        # ── Pass 1: 座標計算 ＋ 円描画 ──
        from reportlab.pdfbase.pdfmetrics import stringWidth
        placed: list[tuple] = []   # (px, py, label, color, by)
        placed_xy: list[tuple[float, float]] = []
        for item in self.items:
            quad  = item.get("quadrant", "BL")
            xf    = max(xf_min_clamp, min(xf_max_clamp, item.get("x_position", 0.5)))
            yf    = max(yf_min_clamp, min(yf_max_clamp, item.get("y_position", 0.5)))
            label = item.get("label", "")
            color = COLOR_MAP.get(item.get("color", "gray"), MG)

            bx = mx + (hw if quad[1] == 'R' else 0)
            by = my + (hh if quad[0] == 'T' else 0)
            px_pos = bx + xf * hw
            py_pos = by + yf * hh

            # 衝突回避（最低 2*R_DOT+4mm 離す）
            min_dist = 2 * R_DOT + 4 * mm
            for _ in range(10):
                collision = any(
                    (px_pos - ox) ** 2 + (py_pos - oy) ** 2 < min_dist ** 2
                    for ox, oy in placed_xy
                )
                if not collision:
                    break
                py_pos -= 3 * mm
                if py_pos < by + yf_min_clamp * hh:
                    py_pos = by + yf_min_clamp * hh
                    px_pos += 4 * mm
                if px_pos > bx + xf_max_clamp * hw:
                    px_pos = bx + xf_max_clamp * hw

            placed_xy.append((px_pos, py_pos))
            placed.append((px_pos, py_pos, label, color, by))

            # 影 → 白縁 → 塗り
            c.setFillColor(HexColor("#CCCCCC"))
            c.circle(px_pos + 0.6 * mm, py_pos - 0.6 * mm, R_DOT, fill=1, stroke=0)
            c.setFillColor(W)
            c.circle(px_pos, py_pos, R_DOT + 0.8 * mm, fill=1, stroke=0)
            c.setFillColor(color)
            c.circle(px_pos, py_pos, R_DOT, fill=1, stroke=0)

        # ── Pass 2: 全ラベルを最前面に描画 ──
        # ラベルは円の上方に配置し、象限ヘッダー帯を超えないようクランプ
        for (px_pos, py_pos, label, color, by) in placed:
            # 象限ヘッダー帯の下端（ここより上にラベルを出してはいけない）
            quad_ceiling = by + hh - label_band_h - 0.5 * mm

            # ラベルのY位置（ベースライン）: 円の上に R_DOT + 1.5mm
            label_y = py_pos + R_DOT + 1.5 * mm
            # ボックス上端がヘッダー帯に入り込まないようクランプ
            box_top = label_y - 0.8 * mm + LABEL_H
            if box_top > quad_ceiling:
                label_y = quad_ceiling - LABEL_H + 0.8 * mm

            lw = stringWidth(label, "JP", LABEL_FONT) + 3 * mm

            # 白背景＋アクセントカラーの枠線（コントラスト確保）
            c.setFillColor(W)
            c.setStrokeColor(color)
            c.setLineWidth(0.8)
            c.roundRect(px_pos - lw / 2, label_y - 0.8 * mm, lw, LABEL_H,
                        1.2 * mm, fill=1, stroke=1)
            c.setLineWidth(0.3)

            # ネイビーで文字
            c.setFont("JP", LABEL_FONT)
            c.setFillColor(NAVY)
            c.drawCentredString(px_pos, label_y, label)


# --- Helper functions ---

def _build_ot_legend(items, acc):
    """OTマトリクスの凡例テーブル（マトリクス下に配置）。
    各アイテムの機会・脅威の詳細をコンパクトなテーブルで表示する。
    """
    COLOR_MAP = {
        "red": RED, "orange": acc, "teal": TEAL,
        "green": GRN, "gray": MG,
    }
    QUAD_LABEL = {"TL": "機会大", "TR": "注意", "BL": "参考", "BR": "脅威大"}

    hs = ParagraphStyle('oth', fontName='JP', fontSize=5.5, textColor=W, leading=8, alignment=TA_CENTER)
    tc = ParagraphStyle('otc', fontName='JP', fontSize=6.5, textColor=TXT, leading=10)
    tg = ParagraphStyle('otg', fontName='JP', fontSize=6.5, textColor=HexColor("#27AE60"), leading=10)
    tr = ParagraphStyle('otr', fontName='JP', fontSize=6.5, textColor=HexColor("#C0392B"), leading=10)

    header = [
        Paragraph('項目', hs),
        Paragraph('象限', hs),
        Paragraph('↑ 機会', hs),
        Paragraph('↓ 脅威', hs),
    ]
    rows = [header]

    for item in items:
        quad = item.get("quadrant", "BL")
        label = item.get("label", "")
        opp = item.get("opportunity", "")
        thr = item.get("threat", "")
        color = COLOR_MAP.get(item.get("color", "gray"), MG)

        # ラベルセルに色付き丸を表現（文字で代用）
        # hexval() は '0xRRGGBB' 形式なので '#RRGGBB' に変換
        hex_raw = color.hexval().lstrip("0x").lstrip("0X")
        hex_str = f"#{hex_raw[-6:].upper()}"
        label_cell = Paragraph(
            f'<font color="{hex_str}">\u25cf</font> {label}',
            ParagraphStyle('otl', fontName='JP', fontSize=6.5, textColor=TXT, leading=10)
        )
        rows.append([
            label_cell,
            Paragraph(QUAD_LABEL.get(quad, quad), tc),
            Paragraph(opp if opp else "—", tg if opp else tc),
            Paragraph(thr if thr else "—", tr if thr else tc),
        ])

    col_w = [30 * mm, 14 * mm, CW / 2 - 22 * mm, CW / 2 - 22 * mm]
    t = Table(rows, colWidths=col_w)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('TEXTCOLOR', (0, 0), (-1, 0), W),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [W, LG]),
        ('GRID', (0, 0), (-1, -1), 0.3, HexColor("#D0D4DE")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 1.5 * mm),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5 * mm),
        ('LEFTPADDING', (0, 0), (-1, -1), 2 * mm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2 * mm),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
    ]))
    return t


def _score_dots(n, max_n=5):
    filled = "●" * n
    empty = "○" * (max_n - n)
    if n >= 4:
        color = "#C0392B"
    elif n >= 3:
        color = "#E8913A"
    elif n >= 2:
        color = "#8A8FA0"
    else:
        color = "#C0C4CE"
    return Paragraph(
        f'<font color="{color}" size="7">{filled}</font>'
        f'<font color="#D0D4DE" size="7">{empty}</font>',
        ParagraphStyle('dots', fontName='JP', fontSize=7, alignment=TA_CENTER, leading=10))


def _build_triage_table(evaluations, S):
    hs = ParagraphStyle('th', fontName='JP', fontSize=5.5, textColor=W, leading=8, alignment=TA_CENTER)
    hl = ParagraphStyle('thl', fontName='JP', fontSize=5.5, textColor=W, leading=8)
    header = [
        Paragraph('#', hs), Paragraph('記事タイトル', hl), Paragraph('種別', hs),
        Paragraph('技術<br/>新規性', hs), Paragraph('実用化<br/>距離', hs),
        Paragraph('市場<br/>インパクト', hs), Paragraph('データ<br/>信頼性', hs),
        Paragraph('日本<br/>関連度', hs), Paragraph('一行サマリ', hl),
    ]
    rows = [header]
    highlight_rows = []
    for idx, ev in enumerate(evaluations):
        row_num = idx + 1
        rows.append([
            Paragraph(f'<b>{ev["id"]}</b>', S['tcb']),
            Paragraph(ev.get("title", ""), S['tcb']),
            Paragraph(ev.get("type", ""),
                      ParagraphStyle('at', fontName='JP', fontSize=6, textColor=MG,
                                     leading=9, alignment=TA_CENTER)),
            _score_dots(ev.get("tech_novelty", 1)),
            _score_dots(ev.get("proximity", 1)),
            _score_dots(ev.get("market_impact", 1)),
            _score_dots(ev.get("data_reliability", 1)),
            _score_dots(ev.get("japan_relevance", 1)),
            Paragraph(ev.get("one_line_summary", ""), S['tc']),
        ])
        if ev.get("is_highlight"):
            highlight_rows.append(row_num)

    col_w = [8 * mm, 26 * mm, 12 * mm, 14 * mm, 14 * mm, 14 * mm, 14 * mm, 14 * mm, CW - 116 * mm]
    t = Table(rows, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('TEXTCOLOR', (0, 0), (-1, 0), W),
        ('BACKGROUND', (0, 1), (-1, -1), W),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [W, LG]),
        ('GRID', (0, 0), (-1, -1), 0.3, HexColor("#D0D4DE")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 1.2 * mm),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1.2 * mm),
        ('LEFTPADDING', (0, 0), (-1, -1), 1 * mm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 1 * mm),
        ('ALIGN', (2, 0), (7, -1), 'CENTER'),
    ]
    for r in highlight_rows:
        style_cmds.append(('BACKGROUND', (0, r), (-1, r), HexColor("#FFF8E1")))
    t.setStyle(TableStyle(style_cmds))
    return t


def _build_legend():
    return Paragraph(
        '<font color="#C0392B" size="7">●●●●</font><font color="#D0D4DE" size="7">○</font> 高　'
        '<font color="#E8913A" size="7">●●●</font><font color="#D0D4DE" size="7">○○</font> 中高　'
        '<font color="#8A8FA0" size="7">●●</font><font color="#D0D4DE" size="7">○○○</font> 中　'
        '<font color="#C0C4CE" size="7">●</font><font color="#D0D4DE" size="7">○○○○</font> 低　'
        '| 背景黄色＝注目記事',
        ParagraphStyle('lg', fontName='JP', fontSize=7, textColor=MG, leading=11))


def _build_reading_guide():
    return Paragraph(
        "各列の見方 ― "
        "<b>技術新規性</b>：ブレークスルー度合い　"
        "<b>実用化距離</b>：製品として使える近さ　"
        "<b>市場インパクト</b>：業界全体への影響規模　"
        "<b>データ信頼性</b>：定量データ・査読の有無　"
        "<b>日本関連度</b>：日本の企業・サプライチェーンとの直接的関連性",
        ParagraphStyle('rg', fontName='JP', fontSize=7, textColor=TXT, leading=11, spaceAfter=2 * mm))


# --- Story builder ---

def _build_story(data: dict):
    """front_report.json データからreportlab Storyを構築。"""
    meta = data.get("report_metadata", {})
    acc_hex = meta.get("accent_color", "#E8913A")
    ACC = HexColor(acc_hex)
    S = _get_styles(ACC)

    category = meta.get("category", "")
    date_str = meta.get("date", "")
    total = meta.get("total_articles", 0)

    story = []

    # === P1: Header + Evaluation Matrix ===
    story.append(_Header(CW, meta, ACC))
    story.append(Spacer(1, 4 * mm))

    # Key Metrics
    metrics = data.get("key_metrics", [])
    if metrics:
        story.append(_Metrics(CW, metrics, ACC))
        story.append(Spacer(1, 4 * mm))

    # Evaluation table
    evals = data.get("article_evaluations", [])
    story.append(_Div(CW, f"今週の全{total}記事 ― 5軸評価で読むべき記事を選ぶ"))
    story.append(Spacer(1, 1 * mm))
    story.append(_build_reading_guide())
    story.append(_build_triage_table(evals, S))
    story.append(Spacer(1, 1 * mm))
    story.append(_build_legend())

    story.append(PageBreak())

    # === P2: 3 Questions + OT Matrix ===
    story.append(_Div(CW, "今週、判断に影響しうる3つの問い"))
    story.append(Spacer(1, 3 * mm))

    for q in data.get("three_questions", []):
        story.append(Paragraph(q.get("title", ""), S['qt']))
        story.append(Paragraph(q.get("body", ""), S['qb']))

    story.append(Spacer(1, 2 * mm))
    story.append(_Div(CW, "日本企業にとっての「機会 vs 脅威」", ACC))
    story.append(Spacer(1, 3 * mm))

    ot_items = data.get("opportunity_threat", [])
    story.append(_OTMatrix(CW, ot_items, ACC))
    story.append(Spacer(1, 2 * mm))
    # 凡例テーブル（機会・脅威の詳細はマトリクス下のテーブルに表示）
    if ot_items:
        story.append(_build_ot_legend(ot_items, ACC))

    story.append(PageBreak())

    # === P3+P4: Deep Dives ===
    deep_dives = data.get("deep_dives", [])
    for i, dd in enumerate(deep_dives):
        n = i + 1
        story.append(_Div(CW, f"深掘り {'①②③④⑤'[i]} ― {dd.get('section_title', '')}"))
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(
            f'<font color="#8A8FA0">{dd.get("source_line", "")} | {dd.get("score_line", "")}</font>',
            S['bs']))

        for para in dd.get("body_paragraphs", []):
            story.append(Paragraph(para, S['body']))

        story.append(_ExpBox(CW, dd.get("expert_comment", ""),
                             dd.get("expert_label", "技術者の視点"),
                             S['ACC'], S['ACC_LT'], S))

        # 2件目の後にページブレーク
        if i == 1 and len(deep_dives) > 2:
            story.append(PageBreak())
        else:
            story.append(Spacer(1, 4 * mm))

    # Other notable
    other = data.get("other_notable", [])
    if other:
        story.append(Spacer(1, 2 * mm))
        story.append(_Div(CW, "その他の注目記事"))
        story.append(Spacer(1, 2 * mm))
        for item in other:
            story.append(Paragraph(f'<b>{item.get("title", "")}</b>', S['tcb']))
            story.append(Paragraph(
                f'<font color="#8A8FA0">{item.get("score_line", "")}</font>', S['bs']))
            story.append(Paragraph(item.get("comment", ""), S['body']))

    story.append(PageBreak())

    # === Last page: Actions ===
    story.append(_Div(CW, "今週のアクション提案", ACC))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "記事評価マトリクスと機会/脅威分析を踏まえたアクション提案です。", S['body']))

    ACTION_COLORS = {"red": RED, "orange": ACC, "blue": DBLUE}
    for action in data.get("action_items", []):
        color = ACTION_COLORS.get(action.get("color", "blue"), DBLUE)
        ps = ParagraphStyle('ap', parent=S['sub'], textColor=color)
        story.append(Paragraph(f"▍{action.get('timeframe', '')}", ps))
        for item in action.get("items", []):
            story.append(Paragraph(f"• {item}", S['act']))

    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=MG))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        f"troy-technical.jp 独自キュレーション。記事著作権は各原著作者に帰属。"
        f" | Gemini API + Claude | {date_str}",
        S['foot']))

    return story


def generate_front_report_pdf(front_report_data: dict, output_path: Path) -> Path:
    """フロントレポート JSON → PDF を生成して返す。"""
    _ensure_fonts()

    meta = front_report_data.get("report_metadata", {})
    acc_hex = meta.get("accent_color", "#E8913A")
    ACC = HexColor(acc_hex)
    category = meta.get("category", "")
    date_str = meta.get("date", "")

    doc = SimpleDocTemplate(
        str(output_path), pagesize=A4,
        leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB,
    )

    def page_footer(canvas, doc_obj):
        canvas.saveState()
        canvas.setFont("JP", 7)
        canvas.setFillColor(MG)
        canvas.drawCentredString(
            PW / 2, 10 * mm,
            f"{category} Weekly Intelligence Report | {date_str} | Page {doc_obj.page}")
        canvas.setStrokeColor(ACC)
        canvas.setLineWidth(2)
        canvas.line(ML, PH - 8 * mm, PW - MR, PH - 8 * mm)
        canvas.restoreState()

    story = _build_story(front_report_data)
    doc.build(story, onFirstPage=page_footer, onLaterPages=page_footer)
    return output_path


# ===================================================================
# Step 4: PDF 連結
# ===================================================================

def concatenate_pdfs(front_pdf: Path, articles_pdf: Path, output_path: Path) -> Path:
    """フロントレポートPDF + 全記事PDF を連結。"""
    writer = PdfWriter()

    for pdf_path in [front_pdf, articles_pdf]:
        if pdf_path.exists():
            reader = PdfReader(str(pdf_path))
            for page in reader.pages:
                writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


# ===================================================================
# 統合エントリポイント
# ===================================================================

def generate_weekly_report(
    client: genai.Client,
    articles,
    topic_name: str,
    output_folder: Path,
    collection_date: str,
    existing_articles_pdf: Optional[Path] = None,
    progress_callback: Optional[Callable] = None,
) -> Path:
    """WeeklyReport の全ステップを実行する。

    Returns
    -------
    Path : 最終レポート PDF のパス
    """
    if progress_callback:
        progress_callback("WeeklyReport: articles_data.json 生成中...")

    # Step 1: articles_data.json
    articles_data = build_articles_data(articles, topic_name, collection_date)
    json_path = output_folder / "articles_data.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(articles_data, f, ensure_ascii=False, indent=2)

    # Step 2: Gemini 分析
    if progress_callback:
        progress_callback("WeeklyReport: Gemini 分析中...")
    front_report = generate_front_report_json(client, articles_data, progress_callback)

    front_json_path = output_folder / "front_report.json"
    with open(front_json_path, "w", encoding="utf-8") as f:
        json.dump(front_report, f, ensure_ascii=False, indent=2)

    # Step 3: フロントレポート PDF
    if progress_callback:
        progress_callback("WeeklyReport: フロントレポートPDF 生成中...")
    front_pdf_path = output_folder / f"{topic_name}_フロントレポート.pdf"
    generate_front_report_pdf(front_report, front_pdf_path)

    # Step 4: 連結 (既存記事PDFがある場合)
    date_m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", collection_date)
    date_suffix = ""
    if date_m:
        date_suffix = f"{date_m.group(1)}{int(date_m.group(2)):02d}{int(date_m.group(3)):02d}"

    final_name = f"{topic_name}ウィークリーレポート{date_suffix}.pdf"
    final_path = output_folder / final_name

    if existing_articles_pdf and existing_articles_pdf.exists():
        if progress_callback:
            progress_callback("WeeklyReport: PDF連結中...")
        concatenate_pdfs(front_pdf_path, existing_articles_pdf, final_path)
    else:
        # 記事PDFがない場合はフロントレポート単体
        import shutil
        shutil.copy2(front_pdf_path, final_path)

    if progress_callback:
        progress_callback(f"WeeklyReport 完了: {final_path.name}")

    return final_path
