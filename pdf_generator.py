"""全記事を1つのPDFにまとめるモジュール（fpdf2使用）"""

import io
import json
import re
import tempfile
from pathlib import Path
from datetime import datetime
from typing import Optional

import requests
from fpdf import FPDF
from fpdf.fonts import FontFace

from article_parser import Article
from config import HTTP_HEADERS, HTTP_TIMEOUT

# =====================================================================
# カラー定数（HTMLテーマと統一）
# =====================================================================
DARK_BLUE = (26, 35, 126)         # #1a237e
MEDIUM_BLUE = (40, 53, 147)       # #283593
LIGHT_BLUE_BG = (232, 234, 246)   # #e8eaf6
WHITE = (255, 255, 255)
TEXT_DARK = (51, 51, 51)           # #333
TEXT_GRAY = (136, 136, 136)       # #888
SECTION_BORDER = DARK_BLUE
ACCENT_GREEN = (46, 125, 50)       # #2e7d32
ACCENT_ORANGE = (230, 126, 34)     # #e67e22
ACCENT_RED = (211, 47, 47)         # #d32f2f
HIGHLIGHT_BG = (255, 243, 224)     # #fff3e0
CATEGORY_COLORS = [
    (25, 118, 210),   # Blue
    (56, 142, 60),    # Green
    (245, 124, 0),    # Orange
    (156, 39, 176),   # Purple
    (0, 151, 167),    # Teal
]

# フォントパス
MEIRYO_PATH = "C:/Windows/Fonts/meiryo.ttc"
MEIRYO_BOLD_PATH = "C:/Windows/Fonts/meiryob.ttc"
GOTHIC_PATH = "C:/Windows/Fonts/msgothic.ttc"


# =====================================================================
# PDF クラス
# =====================================================================

class ArticlePDF(FPDF):
    """記事まとめPDF用のカスタムFPDFクラス"""

    def __init__(self, topic_name: str, collection_date: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.topic_name = topic_name
        self.collection_date = collection_date
        self._font_family_name = "Meiryo"

        # フォント登録
        self._register_fonts()

        # ページ設定
        self.set_auto_page_break(auto=True, margin=20)
        self.set_margins(15, 15, 15)

    def _register_fonts(self):
        """日本語フォントを登録（メイリオ優先、フォールバックはMSゴシック）"""
        font_path = Path(MEIRYO_PATH)
        bold_path = Path(MEIRYO_BOLD_PATH)

        if not font_path.exists():
            # メイリオがない場合はMSゴシックにフォールバック
            font_path = Path(GOTHIC_PATH)
            bold_path = Path(GOTHIC_PATH)
            self._font_family_name = "Gothic"

        if not font_path.exists():
            # どちらもない場合は組み込みフォント
            self._font_family_name = "Helvetica"
            return

        self.add_font(
            self._font_family_name, style="",
            fname=str(font_path), collection_font_number=0,
        )
        self.add_font(
            self._font_family_name, style="B",
            fname=str(bold_path), collection_font_number=0,
        )

    def _set_font(self, style: str = "", size: int = 10):
        """日本語フォントをセット"""
        self.set_font(self._font_family_name, style=style, size=size)

    def footer(self):
        """ページフッター"""
        self.set_y(-15)
        self._set_font(size=7)
        self.set_text_color(*TEXT_GRAY)
        self.cell(
            0, 10,
            f"{self.topic_name}  |  {self.collection_date}  |  "
            f"Search Grounding (Gemini API)  |  Page {self.page_no()}/{{nb}}",
            align="C",
        )


# =====================================================================
# 表紙
# =====================================================================

def _add_cover_page(pdf: ArticlePDF, topic_name: str,
                    collection_date: str, article_count: int):
    """表紙ページを追加"""
    pdf.add_page()

    # 上部のダークブルー背景（ページ上半分）
    pdf.set_fill_color(*DARK_BLUE)
    pdf.rect(0, 0, 210, 140, style="F")

    # タイトル（白文字）
    pdf.set_text_color(*WHITE)
    pdf._set_font(style="B", size=24)
    pdf.set_y(45)
    pdf.multi_cell(0, 14, topic_name, align="C")

    # サブタイトル
    pdf._set_font(size=14)
    pdf.ln(8)
    pdf.cell(0, 10, "調査レポート", align="C", new_x="LMARGIN", new_y="NEXT")

    # 日付
    pdf.ln(5)
    pdf._set_font(size=12)
    pdf.set_text_color(200, 200, 230)
    pdf.cell(0, 10, f"収集日: {collection_date}", align="C", new_x="LMARGIN", new_y="NEXT")

    # 記事件数
    pdf.ln(3)
    pdf.cell(0, 10, f"全 {article_count} 件", align="C", new_x="LMARGIN", new_y="NEXT")

    # 下部の情報
    pdf.set_y(200)
    pdf.set_text_color(*TEXT_GRAY)
    pdf._set_font(size=9)
    pdf.cell(
        0, 10,
        "自動記事収集・翻訳システム (Gemini API使用)",
        align="C",
    )


# =====================================================================
# サマリーページ（1ページ動向要約）
# =====================================================================

def _check_remaining(pdf, needed: float):
    """ページ下端までの残りスペースを確認し、足りなければ改ページする"""
    page_bottom = pdf.h - pdf.b_margin
    if pdf.get_y() + needed > page_bottom:
        pdf.add_page()


def _add_summary_page(pdf: ArticlePDF, summary_data: dict,
                      topic_name: str, collection_date: str):
    """エグゼクティブサマリーページを追加（表紙の直後、記事の前）"""
    pdf.add_page()
    pw = pdf.w - pdf.l_margin - pdf.r_margin  # 有効幅 180mm

    # --- ヘッダーバー ---
    pdf.set_fill_color(*DARK_BLUE)
    pdf.rect(0, 0, 210, 32, style="F")

    pdf.set_y(6)
    pdf.set_text_color(*WHITE)
    pdf._set_font(style="B", size=16)
    pdf.cell(0, 8, f"{topic_name}  Weekly Report", align="C",
             new_x="LMARGIN", new_y="NEXT")

    pdf._set_font(size=9)
    pdf.set_text_color(200, 200, 230)
    stats = summary_data.get("stats", {})
    total = stats.get("total_articles", "")
    countries = stats.get("countries", "")
    pdf.cell(0, 5, f"{collection_date}  |  {total}件  |  {countries}カ国",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_y(35)

    # --- 今週の概要 ---
    pdf.set_text_color(*DARK_BLUE)
    pdf._set_font(style="B", size=10)
    pdf.cell(0, 5, "■ 今週の動向", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)

    pdf.set_text_color(*TEXT_DARK)
    pdf._set_font(size=8.5)
    overview = summary_data.get("overview", "")
    pdf.multi_cell(pw, 4.2, overview, align="L")
    pdf.ln(2)

    # --- ハイライト（横並びボックス） ---
    highlights = summary_data.get("key_highlights", [])[:5]
    if highlights:
        row_h = 20  # 1行あたりの高さ（カラーバー3mm + 本体15mm + 余白2mm）
        total_rows = (len(highlights) + 2) // 3
        _check_remaining(pdf, 7 + row_h)  # 最低1行分は確保

        pdf.set_text_color(*DARK_BLUE)
        pdf._set_font(style="B", size=10)
        pdf.cell(0, 5, "■ 注目トピック", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)

        box_w = (pw - 4 * 2) / min(len(highlights), 3)  # 3列まで
        start_x = pdf.l_margin
        start_y = pdf.get_y()

        # 自動改ページを一時無効化（横並び描画中の分断を防止）
        pdf.set_auto_page_break(auto=False)

        for i, hl in enumerate(highlights):
            col = i % 3
            if i > 0 and col == 0:
                start_y = start_y + row_h

                # 新しい行がページに収まるか確認
                page_bottom = pdf.h - pdf.b_margin
                if start_y + row_h > page_bottom:
                    pdf.set_auto_page_break(auto=True, margin=20)
                    pdf.add_page()
                    pdf.set_auto_page_break(auto=False)
                    start_y = pdf.get_y()

            x = start_x + col * (box_w + 2)
            y = start_y

            # ボックス背景
            color = CATEGORY_COLORS[i % len(CATEGORY_COLORS)]
            pdf.set_fill_color(*color)
            pdf.rect(x, y, box_w, 3, style="F")  # カラーバー

            pdf.set_fill_color(245, 245, 250)
            pdf.rect(x, y + 3, box_w, 15, style="F")  # 本体

            # ラベル
            pdf.set_xy(x + 1, y + 4)
            pdf.set_text_color(*color)
            pdf._set_font(style="B", size=8)
            ref = hl.get("article_ref", "")
            pdf.cell(box_w - 2, 3.5, f'{hl.get("label", "")} {ref}', align="L")

            # 詳細
            pdf.set_xy(x + 1, y + 8)
            pdf.set_text_color(80, 80, 80)
            pdf._set_font(size=6.5)
            detail = hl.get("detail", "")[:50]
            pdf.multi_cell(box_w - 2, 3, detail, align="L")

        pdf.set_auto_page_break(auto=True, margin=20)
        pdf.set_y(start_y + row_h)
        pdf.ln(2)

    # --- カテゴリー別動向 ---
    categories = summary_data.get("categories", [])
    if categories:
        _check_remaining(pdf, 20)

        pdf.set_text_color(*DARK_BLUE)
        pdf._set_font(style="B", size=10)
        pdf.cell(0, 5, "■ カテゴリー別動向", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)

        for i, cat in enumerate(categories):
            # 各カテゴリーが収まるか確認（約14mm必要）
            _check_remaining(pdf, 14)

            color = CATEGORY_COLORS[i % len(CATEGORY_COLORS)]
            cat_y = pdf.get_y()

            # カテゴリーカラーバー
            pdf.set_fill_color(*color)
            pdf.rect(pdf.l_margin, cat_y, 2, 8, style="F")

            # カテゴリー名 + 件数
            pdf.set_x(pdf.l_margin + 4)
            pdf.set_text_color(*color)
            pdf._set_font(style="B", size=8.5)
            name = cat.get("name", "")
            count = cat.get("count", 0)
            refs = ", ".join(cat.get("article_refs", []))
            pdf.cell(50, 4, f"{name}（{count}件）", align="L")

            # 記事参照
            pdf.set_text_color(*TEXT_GRAY)
            pdf._set_font(size=7)
            pdf.cell(0, 4, refs, align="L", new_x="LMARGIN", new_y="NEXT")

            # サマリー
            pdf.set_x(pdf.l_margin + 4)
            pdf.set_text_color(80, 80, 80)
            pdf._set_font(size=7.5)
            summary_text = cat.get("summary", "")[:100]
            pdf.multi_cell(pw - 4, 3.5, summary_text, align="L")
            pdf.ln(0.5)

        pdf.ln(1)

    # --- ロードマップ（タイムライン） ---
    timeline = summary_data.get("timeline", [])
    if timeline:
        _check_remaining(pdf, 30)  # タイムラインは約25mm必要

        pdf.set_text_color(*DARK_BLUE)
        pdf._set_font(style="B", size=10)
        pdf.cell(0, 5, "■ 今後のロードマップ", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

        tl_y = pdf.get_y()
        tl_x_start = pdf.l_margin + 5
        tl_width = pw - 10
        step = tl_width / max(len(timeline), 1)

        # 自動改ページを一時無効化
        pdf.set_auto_page_break(auto=False)

        # 横線（タイムラインバー）
        pdf.set_draw_color(*DARK_BLUE)
        pdf.set_line_width(0.8)
        pdf.line(tl_x_start, tl_y + 4, tl_x_start + tl_width, tl_y + 4)

        for i, item in enumerate(timeline):
            cx = tl_x_start + i * step + step / 2
            color = CATEGORY_COLORS[i % len(CATEGORY_COLORS)]

            # ドット
            pdf.set_fill_color(*color)
            pdf.ellipse(cx - 2, tl_y + 2, 4, 4, style="F")

            # 期間ラベル
            pdf.set_xy(cx - step / 2, tl_y + 7)
            pdf.set_text_color(*color)
            pdf._set_font(style="B", size=7)
            pdf.cell(step, 3.5, item.get("period", ""), align="C")

            # イベント
            pdf.set_xy(cx - step / 2, tl_y + 10.5)
            pdf.set_text_color(80, 80, 80)
            pdf._set_font(size=6)
            event_text = item.get("event", "")[:30]
            pdf.multi_cell(step, 2.8, event_text, align="C")

        pdf.set_auto_page_break(auto=True, margin=20)
        pdf.set_y(tl_y + 22)
        pdf.ln(1)

    # --- 今後の展望 ---
    outlook = summary_data.get("outlook", "")
    if outlook:
        # dry_runで高さを先に計測
        pdf._set_font(size=8)
        result = pdf.multi_cell(pw - 6, 3.8, outlook, align="L",
                                dry_run=True, output="LINES")
        line_count = len(result) if result else 1
        text_h = line_count * 3.8 + 4

        _check_remaining(pdf, text_h + 8)  # タイトル分も含めて確認

        pdf.set_text_color(*DARK_BLUE)
        pdf._set_font(style="B", size=10)
        pdf.cell(0, 5, "■ 今後の展望", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)

        # 展望ボックス — 背景→テキストの順に描画
        pdf.set_fill_color(*LIGHT_BLUE_BG)
        pdf.set_text_color(*TEXT_DARK)
        pdf._set_font(size=8)

        start_y = pdf.get_y()

        # 背景を先に描画
        pdf.rect(pdf.l_margin, start_y - 1, pw, text_h, style="F")

        # テキストを描画
        pdf.set_y(start_y)
        pdf.set_x(pdf.l_margin + 3)
        pdf.multi_cell(pw - 6, 3.8, outlook, align="L")

    # --- キーメトリクス ---
    metrics = stats.get("key_metrics", [])
    if metrics:
        _check_remaining(pdf, 15)  # メトリクス行に必要な高さ

        pdf.ln(3)

        # 自動改ページを一時無効化（横並び描画中の分断を防止）
        pdf.set_auto_page_break(auto=False)

        metric_w = pw / len(metrics)
        metric_y = pdf.get_y()
        for i, m in enumerate(metrics):
            x = pdf.l_margin + i * metric_w
            pdf.set_xy(x, metric_y)
            pdf.set_text_color(*DARK_BLUE)
            pdf._set_font(style="B", size=9)
            pdf.cell(metric_w, 4, m.get("value", ""), align="C",
                     new_x="LEFT", new_y="NEXT")
            pdf.set_x(x)
            pdf.set_text_color(*TEXT_GRAY)
            pdf._set_font(size=6.5)
            pdf.cell(metric_w, 3, m.get("label", ""), align="C")

        pdf.set_auto_page_break(auto=True, margin=20)


# =====================================================================
# 画像ダウンロード
# =====================================================================

def _download_image(url: str) -> Optional[str]:
    """画像をダウンロードして一時ファイルパスを返す"""
    if not url:
        return None
    try:
        resp = requests.get(url, headers=HTTP_HEADERS, timeout=5, stream=True)
        resp.raise_for_status()

        content_type = resp.headers.get("Content-Type", "")
        if "image" not in content_type and not url.lower().endswith(
            (".jpg", ".jpeg", ".png", ".gif", ".webp")
        ):
            return None

        # 拡張子判定
        if "png" in content_type or url.lower().endswith(".png"):
            ext = ".png"
        elif "gif" in content_type or url.lower().endswith(".gif"):
            ext = ".gif"
        else:
            ext = ".jpg"

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        for chunk in resp.iter_content(8192):
            tmp.write(chunk)
        tmp.close()
        return tmp.name
    except Exception:
        return None


# =====================================================================
# HTMLコンテンツの前処理
# =====================================================================

def _strip_img_tags(html: str) -> str:
    """HTMLから<img>タグを除去"""
    if not html:
        return ""
    return re.sub(r"<img[^>]*/?>", "", html, flags=re.IGNORECASE)


def _clean_html_for_pdf(html: str) -> str:
    """write_html()用にHTMLをクリーニング"""
    if not html:
        return "<p></p>"

    html = _strip_img_tags(html)

    # 空のタグを除去
    html = re.sub(r"<(\w+)>\s*</\1>", "", html)

    # write_html()がサポートしないタグを変換
    html = re.sub(r"</?div[^>]*>", "", html, flags=re.IGNORECASE)
    html = re.sub(r"</?span[^>]*>", "", html, flags=re.IGNORECASE)

    # pタグで囲まれていない文章はpで囲む
    if not html.strip().startswith("<"):
        html = f"<p>{html}</p>"

    return html


# =====================================================================
# 記事ページ
# =====================================================================

def _add_article(pdf: ArticlePDF, article: Article,
                 article_number: int, collection_date: str):
    """1記事分のページを追加"""
    pdf.add_page()
    page_width = pdf.w - pdf.l_margin - pdf.r_margin  # 有効幅

    # --- ヘッダーバー（ダークブルー） ---
    header_y = pdf.get_y()
    pdf.set_fill_color(*DARK_BLUE)
    pdf.rect(0, header_y - 5, 210, 38, style="F")

    # 記事番号バッジ
    pdf.set_y(header_y)
    pdf.set_text_color(*WHITE)
    pdf._set_font(style="B", size=11)
    pdf.cell(12, 8, f"#{article_number:02d}", align="C")

    # タイトル
    pdf._set_font(style="B", size=13)
    title_x = pdf.get_x()
    title_w = page_width - 12
    pdf.multi_cell(title_w, 7, article.title_ja or "無題", align="L")

    # メタ情報行
    meta_y = pdf.get_y()
    if meta_y < header_y + 22:
        meta_y = header_y + 22
    pdf.set_y(meta_y)
    pdf.set_x(pdf.l_margin + 12)
    pdf._set_font(size=8)
    pdf.set_text_color(200, 200, 230)

    # 公開日を整形
    pub_date = article.publish_date or "日付不明"
    if pub_date and re.match(r"\d{4}-\d{2}-\d{2}", pub_date):
        try:
            dt = datetime.strptime(pub_date, "%Y-%m-%d")
            pub_date = dt.strftime("%Y年%m月%d日")
        except ValueError:
            pass

    meta_text = f"公開日 {pub_date}  |  {article.source_name or '出典不明'}  |  {article.country or ''}"
    pdf.cell(title_w, 5, meta_text, align="L")

    # ヘッダー下に余白
    pdf.set_y(header_y + 40)
    pdf.set_text_color(*TEXT_DARK)

    # --- 画像 ---
    if article.image_url:
        img_path = _download_image(article.image_url)
        if img_path:
            try:
                img_y = pdf.get_y()
                # 幅いっぱいに表示、高さは自動
                pdf.image(img_path, x=pdf.l_margin, w=page_width)
                pdf.ln(5)
            except Exception:
                pass
            finally:
                # 一時ファイルを削除
                try:
                    Path(img_path).unlink(missing_ok=True)
                except Exception:
                    pass

    # --- 概要セクション ---
    pdf.ln(3)
    _draw_section_title(pdf, "概要")
    pdf.ln(2)

    # 薄青背景ボックス
    box_x = pdf.l_margin
    pdf.set_fill_color(*LIGHT_BLUE_BG)

    # 概要テキスト
    pdf._set_font(size=10)
    pdf.set_text_color(*TEXT_DARK)

    summary = article.summary_ja or ""

    # dry_runで高さを計測
    start_y = pdf.get_y()
    result = pdf.multi_cell(page_width - 6, 6, summary, align="L",
                            dry_run=True, output="LINES")
    line_count = len(result) if result else 1
    box_height = line_count * 6 + 4

    # 背景を先に描画
    pdf.set_fill_color(*LIGHT_BLUE_BG)
    pdf.rect(box_x, start_y - 2, page_width, box_height, style="F")

    # テキストを描画
    pdf.set_y(start_y)
    pdf.set_x(box_x + 3)
    pdf._set_font(size=10)
    pdf.set_text_color(*TEXT_DARK)
    pdf.multi_cell(page_width - 6, 6, summary, align="L")

    pdf.ln(5)

    # --- 詳細セクション ---
    _draw_section_title(pdf, "詳細")
    pdf.ln(2)

    detail_html = _clean_html_for_pdf(article.content_ja)
    pdf._set_font(size=10)
    pdf.set_text_color(*TEXT_DARK)

    try:
        pdf.write_html(
            detail_html,
            tag_styles={
                "h3": FontFace(
                    family=pdf._font_family_name,
                    size_pt=12,
                    color=MEDIUM_BLUE,
                ),
            },
        )
    except Exception:
        # write_htmlが失敗した場合はプレーンテキストで表示
        plain = re.sub(r"<[^>]+>", "", detail_html)
        pdf.multi_cell(0, 6, plain, align="L")

    # --- ソースURL ---
    if article.url:
        pdf.ln(5)
        pdf.set_text_color(*TEXT_GRAY)
        pdf._set_font(size=8)
        pdf.cell(0, 5, f"元記事: {article.url}", align="R")

    pdf.ln(3)


def _draw_section_title(pdf: ArticlePDF, title: str):
    """セクションタイトルを左ボーダー付きで描画"""
    y = pdf.get_y()
    # 左ボーダー（4px相当）
    pdf.set_fill_color(*DARK_BLUE)
    pdf.rect(pdf.l_margin, y, 1.5, 7, style="F")

    # テキスト
    pdf.set_x(pdf.l_margin + 4)
    pdf._set_font(style="B", size=11)
    pdf.set_text_color(*DARK_BLUE)
    pdf.cell(0, 7, title, align="L")
    pdf.ln(7)

    # テキスト色を戻す
    pdf.set_text_color(*TEXT_DARK)


# =====================================================================
# メイン公開関数
# =====================================================================

def generate_combined_pdf(
    articles: list[Article],
    topic_name: str,
    output_folder: Path,
    collection_date: str | None = None,
    summary_data: dict | None = None,
) -> Path:
    """全記事を1つのPDFにまとめて保存する

    Args:
        articles: 記事リスト
        topic_name: 調査テーマ名（ファイル名に使用）
        output_folder: 出力フォルダ
        collection_date: 収集日表示テキスト（例: "2026年03月02日"）
        summary_data: サマリーページ用データ（Gemini生成JSON）。Noneならサマリーなし。

    Returns:
        保存されたPDFファイルのPath
    """
    if not collection_date:
        collection_date = datetime.now().strftime("%Y年%m月%d日")

    date_str = datetime.now().strftime("%Y%m%d")

    # PDF作成
    pdf = ArticlePDF(topic_name, collection_date)
    pdf.alias_nb_pages()

    # 表紙
    _add_cover_page(pdf, topic_name, collection_date, len(articles))

    # サマリーページ（表紙の直後）
    if summary_data:
        _add_summary_page(pdf, summary_data, topic_name, collection_date)

    # 各記事
    for idx, article in enumerate(articles):
        _add_article(pdf, article, idx + 1, collection_date)

    # ファイル名: {調査テーマ名}ウィークリーレポート{YYYYMMDD}.pdf
    filename = f"{topic_name}ウィークリーレポート{date_str}.pdf"
    filepath = output_folder / filename

    pdf.output(str(filepath))
    return filepath
