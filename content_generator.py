"""コンテンツジェネレーター - WeeklyReport(PDF)とポッドキャスト原稿を生成

採用済み記事のみを対象に:
1. Gemini APIで分析テキストを生成 → reportlabでPDF化
2. ポッドキャスト原稿テキストを生成
"""
import sys
import os
import json
import re
import threading
from pathlib import Path
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Text, Scrollbar,
    StringVar, messagebox, filedialog,
    BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, N, S,
    WORD, END, DISABLED, NORMAL,
)
from tkinter import ttk

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ─── パス設定 ───
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "調査内容ファイル"
SURVEY_OUTPUT_DIR = BASE_DIR / "調査アウトプット"

try:
    from dotenv import load_dotenv
    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# ─── 依存インポート ───
try:
    from bs4 import BeautifulSoup
    BS4_OK = True
except ImportError:
    BS4_OK = False

try:
    from google import genai
    from google.genai import types as genai_types
    GENAI_OK = True
except ImportError:
    GENAI_OK = False

# lib/weekly_report_generator を優先使用
sys.path.insert(0, str(BASE_DIR / "lib"))
try:
    from weekly_report_generator import (
        generate_front_report_json,
        generate_front_report_pdf,
        concatenate_pdfs,
    )
    RICH_REPORT_OK = True
except Exception:
    RICH_REPORT_OK = False

# podcast_reviewer からVoicePeak音声生成関数をインポート
try:
    from podcast_reviewer import (
        parse_dialogue_script,
        build_segments,
        generate_wav,
        combine_wavs_to_mp3,
        check_missing_segments,
        _kill_stale_voicepeak,
        NARRATOR_F,
        NARRATOR_M,
    )
    PODCAST_AUDIO_OK = True
except Exception:
    PODCAST_AUDIO_OK = False

# ─── パイプライン状態管理（article_curatorと共通） ───

def load_pipeline_state(folder: Path) -> dict:
    state_file = folder / "_pipeline_state.json"
    if state_file.exists():
        try:
            return json.loads(state_file.read_text(encoding="utf-8"))
        except Exception:
            pass
    return _default_state(folder.name)


def save_pipeline_state(folder: Path, state: dict):
    state_file = folder / "_pipeline_state.json"
    state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _default_state(folder_name: str) -> dict:
    return {
        "folder": folder_name,
        "created": datetime.now().isoformat(timespec="seconds"),
        "stages": {
            "collection": {"status": "pending", "article_count": 0, "completed_at": None},
            "curation": {
                "status": "pending", "articles": {}, "adopted_count": 0,
                "rejected_count": 0, "completed_at": None,
            },
            "generation": {"status": "pending", "pdf_path": None, "script_path": None, "completed_at": None},
            "podcast_review": {"status": "unreviewed", "review_count": 0, "last_position_sec": 0.0, "completed_at": None},
            "upload": {"status": "pending", "uploaded_count": 0, "completed_at": None},
        },
    }


# ─── HTML記事パース ───

def parse_html_article(filepath: Path) -> dict:
    """採用記事HTMLからタイトル・本文を抽出"""
    html = filepath.read_text(encoding="utf-8")
    if BS4_OK:
        soup = BeautifulSoup(html, "html.parser")
        title_tag = soup.select_one(".header h1") or soup.find("h1") or soup.find("title")
        title = title_tag.get_text(strip=True) if title_tag else filepath.stem
        for tag in soup(["script", "style"]):
            tag.decompose()
        body = soup.get_text(separator="\n", strip=True)
    else:
        title = filepath.stem
        body = re.sub(r"<[^>]+>", "", html)

    return {"title": title, "body": body[:4000], "filepath": filepath}


def get_adopted_articles(folder: Path, state: dict) -> list[dict]:
    """採用記事（不採用_から始まらないNN_*.html）を返す"""
    curation = state.get("stages", {}).get("curation", {})

    # まずファイルシステムから採用ファイルを収集
    all_files = sorted(
        f for f in folder.glob("*.html")
        if re.match(r"^\d+", f.name) and not f.name.startswith("不採用_")
    )
    articles = []
    for f in all_files:
        art = parse_html_article(f)
        articles.append(art)
    return articles


# ─── プロンプトファイル読み込み ───

DEFAULT_REPORT_PROMPT = """以下の記事群を分析し、ウィークリーレポートを作成してください。

## 分析の切り口

1. 今週の最重要ニュース（1〜2件に絞って重要度を説明する）
2. 技術トレンドの方向性（材料・製造プロセス・コスト動向）
3. 注目すべき企業・研究機関の動向
4. 市場・量産化スケジュールへの影響
5. 来週以降の注目ポイントと今後の展望

## 出力形式
- 日本語で記述
- 各セクションに見出し（##）を付ける
- 重要な数値・固有名詞は正確に記載
- 全体で1500〜2500字程度
"""

DEFAULT_PODCAST_PROMPT = """以下の採用記事をもとに、技術情報ポッドキャスト番組の原稿を作成してください。

## 目指すトーン
NHKラジオの情報番組のような、落ち着いた知的な対話。
テレビのバラエティ的な大げさなリアクションは一切不要。
スージーの関心・驚きは「言葉の内容」で自然に表現する。
ただし、音声合成で抑揚を付けるため、感情タグはバランスよく織り交ぜること。

## 登場人物（台詞中では必ず名前で呼ぶこと。記号 F / M は書かない）
- F = スージー: 女性ホスト。技術の専門家ではないが、好奇心があり的確な質問をする。
  普段は落ち着いたトーン。本当に良いニュースの時だけ少し明るくなる。
- M = トロイ: 男性専門家。論理的で丁寧。難しい技術をわかりやすく解説する。

## 【重要】話者記号 F / M について
F と M は**行頭にだけ書く内部マーカー**で、台詞テキストの中には絶対に入れないこと。
台詞中では必ず「スージー」「トロイ」と名前で呼ぶ。

NG例: 「ホストのF、スージーです」「アナリストのM、トロイさんです」
OK例: 「ホストのスージーです」「アナリストのトロイです」

## 話者記号と感情コード（必ず行頭にこの形式で記述）
- F[N]: スージーの通常発言（相槌・コメント・つなぎ）
- F[Q]: スージーの疑問・質問（語尾上がり）
- F[S]: スージーの落ち着いたまとめ・重要な指摘
- F[H]: スージーの明るい発言（良いニュース・ポジティブ反応）
- F[T]: スージーの軽いツッコミ・冗談
- F[E]: スージーの驚き（本当に驚くべき時のみ）
- M[N]: トロイの通常説明
- M[S]: トロイの重要な技術的説明
- M[H]: トロイの温かみのある明るい発言
- M[Q]: トロイの問いかけ

## 感情コードの使用比率（程よいメリハリを必ず付けること）
F[N] 30〜40% / F[Q] 20〜30% / F[S] 10〜15% / F[H] 10〜15% / F[T] 5〜10% / F[E] 2〜5%
M[N] 40〜55% / M[S] 20〜30% / M[H] 5〜10% / M[Q] 5%程度

※ [N] ばかりにすると音声が平坦になるため、必ず他の感情タグも織り交ぜる。
※ 1記事（8〜12行）の中で最低3種類以上の感情コードを使うこと。

## 原稿の構成
1. オープニング（F[N]またはF[H]で挨拶、今週のテーマ紹介）
2. 各記事のトピック解説（記事1件につき8〜12行のやりとり）
   - F[Q]でスージーが質問 → M[N]/M[S]でトロイが解説 → F[N]/F[S]でスージーが受ける
3. クロージング（今週のまとめ、来週への展望）

## 分量
- 記事1件につき会話8〜12行（約400〜600字）
- 全体で原稿本文2000〜4000字程度

## 出力形式（厳守）
各行は必ず「話者コード: テキスト」の形式。
空行でトピックを区切ること。
HTMLタグ・markdown・余分な説明文は一切含めないこと。

## 良い例と悪い例

【悪い例 NG — 台詞中に F / M が混入】
F[N]: ホストのF、スージーです。             ← F が台詞内に混入
F[N]: アナリストのM、トロイさんです。        ← M が台詞内に混入
F[N]: ナビゲーターのF、スージーです。        ← F が台詞内に混入

【悪い例 NG — 大げさすぎるリアクション】
F[E]: すごーい！それって本当にすごいことじゃないですか！
F[H]: うわー！なんかもう、毎日が驚きの連続ですね！
F[Q]: ちょ、ちょっと待って！それってどういうことですか？

【良い例 OK — 自己紹介】
F[H]: 皆さん、こんにちは！ホストのスージーです。今週も最新のAI動向をお届けします。
M[N]: 皆さん、こんにちは。アナリストのトロイです。今週もよろしくお願いします。

【良い例 OK — 通常の対話】
F[N]: なるほど。そうなると、製造コストにも大きく影響してくるんですね。
F[Q]: 具体的にはどのくらいのコスト削減が見込めるんですか？
F[S]: 今週の話をまとめると、量産化に向けた具体的な動きが各社で出てきた、ということですね。
F[H]: 患者さんにとっては、本当に心強い話ですね。
F[T]: トロイさん、それ先週も同じこと言ってませんでした？

## 注意事項
- 専門用語は必ずスージーの質問→トロイの解説という形でフォローする
- 数値・企業名・技術名は正確に記載する
- 感嘆符「！」は1発言に1個まで。連発しない
- 「すごーい」「うわー」「えーっ」「ちょ待って」などのバラエティ的表現は禁止
"""

DEFAULT_PODCAST_PROMPT_DETAIL = """
以下の採用記事の内容を参考にポッドキャスト原稿を作成してください。

{articles_text}

上記の記事を読みやすい対話形式にまとめ、リスナーが楽しめる内容にしてください。
"""


def load_report_prompt(topic_name: str) -> str:
    prompt_file = BASE_DIR / "調査内容ファイル" / f"{topic_name}_report_prompt.txt"
    if prompt_file.exists():
        return prompt_file.read_text(encoding="utf-8")
    return DEFAULT_REPORT_PROMPT


def load_podcast_prompt(topic_name: str) -> str:
    """トピック固有プロンプト + 共通ルールを結合して返す。

    読み込み優先順:
      1. {topic}_podcast_prompt.txt  （トピック固有・存在すれば使用）
      2. DEFAULT_PODCAST_PROMPT      （ファイルなし時のフォールバック）
    その後、共通_podcast_rules.txt を末尾に追加する（存在する場合）。
    """
    prompt_file = BASE_DIR / "調査内容ファイル" / f"{topic_name}_podcast_prompt.txt"
    base_prompt = (
        prompt_file.read_text(encoding="utf-8")
        if prompt_file.exists()
        else DEFAULT_PODCAST_PROMPT
    )

    # 共通ルールを末尾に追加
    common_rules_file = BASE_DIR / "調査内容ファイル" / "共通_podcast_rules.txt"
    if common_rules_file.exists():
        common_rules = common_rules_file.read_text(encoding="utf-8")
        return base_prompt.rstrip() + "\n\n" + common_rules
    return base_prompt


# ─── Gemini API呼び出し ───

def call_gemini(prompt: str, api_key: str, model: str = "gemini-2.5-flash") -> str:
    if not GENAI_OK:
        raise ImportError("google-genai ライブラリがインストールされていません")
    # gemini-1.5-flash は v1beta API で廃止済み（404 NOT_FOUND）のため除外
    _MODELS = [model, "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.0-flash-lite"]
    seen: set = set()
    last_err = None

    def _is_transient(exc):
        s = str(exc)
        return any(k in s for k in ("429", "503", "RESOURCE_EXHAUSTED",
                                    "overloaded", "rate", "quota", "too many"))

    def _is_not_found(exc):
        s = str(exc)
        return any(k in s for k in ("404", "NOT_FOUND", "not found", "not supported"))

    for m in _MODELS:
        if m in seen:
            continue
        seen.add(m)
        for attempt in range(3):
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(model=m, contents=prompt)
                return response.text
            except Exception as e:
                last_err = e
                if _is_not_found(e):
                    break  # 次モデルへ
                if _is_transient(e) and attempt < 2:
                    import time as _t
                    _t.sleep(10 * (2 ** attempt))  # 10s → 20s
                    continue
                break  # 次モデルへ
    raise RuntimeError(f"Gemini API呼び出し失敗（全モデル試行済み）: {last_err}")


# ─── PDF生成 ───

def generate_pdf(report_text: str, output_path: Path, title: str):
    """reportlab または fpdf2 でPDFを生成"""
    # reportlab を試す
    try:
        _generate_pdf_reportlab(report_text, output_path, title)
        return
    except ImportError:
        pass

    # fpdf2 を試す
    try:
        _generate_pdf_fpdf2(report_text, output_path, title)
        return
    except ImportError:
        pass

    # どちらもなければテキストファイルとして保存
    txt_path = output_path.with_suffix(".txt")
    txt_path.write_text(f"{title}\n\n{report_text}", encoding="utf-8")
    raise Exception(f"reportlab/fpdf2 が未インストールのため、テキストファイルとして保存しました: {txt_path.name}")


def _generate_pdf_reportlab(text: str, output_path: Path, title: str):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.enums import TA_LEFT
    import reportlab.pdfbase.pdfmetrics as metrics
    from reportlab.pdfbase.ttfonts import TTFont

    # 日本語フォント登録を試みる
    font_name = "Helvetica"
    jp_font_paths = [
        r"C:\Windows\Fonts\YuGothM.ttc",
        r"C:\Windows\Fonts\meiryo.ttc",
        r"C:\Windows\Fonts\msgothic.ttc",
    ]
    for fp in jp_font_paths:
        if Path(fp).exists():
            try:
                metrics.registerFont(TTFont("JpFont", fp))
                font_name = "JpFont"
                break
            except Exception:
                continue

    doc = SimpleDocTemplate(str(output_path), pagesize=A4,
                             rightMargin=20*mm, leftMargin=20*mm,
                             topMargin=25*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", fontName=font_name, fontSize=16,
                                  spaceAfter=12, leading=20)
    body_style = ParagraphStyle("Body", fontName=font_name, fontSize=10,
                                 leading=16, spaceAfter=6)
    h2_style = ParagraphStyle("H2", fontName=font_name, fontSize=13,
                               spaceBefore=12, spaceAfter=6, leading=18)

    story = [Paragraph(title, title_style), Spacer(1, 6*mm)]

    for line in text.split("\n"):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 3*mm))
            continue
        if line.startswith("## "):
            story.append(Paragraph(line[3:], h2_style))
        elif line.startswith("# "):
            story.append(Paragraph(line[2:], title_style))
        else:
            # HTMLタグをエスケープ
            line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(line, body_style))

    doc.build(story)


def _generate_pdf_fpdf2(text: str, output_path: Path, title: str):
    from fpdf import FPDF

    pdf = FPDF()
    pdf.add_page()

    # 日本語フォント登録
    jp_font_paths = [
        r"C:\Windows\Fonts\YuGothM.ttc",
        r"C:\Windows\Fonts\meiryo.ttc",
        r"C:\Windows\Fonts\msgothic.ttc",
    ]
    font_added = False
    for fp in jp_font_paths:
        if Path(fp).exists():
            try:
                pdf.add_font("JpFont", "", fp, uni=True)
                font_added = True
                break
            except Exception:
                continue

    font_name = "JpFont" if font_added else "Helvetica"

    pdf.set_font(font_name, size=16)
    pdf.cell(0, 12, title, ln=True)
    pdf.ln(4)
    pdf.set_font(font_name, size=10)

    for line in text.split("\n"):
        line = line.strip()
        if not line:
            pdf.ln(4)
            continue
        if line.startswith("## "):
            pdf.set_font(font_name, size=13)
            pdf.multi_cell(0, 8, line[3:])
            pdf.set_font(font_name, size=10)
        else:
            pdf.multi_cell(0, 6, line)

    pdf.output(str(output_path))


# ─── 採用記事HTML → Chromium(Playwright)で画像付きPDF化 ───

# 改ページ制御用CSS (文章の途中で切れないようにする)
_ARTICLE_PDF_CSS = """
@page { size: A4; margin: 18mm 14mm 18mm 14mm; }
html, body { background: #ffffff !important; }
body { padding: 0 !important; margin: 0 !important; }
.container {
  box-shadow: none !important;
  border-radius: 0 !important;
  margin: 0 !important;
  max-width: 100% !important;
}
/* 見出し直後での改ページを避ける */
h1, h2, h3, h4, .section-title {
  page-break-after: avoid !important;
  break-after: avoid !important;
}
/* 見出しとその直後のブロックをできるだけ一体で扱う */
.header { page-break-after: avoid !important; break-after: avoid !important; }
.section { page-break-inside: avoid !important; break-inside: avoid !important; }
.overview { page-break-inside: avoid !important; break-inside: avoid !important; }
p {
  page-break-inside: avoid !important;
  break-inside: avoid !important;
  orphans: 3;
  widows: 3;
}
li { page-break-inside: avoid !important; break-inside: avoid !important; }
.article-image { page-break-inside: avoid !important; break-inside: avoid !important; }
img {
  max-width: 100% !important;
  height: auto !important;
  page-break-inside: avoid !important;
  break-inside: avoid !important;
}
.footer, .source { page-break-before: avoid !important; break-before: avoid !important; }
/* 各記事末尾で強制改ページしたい場合はこのクラスを付与 */
.force-page-break { page-break-after: always !important; break-after: page !important; }
"""


def _render_html_to_pdf_with_browser(html_path: Path, out_pdf: Path,
                                     browser,
                                     extra_css: str = _ARTICLE_PDF_CSS,
                                     timeout_ms: int = 30000) -> bool:
    """既存のPlaywrightブラウザインスタンスを使って単一HTMLをPDF化する。

    ブラウザの起動/終了コストを省くため、呼び出し元でブラウザを管理する。
    失敗した場合は1回だけ再試行する。
    """
    for attempt in range(2):
        context = None
        try:
            context = browser.new_context()
            page = context.new_page()
            file_url = html_path.resolve().as_uri()
            page.goto(file_url, wait_until="load", timeout=timeout_ms)
            # 画像などの遅延リソースの読込を少し待つ（ベストエフォート）
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass
            if extra_css:
                page.add_style_tag(content=extra_css)
            page.pdf(
                path=str(out_pdf),
                format="A4",
                print_background=True,
                prefer_css_page_size=True,
                margin={"top": "18mm", "bottom": "18mm",
                        "left": "14mm", "right": "14mm"},
            )
            context.close()
            if out_pdf.exists() and out_pdf.stat().st_size > 0:
                return True
        except Exception as e:
            if context:
                try:
                    context.close()
                except Exception:
                    pass
            if attempt == 0:
                print(f"[chromium render retry] {html_path.name}: {e}")
                time.sleep(1)
            else:
                print(f"[chromium render failed] {html_path.name}: {e}")
    return False


def _render_html_to_pdf_chromium(html_path: Path, out_pdf: Path,
                                 extra_css: str = _ARTICLE_PDF_CSS,
                                 timeout_ms: int = 30000) -> bool:
    """Playwright(Chromium)で単一HTMLをPDF化（ブラウザを都度起動するスタンドアロン版）。

    単体テストや単発呼び出し用。バッチ処理には build_articles_pdf_via_chromium を使うこと。
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            ok = _render_html_to_pdf_with_browser(
                html_path, out_pdf, browser, extra_css=extra_css, timeout_ms=timeout_ms)
            browser.close()
        return ok
    except Exception as e:
        print(f"[chromium render failed] {html_path.name}: {e}")
        return False


def _make_cover_and_toc_html(articles: list[dict], topic_name: str) -> str:
    """表紙＋目次のHTMLを文字列で返す（Chromiumに直接渡すため一時ファイルを使わない）。"""
    today_str = datetime.now().strftime("%Y-%m-%d")
    rows = "".join(
        f"<li>{i:02d}. {(a.get('title') or '').replace('<','&lt;')}</li>"
        for i, a in enumerate(articles, 1)
    )
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"/>
<style>
@page {{ size: A4; margin: 22mm 18mm; }}
body {{ font-family: 'Meiryo','Yu Gothic UI','Hiragino Kaku Gothic ProN',sans-serif;
       color:#222; line-height:1.8; }}
.cover {{ text-align:center; padding:40mm 0 20mm; }}
.cover h1 {{ font-size: 26pt; color:#0D47A1; margin:0 0 12mm; line-height:1.3; }}
.cover p {{ font-size: 12pt; color:#555; margin:2mm 0; }}
h2 {{ color:#1A237E; border-left:6px solid #1A237E; padding-left:10px;
      margin: 0 0 10mm 0; page-break-before: always; }}
ol, ul {{ padding-left: 24px; }}
li {{ font-size: 10.5pt; margin: 3px 0; page-break-inside: avoid; }}
</style></head><body>
<div class="cover">
  <h1>{topic_name} 採用記事全文集</h1>
  <p>出力日: {today_str}</p>
  <p>採用記事数: {len(articles)} 件</p>
</div>
<h2>収録記事一覧</h2>
<ol>{rows}</ol>
</body></html>"""


def build_articles_pdf_via_chromium(articles: list[dict], output_path: Path,
                                    topic_name: str = "",
                                    progress_callback=None) -> Path | None:
    """採用記事HTMLをChromiumレンダリングで画像込みPDF化し、全てを1ファイルに連結。

    - ブラウザを1回だけ起動して全記事を処理する（起動コスト削減・安定性向上）
    - 各HTMLは元のデザイン/画像を保ったまま PDF 化される
    - 失敗した記事は1回リトライする
    - 先頭に表紙 + 目次を付ける
    """
    import sys, traceback
    # Windows の daemon スレッド内では asyncio の current event loop が未設定で
    # `get_event_loop()` が RuntimeError を投げるため、Playwright のサブプロセス
    # 起動が失敗する。ここで明示的に ProactorEventLoop をセットして対処する。
    # （過去事故: 修正後も GUI プロセスに古いモジュールがキャッシュされて
    #   このセットアップが走らず、無音でテキストフォールバックに落ちた）
    if sys.platform == "win32":
        import asyncio
        try:
            # 既存ポリシーが ProactorEventLoopPolicy でなければ設定
            policy = asyncio.get_event_loop_policy()
            if not isinstance(policy, asyncio.WindowsProactorEventLoopPolicy):
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        except Exception as e:
            # セットアップ自体が失敗した場合も検出できるようログに出す
            msg = f"event loop setup失敗: {type(e).__name__}: {e}"
            print(f"[build_articles_pdf_via_chromium] {msg}")
            if progress_callback:
                progress_callback(msg)

    try:
        from playwright.sync_api import sync_playwright
        from pypdf import PdfWriter, PdfReader
    except ImportError as e:
        # Silent fallback を避け、呼び出し側の GUI ログに確実に表示する
        msg = f"Chromium用ライブラリのインポートに失敗（HTMLレンダリング不可）: {e}"
        print(f"[build_articles_pdf_via_chromium] {msg}")
        if progress_callback:
            progress_callback(msg)
        return None

    if not articles:
        return None

    work_dir = output_path.parent / "_articles_pdf_work"
    work_dir.mkdir(exist_ok=True)

    # ── Chromiumブラウザ本体の存在確認・自動インストール ──
    # Playwright Pythonパッケージだけインストールされて
    # ブラウザ本体が未ダウンロードの場合、launch()が失敗する。
    # "playwright install chromium" を自動実行して解決する。
    def _ensure_chromium(pcb=None):
        import subprocess, sys
        try:
            from playwright.sync_api import sync_playwright as _spw
            with _spw() as _pw:
                _pw.chromium.launch().close()
            return True  # 起動できた → インストール済み
        except Exception as _e:
            err_s = str(_e).lower()
            needs_install = any(k in err_s for k in
                                ("executable doesn't exist", "not found",
                                 "playwright install", "browser is not installed"))
            if needs_install:
                msg = "Chromiumブラウザが未インストールです。自動インストール中... (数分かかります)"
                print(f"[build_articles_pdf_via_chromium] {msg}")
                if pcb:
                    pcb(msg)
                try:
                    result = subprocess.run(
                        [sys.executable, "-m", "playwright", "install", "chromium"],
                        capture_output=True, text=True, timeout=300)
                    if result.returncode == 0:
                        ok_msg = "Chromiumインストール完了。レンダリングを続行します。"
                        print(f"[build_articles_pdf_via_chromium] {ok_msg}")
                        if pcb:
                            pcb(ok_msg)
                        return True
                    else:
                        fail_msg = (f"playwright install 失敗 (code={result.returncode}):\n"
                                    f"{result.stderr[:300]}")
                        print(f"[build_articles_pdf_via_chromium] {fail_msg}")
                        if pcb:
                            pcb(fail_msg)
                        return False
                except Exception as install_err:
                    err_msg = f"playwright install 実行エラー: {install_err}"
                    print(f"[build_articles_pdf_via_chromium] {err_msg}")
                    if pcb:
                        pcb(err_msg)
                    return False
            # インストール済みだが別原因のエラー → そのまま続行（launch時に再度捕捉）
            return True

    _ensure_chromium(progress_callback)

    per_article_pdfs: list[Path] = []

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()

            # 表紙＋目次PDF（HTMLを一時ファイル経由でレンダリング）
            cover_pdf = work_dir / "00_cover.pdf"
            if progress_callback:
                progress_callback("表紙・目次PDFを生成中...")
            cover_html_path = work_dir / "00_cover.html"
            try:
                cover_html_path.write_text(
                    _make_cover_and_toc_html(articles, topic_name), encoding="utf-8")
                if _render_html_to_pdf_with_browser(
                        cover_html_path, cover_pdf, browser, extra_css=""):
                    per_article_pdfs.append(cover_pdf)
            except Exception as e:
                print(f"[cover pdf failed] {e}")
            finally:
                try:
                    cover_html_path.unlink(missing_ok=True)
                except Exception:
                    pass

            # 各記事PDF
            for i, art in enumerate(articles, 1):
                html_path = art.get("filepath")
                if not (isinstance(html_path, Path) and html_path.exists()):
                    if progress_callback:
                        progress_callback(f"記事 {i}/{len(articles)} スキップ（HTMLファイルなし）")
                    continue
                out_pdf = work_dir / f"{i:02d}_article.pdf"
                if progress_callback:
                    progress_callback(
                        f"記事 {i}/{len(articles)} をPDFレンダリング中: {html_path.name[:40]}")
                ok = _render_html_to_pdf_with_browser(html_path, out_pdf, browser)
                if ok:
                    per_article_pdfs.append(out_pdf)
                else:
                    if progress_callback:
                        progress_callback(f"  ⚠ 記事 {i} のレンダリングに失敗（スキップ）")

            browser.close()

    except Exception as e:
        # 例外の種類とトレースバックを完全に出して、GUIログで silent fallback の原因が
        # 追跡できるようにする（過去事故: 原因不明のテキストフォールバックで再発）
        tb = traceback.format_exc()
        msg = f"Chromium起動エラー: {type(e).__name__}: {e}"
        print(f"[build_articles_pdf_via_chromium] {msg}\n{tb}")
        if progress_callback:
            progress_callback(msg)
            # トレースバックもログに投入（原因解析用）
            for tb_line in tb.splitlines()[-8:]:  # 末尾8行のみ
                progress_callback(f"    {tb_line}")
        return None

    if not per_article_pdfs:
        msg = "Chromiumレンダリング: 成功ファイルが0件（全記事のHTML→PDF変換が失敗）"
        print(f"[build_articles_pdf_via_chromium] {msg}")
        if progress_callback:
            progress_callback(msg)
        return None

    # 連結
    writer = PdfWriter()
    for pdf_path in per_article_pdfs:
        try:
            reader = PdfReader(str(pdf_path))
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"[merge skip] {pdf_path.name}: {e}")

    with open(output_path, "wb") as f:
        writer.write(f)

    # 作業ファイル片付け
    try:
        for p in per_article_pdfs:
            try:
                p.unlink()
            except Exception:
                pass
        if work_dir.exists() and not any(work_dir.iterdir()):
            work_dir.rmdir()
    except Exception:
        pass

    return output_path if output_path.exists() else None


# ─── 採用記事HTMLを1つのPDFに纏める（フォールバック: テキストベース） ───

def build_articles_pdf_from_html(articles: list[dict], output_path: Path,
                                 topic_name: str = "") -> Path | None:
    """採用記事HTML群を1つの連結PDFとして出力する。

    Parameters
    ----------
    articles : list[dict]
        parse_html_article() の出力 ({"title","body","filepath"}) のリスト
    output_path : Path
        出力PDFパス
    topic_name : str
        表紙・ヘッダ用のトピック名

    Returns
    -------
    Path or None : 生成に成功したパス、失敗時は None
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, PageBreak
        )
        import reportlab.pdfbase.pdfmetrics as metrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return None

    if not articles:
        return None

    # 日本語フォント登録
    font_name = "Helvetica"
    for fp in (r"C:\Windows\Fonts\meiryo.ttc",
               r"C:\Windows\Fonts\YuGothM.ttc",
               r"C:\Windows\Fonts\msgothic.ttc"):
        if Path(fp).exists():
            try:
                metrics.registerFont(TTFont("JpArtFont", fp))
                font_name = "JpArtFont"
                break
            except Exception:
                continue

    doc = SimpleDocTemplate(
        str(output_path), pagesize=A4,
        rightMargin=18 * mm, leftMargin=18 * mm,
        topMargin=22 * mm, bottomMargin=18 * mm,
    )

    cover_title_style = ParagraphStyle(
        "CoverTitle", fontName=font_name, fontSize=22,
        leading=30, spaceAfter=14, textColor="#0D47A1",
    )
    cover_sub_style = ParagraphStyle(
        "CoverSub", fontName=font_name, fontSize=12,
        leading=18, spaceAfter=6, textColor="#555555",
    )
    art_title_style = ParagraphStyle(
        "ArtTitle", fontName=font_name, fontSize=15,
        leading=22, spaceAfter=6, textColor="#1A237E",
    )
    meta_style = ParagraphStyle(
        "Meta", fontName=font_name, fontSize=8.5,
        leading=13, spaceAfter=8, textColor="#666666",
    )
    body_style = ParagraphStyle(
        "Body", fontName=font_name, fontSize=10,
        leading=16, spaceAfter=4,
    )
    h2_style = ParagraphStyle(
        "H2", fontName=font_name, fontSize=12,
        leading=18, spaceBefore=8, spaceAfter=4, textColor="#1976D2",
    )
    toc_item_style = ParagraphStyle(
        "Toc", fontName=font_name, fontSize=10,
        leading=16, leftIndent=6,
    )

    today_str = datetime.now().strftime("%Y-%m-%d")

    def _esc(s: str) -> str:
        return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _extract_meta(filepath: Path) -> dict:
        info = {"date": "", "url": "", "source": ""}
        try:
            if BS4_OK:
                html = filepath.read_text(encoding="utf-8")
                soup = BeautifulSoup(html, "html.parser")
                for s in soup.stripped_strings:
                    m = re.search(r"(\d{4}[-年/]\d{1,2}[-月/]\d{1,2})", s)
                    if m and not info["date"]:
                        info["date"] = m.group(1)
                    if ("ソース" in s or "source" in s.lower()) and not info["source"]:
                        info["source"] = s[:40]
                first_link = soup.find("a", href=re.compile(r"^https?://"))
                if first_link:
                    info["url"] = first_link.get("href", "")[:200]
        except Exception:
            pass
        return info

    def _full_body(art: dict) -> str:
        """HTMLファイルから本文を再抽出（self.articlesは本文が切り詰められている可能性があるため）"""
        fp = art.get("filepath")
        if isinstance(fp, Path) and fp.exists():
            try:
                html = fp.read_text(encoding="utf-8")
                if BS4_OK:
                    soup = BeautifulSoup(html, "html.parser")
                    for tag in soup(["script", "style", "head"]):
                        tag.decompose()
                    return soup.get_text(separator="\n", strip=True)
                return re.sub(r"<[^>]+>", "", html)
            except Exception:
                pass
        return art.get("body", "")

    story = []

    # ── 表紙 ──
    story.append(Spacer(1, 30 * mm))
    story.append(Paragraph(f"{_esc(topic_name)} 採用記事全文集", cover_title_style))
    story.append(Paragraph(f"出力日: {today_str}", cover_sub_style))
    story.append(Paragraph(f"採用記事数: {len(articles)}件", cover_sub_style))
    story.append(Spacer(1, 10 * mm))

    # ── 目次 ──
    story.append(Paragraph("■ 収録記事一覧", h2_style))
    for i, art in enumerate(articles, 1):
        story.append(Paragraph(
            f"{i:02d}. {_esc(art.get('title', ''))}", toc_item_style))
    story.append(PageBreak())

    # ── 各記事 ──
    for i, art in enumerate(articles, 1):
        title = art.get("title", f"記事{i}")
        filepath = art.get("filepath")
        meta = _extract_meta(filepath) if isinstance(filepath, Path) else {}
        body = _full_body(art)

        story.append(Paragraph(f"【記事{i:02d}】 {_esc(title)}", art_title_style))
        meta_lines = []
        if meta.get("date"):
            meta_lines.append(f"公開日: {_esc(meta['date'])}")
        if meta.get("url"):
            meta_lines.append(f"URL: {_esc(meta['url'])}")
        if meta.get("source"):
            meta_lines.append(f"ソース: {_esc(meta['source'])}")
        if meta_lines:
            story.append(Paragraph(" ｜ ".join(meta_lines), meta_style))

        # 本文を段落単位で流し込み
        for raw_line in body.split("\n"):
            line = raw_line.strip()
            if not line:
                story.append(Spacer(1, 2 * mm))
                continue
            # 過度に長い1行は Paragraph で自動折返しされるので特別処理不要
            story.append(Paragraph(_esc(line), body_style))

        if i < len(articles):
            story.append(PageBreak())

    try:
        doc.build(story)
        return output_path
    except Exception:
        return None


# ─── メインGUIアプリ ───

class ContentGeneratorApp:
    def __init__(self, initial_folder: Path | None = None):
        self.root = Tk()
        self.root.title("コンテンツジェネレーター - WeeklyReport & ポッドキャスト原稿生成")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

        self.folder: Path | None = None
        self.pipeline_state: dict = {}
        self.topic_name: str = ""
        self.articles: list[dict] = []

        self._build_ui()

        if initial_folder and initial_folder.is_dir():
            self.root.after(100, lambda: self._load_folder(initial_folder))

    def _build_ui(self):
        root = self.root

        # タイトル
        title_frame = Frame(root, bg="#4CAF50", pady=8)
        title_frame.pack(fill=X)
        Label(title_frame, text="コンテンツジェネレーター  WeeklyReport & ポッドキャスト原稿生成",
              font=("Yu Gothic UI", 13, "bold"), fg="white", bg="#4CAF50").pack()

        # フォルダ選択
        folder_frame = Frame(root, bg="#E8F5E9", pady=4, padx=10)
        folder_frame.pack(fill=X)
        Button(folder_frame, text="フォルダを開く", font=("Yu Gothic UI", 10),
               bg="#388E3C", fg="white", command=self._select_folder).pack(side=LEFT, padx=(0, 8))
        self.folder_label = Label(folder_frame, text="フォルダを選択してください",
                                   font=("Yu Gothic UI", 10), bg="#E8F5E9", fg="#555")
        self.folder_label.pack(side=LEFT)

        # 状態パネル
        info_lf = ttk.LabelFrame(root, text="状態", padding=8)
        info_lf.pack(fill=X, padx=10, pady=6)
        self.info_label = Label(info_lf, text="フォルダを選択してください",
                                 font=("Yu Gothic UI", 10), bg="#f5f5f5", anchor=W)
        self.info_label.pack(fill=X)

        # 生成ボタン
        btn_frame = Frame(root, bg="#f5f5f5", pady=6, padx=10)
        btn_frame.pack(fill=X)

        self.btn_report = Button(btn_frame, text="WeeklyReport生成（PDF）",
                                  font=("Yu Gothic UI", 11, "bold"),
                                  bg="#1976D2", fg="white", height=2,
                                  command=self._generate_report, state=DISABLED)
        self.btn_report.pack(side=LEFT, padx=(0, 8), fill=Y)

        self.btn_podcast = Button(btn_frame, text="ポッドキャスト原稿生成",
                                   font=("Yu Gothic UI", 11, "bold"),
                                   bg="#7B1FA2", fg="white", height=2,
                                   command=self._generate_podcast, state=DISABLED)
        self.btn_podcast.pack(side=LEFT, padx=(0, 8), fill=Y)

        self.btn_both = Button(btn_frame, text="両方まとめて生成",
                                font=("Yu Gothic UI", 11, "bold"),
                                bg="#E65100", fg="white", height=2,
                                command=self._generate_both, state=DISABLED)
        self.btn_both.pack(side=LEFT, fill=Y)

        self.btn_combine_mp3 = Button(btn_frame, text="既存WAV→MP3結合",
                                       font=("Yu Gothic UI", 10),
                                       bg="#00695C", fg="white", height=2,
                                       command=self._combine_existing_wavs, state=DISABLED)
        self.btn_combine_mp3.pack(side=LEFT, padx=(8, 0), fill=Y)

        # プログレスバー
        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill=X, padx=10)

        # ログエリア
        log_lf = ttk.LabelFrame(root, text="ログ", padding=4)
        log_lf.pack(fill=BOTH, expand=True, padx=10, pady=6)
        log_frame = Frame(log_lf)
        log_frame.pack(fill=BOTH, expand=True)
        self.log_text = Text(log_frame, wrap=WORD, font=("Yu Gothic UI", 9),
                              state=DISABLED, bg="#1e1e1e", fg="#d4d4d4")
        log_scroll = Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_scroll.pack(side=RIGHT, fill=Y)

    def _log(self, msg: str):
        def _do():
            self.log_text.config(state=NORMAL)
            self.log_text.insert(END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
            self.log_text.see(END)
            self.log_text.config(state=DISABLED)
        self.root.after(0, _do)

    # ─── フォルダ選択 ───

    def _select_folder(self):
        folder = filedialog.askdirectory(
            initialdir=str(SURVEY_OUTPUT_DIR) if SURVEY_OUTPUT_DIR.exists() else str(BASE_DIR),
            title="調査アウトプットフォルダを選択",
        )
        if folder:
            self._load_folder(Path(folder))

    def _load_folder(self, folder: Path):
        self.folder = folder
        self.folder_label.config(text=folder.name, fg="#1a1a1a")

        self.pipeline_state = load_pipeline_state(folder)
        curation = self.pipeline_state["stages"]["curation"]

        # トピック名抽出（フォルダ名から日付部分を除く）
        folder_name = folder.name
        self.topic_name = re.sub(r"_\d{8}.*$", "", folder_name)

        # 採用記事を収集
        self.articles = get_adopted_articles(folder, self.pipeline_state)

        adopted_count = len(self.articles)
        curation_status = curation.get("status", "pending")
        gen_state = self.pipeline_state["stages"]["generation"]

        # 全記事に判定が付いていれば（pending=0）事実上完了済みとみなす
        articles_map = curation.get("articles", {})
        pending_count = sum(1 for v in articles_map.values() if v == "pending")
        if curation_status != "completed" and len(articles_map) > 0 and pending_count == 0:
            curation_status = "completed"

        info_lines = [
            f"トピック: {self.topic_name}",
            f"採用記事: {adopted_count} 件",
            f"キュレーション状態: {curation_status}",
        ]
        if gen_state.get("pdf_path"):
            info_lines.append(f"PDF: {Path(gen_state['pdf_path']).name}")
        if gen_state.get("script_path"):
            info_lines.append(f"原稿: {Path(gen_state['script_path']).name}")

        self.info_label.config(text="  |  ".join(info_lines))
        self._log(f"フォルダ読み込み完了: {folder.name}")
        self._log(f"採用記事数: {adopted_count}")

        if curation_status != "completed":
            self._log("警告: キュレーションが完了していません。採用記事のみを対象に生成します。")

        if adopted_count == 0:
            messagebox.showwarning("記事なし", "採用済み記事が見つかりません。\n先にキュレーションを完了させてください。")
            return

        if not GEMINI_API_KEY:
            self._log("警告: GEMINI_API_KEY が設定されていません")

        self.btn_report.config(state=NORMAL)
        self.btn_podcast.config(state=NORMAL)
        self.btn_both.config(state=NORMAL)

        # _podcast_work フォルダにWAVがあれば結合ボタンを有効化
        work_dir = folder / "_podcast_work"
        wav_files = list(work_dir.glob("seg_*.wav")) if work_dir.exists() else []
        if wav_files:
            self.btn_combine_mp3.config(state=NORMAL)
            self._log(f"_podcast_work に {len(wav_files)} 個のWAVファイルを検出 → 「既存WAV→MP3結合」が使えます")

    # ─── WeeklyReport生成 ───

    def _generate_report(self):
        if not self.folder or not self.articles:
            return
        self._set_buttons(DISABLED)
        self.progress.start()
        threading.Thread(target=self._do_generate_report, daemon=True).start()

    def _do_generate_report(self):
        import shutil, traceback
        try:
            self._log("WeeklyReport生成を開始...")
            today = datetime.now().strftime("%Y%m%d")
            output_path = None

            # ── リッチレポートモード（lib/weekly_report_generator使用） ──
            if RICH_REPORT_OK and GENAI_OK and GEMINI_API_KEY:
                self._log("リッチレポートモードで生成します（旧来の高品質フォーマット）")

                # articles_data.json 読込 or HTMLから構築
                articles_data_path = self.folder / "articles_data.json"
                if articles_data_path.exists():
                    articles_data = json.loads(
                        articles_data_path.read_text(encoding="utf-8"))
                    self._log(f"articles_data.json 読込完了: {articles_data['total_articles']}件")
                else:
                    self._log("articles_data.json が見つかりません。HTMLから構築します...")
                    articles_data = self._build_articles_data_dict()

                # Gemini クライアント生成
                client = genai.Client(api_key=GEMINI_API_KEY)

                # フロントレポートJSON 生成（Gemini呼び出し）
                # ─ generate_front_report_json 内に429/503リトライあり。
                #   それでも失敗する場合はここで外側リトライを実施する。
                _gemini_max_retries = 3
                front_report = None
                for _ga in range(_gemini_max_retries):
                    try:
                        attempt_label = (f" (試行 {_ga + 1}/{_gemini_max_retries})"
                                         if _ga > 0 else "")
                        self._log(f"Gemini APIでフロントレポートJSON生成中...{attempt_label}")
                        front_report = generate_front_report_json(
                            client, articles_data,
                            progress_callback=lambda msg: self._log(msg),
                        )
                        break  # 成功 → ループ終了
                    except Exception as _ge:
                        if _ga < _gemini_max_retries - 1:
                            import time as _t
                            _wait = 30 * (_ga + 1)  # 30s → 60s
                            self._log(
                                f"⚠ Gemini API 失敗（{_ge}）。"
                                f"{_wait}秒後にリトライします... "
                                f"({_ga + 1}/{_gemini_max_retries})")
                            _t.sleep(_wait)
                        else:
                            raise RuntimeError(
                                f"Gemini API が {_gemini_max_retries} 回試行後も失敗しました: {_ge}"
                            ) from _ge
                if front_report is None:
                    raise RuntimeError("Gemini APIからレスポンスを取得できませんでした")

                # front_report.json 保存
                front_json_path = self.folder / "front_report.json"
                front_json_path.write_text(
                    json.dumps(front_report, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                self._log("front_report.json 保存完了")

                # フロントレポートPDF 生成
                front_pdf_path = self.folder / f"{self.topic_name}_フロントレポート.pdf"
                self._log("リッチPDF生成中（reportlab）...")
                generate_front_report_pdf(front_report, front_pdf_path)
                self._log(f"フロントレポートPDF生成完了: {front_pdf_path.name}")

                # 採用記事HTMLから統合PDFを生成
                articles_combined_pdf = self.folder / f"{self.topic_name}_統合記事.pdf"
                try:
                    self._log(f"採用記事 {len(self.articles)} 件をChromiumで画像付きPDFに変換中...")
                    built = build_articles_pdf_via_chromium(
                        self.articles, articles_combined_pdf,
                        topic_name=self.topic_name,
                        progress_callback=lambda m: self._log(f"  {m}"),
                    )
                    if not (built and articles_combined_pdf.exists()):
                        # Chromium失敗時はテキストベースにフォールバック。
                        # ※ 過去にフォールバック発動が目立たず、WeeklyReportがテキスト版の
                        #   まま納品されてしまう事故があった。ユーザーが気付けるように
                        #   警告マーカーとダイアログで強く通知する。
                        self._log("⚠⚠⚠ Chromium HTMLレンダリングに失敗したため、テキスト版にフォールバックします ⚠⚠⚠")
                        self._log("   → 出力PDFは「HTML表示そのまま」ではなく簡易テキスト版になります")
                        self._log("   → アプリを一度閉じて起動し直してから再実行してください")
                        try:
                            self.root.after(0, lambda: messagebox.showwarning(
                                "HTMLレンダリング失敗",
                                "ChromiumによるHTML→PDF変換に失敗しました。\n"
                                "簡易テキスト版にフォールバックします。\n\n"
                                "HTML表示どおりのPDFが欲しい場合は、\n"
                                "アプリを一度終了してから再起動し、\n"
                                "もう一度「WeeklyReport生成」を押してください。"
                            ))
                        except Exception:
                            pass
                        built = build_articles_pdf_from_html(
                            self.articles, articles_combined_pdf, topic_name=self.topic_name,
                        )
                    if built and articles_combined_pdf.exists():
                        self._log(f"採用記事統合PDF生成完了: {articles_combined_pdf.name}")
                    else:
                        self._log("採用記事統合PDFの生成に失敗しました（スキップして続行）")
                except Exception as e:
                    self._log(f"採用記事統合PDF生成中にエラー: {e}")

                # 最終PDF（フロント + 採用記事）を結合
                final_name = f"{self.topic_name}ウィークリーレポート{today}.pdf"
                final_path = self.folder / final_name
                if articles_combined_pdf.exists():
                    self._log("フロントレポートと採用記事PDFを結合中...")
                    concatenate_pdfs(front_pdf_path, articles_combined_pdf, final_path)
                else:
                    shutil.copy2(front_pdf_path, final_path)
                output_path = final_path
                self._log(f"WeeklyReport生成完了: {final_name}")

            else:
                # ── フォールバック: 簡易テキストベースモード ──
                self._log("フォールバック: 簡易テキストモードで生成します")
                if not RICH_REPORT_OK:
                    self._log("  ※ lib/weekly_report_generator のインポートに失敗")
                output_name = f"{self.topic_name}調査ウィークリーレポート{today}.pdf"
                output_path = self.folder / output_name

                articles_text = self._build_articles_text()
                report_prompt = load_report_prompt(self.topic_name)
                full_prompt = (
                    f"{report_prompt}\n\n"
                    f"## 対象記事（{len(self.articles)}件）\n\n"
                    f"{articles_text}"
                )

                if GEMINI_API_KEY and GENAI_OK:
                    self._log("Gemini APIにリクエスト中...")
                    report_text = call_gemini(full_prompt, GEMINI_API_KEY)
                else:
                    self._log("APIキー未設定: サンプルレポートを生成します")
                    report_text = self._build_sample_report()

                self._log(f"テキスト生成完了（{len(report_text)}字）")
                title = f"{self.topic_name} ウィークリーレポート {today}"
                try:
                    generate_pdf(report_text, output_path, title)
                    self._log(f"PDF生成完了: {output_name}")
                except Exception as pdf_err:
                    self._log(f"PDF生成エラー: {pdf_err}")
                    txt_path = self.folder / f"{self.topic_name}調査ウィークリーレポート{today}.txt"
                    txt_path.write_text(f"{title}\n\n{report_text}", encoding="utf-8")
                    output_path = txt_path

            # 状態更新
            gen_state = self.pipeline_state["stages"]["generation"]
            gen_state["pdf_path"] = str(output_path)
            gen_state["status"] = "in_progress"
            save_pipeline_state(self.folder, self.pipeline_state)

            self.root.after(0, lambda p=output_path.name: messagebox.showinfo(
                "完了", f"WeeklyReport生成完了\n{p}"))
        except Exception as e:
            self._log(f"エラー: {e}")
            self._log(traceback.format_exc())
            self.root.after(0, lambda err=str(e): messagebox.showerror(
                "エラー", f"生成に失敗しました:\n{err}"))
        finally:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self._set_buttons(NORMAL))

    # ─── ポッドキャスト原稿生成 ───

    def _generate_podcast(self):
        if not self.folder or not self.articles:
            return
        self._set_buttons(DISABLED)
        self.progress.start()
        threading.Thread(target=self._do_generate_podcast, daemon=True).start()

    def _do_generate_podcast(self):
        import traceback
        try:
            self._log("ポッドキャスト原稿生成を開始...")
            today = datetime.now().strftime("%Y%m%d")
            output_name = f"{self.topic_name}調査{today}ポッドキャスト原稿.txt"
            output_path = self.folder / output_name

            articles_text = self._build_articles_text()
            podcast_prompt = load_podcast_prompt(self.topic_name)
            full_prompt = (
                f"{podcast_prompt}\n\n"
                f"## 採用記事（{len(self.articles)}件）\n\n"
                f"{articles_text}"
            )

            self._log("Gemini APIにリクエスト中（ポッドキャスト原稿）...")
            if GEMINI_API_KEY and GENAI_OK:
                script_text = call_gemini(full_prompt, GEMINI_API_KEY)
            else:
                self._log("APIキー未設定: サンプル原稿を生成します")
                script_text = self._build_sample_podcast()

            self._log(f"テキスト生成完了（{len(script_text)}字）")
            output_path.write_text(script_text, encoding="utf-8")
            self._log(f"原稿ファイル保存: {output_name}")

            # 状態更新
            gen_state = self.pipeline_state["stages"]["generation"]
            gen_state["script_path"] = str(output_path)
            gen_state["status"] = "in_progress"
            save_pipeline_state(self.folder, self.pipeline_state)

            # ── 音声合成＋MP3生成 ──
            if PODCAST_AUDIO_OK:
                self._log("VoicePeakで音声合成を開始します...")
                mp3_path = self._generate_podcast_audio(script_text, today)
                if mp3_path:
                    self._log(f"MP3生成完了: {mp3_path.name}")
                    gen_state["mp3_path"] = str(mp3_path)
                    save_pipeline_state(self.folder, self.pipeline_state)
                    self.root.after(0, lambda n=mp3_path.name: messagebox.showinfo(
                        "完了", f"ポッドキャスト生成完了\n原稿: {output_name}\nMP3: {n}"))
                    return
                else:
                    self._log("音声合成に失敗しました。原稿ファイルは保存されています。")
            else:
                self._log("VoicePeak関連ライブラリ未インポート: 原稿のみ保存しました")

            self.root.after(0, lambda n=output_name: messagebox.showinfo(
                "完了", f"ポッドキャスト原稿生成完了\n{n}\n（音声合成は未実行）"))
        except Exception as e:
            self._log(f"エラー: {e}")
            self._log(traceback.format_exc())
            self.root.after(0, lambda err=str(e): messagebox.showerror(
                "エラー", f"生成に失敗しました:\n{err}"))
        finally:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self._set_buttons(NORMAL))

    def _generate_podcast_audio(self, script_text: str, date_str: str) -> "Path | None":
        """スクリプトテキストからVoicePeakでWAV生成→ffmpegでMP3化"""
        try:
            work_dir = self.folder / "_podcast_work"
            work_dir.mkdir(exist_ok=True)

            # 既存のWAVをクリア
            for old_wav in work_dir.glob("seg_*.wav"):
                old_wav.unlink()

            # スクリプトをセグメントに分割
            dialogue = parse_dialogue_script(script_text)
            segments = build_segments(dialogue)
            total = len(segments)
            self._log(f"セグメント数: {total}")

            if total == 0:
                self._log("有効なセグメントがありません（F:/M:形式の行が必要です）")
                return None

            # 各セグメントのWAV生成
            ok_count = 0
            failed_indices: list[int] = []
            empty_indices: set[int] = set()
            for i, seg in enumerate(segments):
                speaker = seg["speaker"]
                emotion = seg.get("emotion", "N")
                text = seg["text"].strip()
                if not text:
                    empty_indices.add(i)
                    continue

                narrator = NARRATOR_F if speaker == "F" else NARRATOR_M
                wav_path = work_dir / f"seg_{i:03d}.wav"

                self._log(f"WAV生成中... [{i+1}/{total}] {speaker}[{emotion}]: {text[:30]}...")
                ok = generate_wav(text, narrator, wav_path, speaker=speaker, emotion=emotion)
                if ok:
                    ok_count += 1
                else:
                    self._log(f"  → seg_{i:03d}.wav 生成失敗（後で再試行します）")
                    failed_indices.append(i)

            # ===== 失敗セグメント再生成パス =====
            # スキップせず、VoicePeakをクールダウンしてから最大3ラウンド再試行する。
            # 原稿の全行が音声化されるまで粘る。（空文字セグメントは除外）
            max_retry_rounds = 3
            for retry_round in range(1, max_retry_rounds + 1):
                if not failed_indices:
                    break
                self._log(f"失敗セグメント {len(failed_indices)} 件を再生成します "
                          f"（ラウンド {retry_round}/{max_retry_rounds}）...")
                try:
                    _kill_stale_voicepeak()
                except Exception:
                    pass
                import time as _t
                _t.sleep(3 + retry_round)  # ラウンドごとに待機を伸ばす

                still_failed: list[int] = []
                for i in failed_indices:
                    seg = segments[i]
                    speaker = seg["speaker"]
                    emotion = seg.get("emotion", "N")
                    text = seg["text"].strip()
                    narrator = NARRATOR_F if speaker == "F" else NARRATOR_M
                    wav_path = work_dir / f"seg_{i:03d}.wav"
                    self._log(f"  再試行 [{i+1}/{total}] {speaker}[{emotion}]: {text[:30]}...")
                    if generate_wav(text, narrator, wav_path,
                                    speaker=speaker, emotion=emotion):
                        ok_count += 1
                    else:
                        still_failed.append(i)
                failed_indices = still_failed

            self._log(f"WAV生成完了: {ok_count}/{total - len(empty_indices)} 成功")

            # 欠落セグメントがまだ残っている場合はMP3を作らず中断
            # （スキップしたMP3は全行音声化の要件を満たさないため）
            missing = check_missing_segments(work_dir, total)
            # 空文字セグメントは元々音声化対象外なので missing から除外
            missing = [i for i in missing if i not in empty_indices]
            if missing:
                preview = ", ".join(
                    f"#{i+1}『{segments[i]['text'][:20]}...』" for i in missing[:5]
                )
                self._log(
                    f"【エラー】{len(missing)} セグメントの音声化に失敗しました: {preview}"
                    + ("..." if len(missing) > 5 else "")
                )
                self._log("原稿の全文が音声化されていないためMP3生成を中止します。"
                          "再度「ポッドキャスト原稿生成」ボタンを押してやり直してください。")
                return None

            if ok_count == 0:
                self._log("音声ファイルが1つも生成されませんでした")
                return None

            # WAV → MP3 結合
            mp3_name = f"{self.topic_name}ポッドキャスト{date_str}.mp3"
            mp3_path = self.folder / mp3_name
            self._log(f"MP3に結合中: {mp3_name}")
            success = combine_wavs_to_mp3(work_dir, mp3_path, total)
            if success and mp3_path.exists():
                return mp3_path
            else:
                self._log("MP3結合に失敗しました（ffmpegが利用可能か確認してください）")
                return None
        except Exception as e:
            self._log(f"音声生成エラー: {e}")
            return None

    # ─── 両方まとめて生成 ───

    def _generate_both(self):
        if not self.folder or not self.articles:
            return
        self._set_buttons(DISABLED)
        self.progress.start()
        threading.Thread(target=self._do_generate_both, daemon=True).start()

    def _do_generate_both(self):
        self._do_generate_report()
        self._do_generate_podcast()

        # 両方完了時に generation ステータスを completed に
        gen_state = self.pipeline_state["stages"]["generation"]
        if gen_state.get("pdf_path") and gen_state.get("script_path"):
            gen_state["status"] = "completed"
            gen_state["completed_at"] = datetime.now().isoformat(timespec="seconds")
            save_pipeline_state(self.folder, self.pipeline_state)
            self._log("両方の生成が完了しました")

    # ─── ヘルパー ───

    def _build_articles_text(self) -> str:
        parts = []
        for i, art in enumerate(self.articles, 1):
            parts.append(f"### 記事{i}: {art['title']}\n\n{art['body'][:2000]}\n")
        return "\n---\n".join(parts)

    def _build_articles_data_dict(self) -> dict:
        """articles_data.json がない場合に HTML から articles_data 構造を構築"""
        arts = []
        for i, art in enumerate(self.articles, 1):
            filepath = art["filepath"]
            source, date, url = "", "", ""
            if BS4_OK:
                html_text = filepath.read_text(encoding="utf-8")
                soup = BeautifulSoup(html_text, "html.parser")
                for s in soup.stripped_strings:
                    if "公開日" in s:
                        m = re.search(r"(\d{4}[-年]\d{1,2}[-月]\d{1,2})", s)
                        if m:
                            date = m.group(1)
                first_link = soup.find("a", href=re.compile(r"^https?://"))
                if first_link:
                    url = first_link.get("href", "")
            arts.append({
                "id": f"#{i:02d}",
                "title": art["title"],
                "source": source,
                "country": "",
                "date": date,
                "url": url,
                "summary": art["body"][:500],
                "details": art["body"],
            })
        return {
            "category": self.topic_name,
            "collection_date": datetime.now().strftime("%Y-%m-%d"),
            "total_articles": len(arts),
            "countries": [],
            "articles": arts,
        }

    def _build_sample_report(self) -> str:
        lines = [f"# {self.topic_name} ウィークリーレポート\n"]
        lines.append("## 今週の最重要ニュース\n\n（Gemini APIキーが設定されていないため、サンプルテキストです）\n")
        lines.append("## 技術トレンドの方向性\n\n採用記事を分析した結果を記載します。\n")
        for art in self.articles[:3]:
            lines.append(f"- {art['title']}\n")
        return "\n".join(lines)

    def _build_sample_podcast(self) -> str:
        lines = [
            "F: 今週も最新の技術ニュースをお届けします。",
            "M: よろしくお願いします。今週はどんなニュースがありましたか？",
        ]
        for art in self.articles[:3]:
            lines.append(f"F: 注目の記事として「{art['title']}」があります。")
            lines.append("M: それは興味深いですね。詳しく教えてください。")
            lines.append(f"F: {art['body'][:200]}...といった内容です。")
        lines.append("M: 今週も様々な動きがありましたね。")
        lines.append("F: 引き続き注目していきたいと思います。ありがとうございました。")
        return "\n".join(lines)

    def _combine_existing_wavs(self):
        """_podcast_work 内の既存WAVをMP3に結合する"""
        if not self.folder:
            return
        if not PODCAST_AUDIO_OK:
            messagebox.showwarning("エラー", "podcast_reviewerのインポートに失敗しています")
            return
        work_dir = self.folder / "_podcast_work"
        wav_files = sorted(work_dir.glob("seg_*.wav"))
        if not wav_files:
            messagebox.showwarning("WAVなし", "_podcast_work フォルダにWAVファイルがありません")
            return
        self._set_buttons(DISABLED)
        self.progress.start()
        threading.Thread(target=self._do_combine_wavs, args=(work_dir, len(wav_files)), daemon=True).start()

    def _do_combine_wavs(self, work_dir: Path, seg_count: int):
        try:
            today = datetime.now().strftime("%Y%m%d")
            mp3_name = f"{self.topic_name}ポッドキャスト{today}.mp3"
            mp3_path = self.folder / mp3_name
            self._log(f"既存WAV ({seg_count}個) をMP3に結合中: {mp3_name}")
            success = combine_wavs_to_mp3(work_dir, mp3_path, seg_count)
            if success and mp3_path.exists():
                gen_state = self.pipeline_state["stages"]["generation"]
                gen_state["mp3_path"] = str(mp3_path)
                save_pipeline_state(self.folder, self.pipeline_state)
                self._log(f"MP3結合完了: {mp3_name}")
                self.root.after(0, lambda n=mp3_name: messagebox.showinfo(
                    "完了", f"MP3結合完了\n{n}"))
            else:
                self._log("MP3結合に失敗しました（ffmpegが利用可能か確認してください）")
                self.root.after(0, lambda: messagebox.showerror(
                    "失敗", "MP3結合に失敗しました。\nffmpegがインストールされているか確認してください。"))
        except Exception as e:
            self._log(f"エラー: {e}")
        finally:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self._set_buttons(NORMAL))

    def _set_buttons(self, state):
        self.btn_report.config(state=state)
        self.btn_podcast.config(state=state)
        self.btn_both.config(state=state)
        # combine_mp3ボタンはWAV存在時のみ有効
        if state == NORMAL and self.folder:
            work_dir = self.folder / "_podcast_work"
            wav_files = list(work_dir.glob("seg_*.wav")) if work_dir.exists() else []
            self.btn_combine_mp3.config(state=NORMAL if wav_files else DISABLED)
        else:
            self.btn_combine_mp3.config(state=state)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    initial = None
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        if p.is_dir():
            initial = p

    app = ContentGeneratorApp(initial_folder=initial)
    app.run()
