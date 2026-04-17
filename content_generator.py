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

## 登場人物
- F（スージー）: 女性ホスト。好奇心旺盛でエネルギッシュ。専門知識は浅いが鋭い質問をする。
- M（トロイ）: 男性専門家。冷静で分析的。技術的な内容をわかりやすく解説する。

## 話者記号と感情コード（必ずこの形式で記述すること）
- F[H]: スージーの明るい・嬉しい・驚きの発言
- F[E]: スージーの興奮・テンション高めの発言
- F[Q]: スージーの疑問・質問の発言
- F[S]: スージーの落ち着いた・真面目な発言
- F[N]: スージーの通常発言
- M[N]: トロイの通常説明発言
- M[S]: トロイの専門的・真剣な説明
- M[Q]: トロイの問いかけ・確認の発言

## 原稿の構成
1. オープニング（F[H]とM[N]の挨拶、今週のテーマ紹介）
2. 各記事のトピック解説（記事1件につき6〜10行のやりとり）
   - F[Q]でスージーが質問 → M[N]/M[S]でトロイが解説 → F[H]/F[E]でスージーが反応
3. クロージング（今週のまとめ、来週への期待）

## 分量
- 記事1件につき会話8〜12行（約400〜600字）
- 全体で原稿本文2000〜4000字程度

## 出力形式（厳守）
各行は必ず「話者コード: テキスト」の形式。
空行でトピックを区切ること。
HTMLタグ・markdown・余分な説明文は一切含めないこと。

## 出力例
F[H]: 皆さん、こんにちは！テックトレンドへようこそ。好奇心ナビゲーターのスージーです。
M[N]: こんにちは。テクノロジー動向を分析する専門家のトロイです。今週もよろしくお願いします。
F[E]: トロイさん、今週もすごいニュースがありましたね！
M[S]: そうですね。特に注目すべきは〜
F[Q]: それはどういう意味なんですか？
M[N]: わかりやすく言うと〜ということです。
F[H]: なるほど！それは私たちの生活にも影響しそうですね！

## 注意事項
- 専門用語は必ずスージーの質問→トロイの解説という形でフォローする
- 数値・企業名・技術名は正確に記載する
- 同じ感情コードが連続しすぎないよう自然なメリハリをつける
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
    prompt_file = BASE_DIR / "調査内容ファイル" / f"{topic_name}_podcast_prompt.txt"
    if prompt_file.exists():
        return prompt_file.read_text(encoding="utf-8")
    return DEFAULT_PODCAST_PROMPT


# ─── Gemini API呼び出し ───

def call_gemini(prompt: str, api_key: str, model: str = "gemini-2.5-flash") -> str:
    if not GENAI_OK:
        raise ImportError("google-genai ライブラリがインストールされていません")
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model,
        contents=prompt,
    )
    return response.text


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

    # ─── WeeklyReport生成 ───

    def _generate_report(self):
        if not self.folder or not self.articles:
            return
        self._set_buttons(DISABLED)
        self.progress.start()
        threading.Thread(target=self._do_generate_report, daemon=True).start()

    def _do_generate_report(self):
        try:
            self._log("WeeklyReport生成を開始...")
            today = datetime.now().strftime("%Y%m%d")
            output_name = f"{self.topic_name}調査ウィークリーレポート{today}.pdf"
            output_path = self.folder / output_name

            # 記事テキストを連結
            articles_text = self._build_articles_text()

            # プロンプト読み込み
            report_prompt = load_report_prompt(self.topic_name)
            full_prompt = (
                f"{report_prompt}\n\n"
                f"## 対象記事（{len(self.articles)}件）\n\n"
                f"{articles_text}"
            )

            self._log("Gemini APIにリクエスト中...")
            if GEMINI_API_KEY and GENAI_OK:
                report_text = call_gemini(full_prompt, GEMINI_API_KEY)
            else:
                # APIキーなしの場合はサンプルテキストを生成
                self._log("APIキー未設定: サンプルレポートを生成します")
                report_text = self._build_sample_report()

            self._log(f"テキスト生成完了（{len(report_text)}字）")

            # PDF生成
            self._log("PDFを生成中...")
            title = f"{self.topic_name} ウィークリーレポート {today}"
            try:
                generate_pdf(report_text, output_path, title)
                self._log(f"PDF生成完了: {output_name}")
            except Exception as pdf_err:
                self._log(f"PDF生成エラー: {pdf_err}")
                # テキストファイルとして保存
                txt_path = self.folder / f"{self.topic_name}調査ウィークリーレポート{today}.txt"
                txt_path.write_text(f"{title}\n\n{report_text}", encoding="utf-8")
                self._log(f"テキストファイルとして保存: {txt_path.name}")
                output_path = txt_path

            # 状態更新
            gen_state = self.pipeline_state["stages"]["generation"]
            gen_state["pdf_path"] = str(output_path)
            gen_state["status"] = "in_progress"
            save_pipeline_state(self.folder, self.pipeline_state)

            self.root.after(0, lambda: messagebox.showinfo("完了", f"WeeklyReport生成完了\n{output_path.name}"))
        except Exception as e:
            self._log(f"エラー: {e}")
            self.root.after(0, lambda err=str(e): messagebox.showerror("エラー", f"生成に失敗しました:\n{err}"))
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

            self.root.after(0, lambda: messagebox.showinfo("完了", f"ポッドキャスト原稿生成完了\n{output_name}"))
        except Exception as e:
            self._log(f"エラー: {e}")
            self.root.after(0, lambda err=str(e): messagebox.showerror("エラー", f"生成に失敗しました:\n{err}"))
        finally:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self._set_buttons(NORMAL))

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

    def _set_buttons(self, state):
        self.btn_report.config(state=state)
        self.btn_podcast.config(state=state)
        self.btn_both.config(state=state)

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
