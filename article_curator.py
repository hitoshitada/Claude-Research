"""記事キュレーター - 記事の採否判定＆アイキャッチ画像確定ツール

調査アウトプットフォルダ内の NN_*.html ファイルを1件ずつレビューし、
採用/不採用を判定してアイキャッチ画像を設定する。
"""
import sys
import os
import io
import json
import re
import threading
import urllib.request
from pathlib import Path
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Scrollbar, Text, Listbox,
    Radiobutton, StringVar, messagebox, filedialog,
    BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, N, S,
    VERTICAL, HORIZONTAL, WORD, SINGLE, END, DISABLED, NORMAL,
    Canvas, PhotoImage,
)
from tkinter import ttk

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ─── パス設定 ───
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "調査アウトプット"
WP_ENV_PATH = BASE_DIR / "Python_Auto_Uploader.env"

# ─── wp_uploaderから関数・クラスをインポート ───
try:
    from wp_uploader import (
        generate_image_with_gemini,
        detect_chart_in_image,
        extract_chart_data,
        render_chart_image,
        generate_image_with_chart,
        WordPressClient,
        CATEGORY_MAP,
        detect_category,
    )
    from bs4 import BeautifulSoup
    from PIL import Image as PILImage, ImageTk
    WP_IMPORT_OK = True
except ImportError as _e:
    WP_IMPORT_OK = False
    _WP_IMPORT_ERROR = str(_e)

try:
    from dotenv import load_dotenv
    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# ─── WP認証読み込み ───
_wp_env: dict = {}
if WP_ENV_PATH.exists():
    lines = [ln.strip() for ln in WP_ENV_PATH.read_text(encoding="utf-8").splitlines()
             if ln.strip() and not ln.strip().startswith("#")]
    if len(lines) >= 3:
        _wp_env = {"url": lines[0], "username": lines[1], "password": lines[2]}

TARGET_WIDTH, TARGET_HEIGHT = 1920, 1080


# ─── パイプライン状態管理 ───

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
                "status": "pending",
                "articles": {},
                "adopted_count": 0,
                "rejected_count": 0,
                "completed_at": None,
            },
            "generation": {"status": "pending", "pdf_path": None, "script_path": None, "completed_at": None},
            "podcast_review": {"status": "unreviewed", "review_count": 0, "last_position_sec": 0.0, "completed_at": None},
            "upload": {"status": "pending", "uploaded_count": 0, "completed_at": None},
        },
    }


# ─── HTMLパース ───

def get_bs4():
    if WP_IMPORT_OK:
        return BeautifulSoup
    try:
        from bs4 import BeautifulSoup as BS
        return BS
    except ImportError:
        return None


def parse_html_content(html_path: Path) -> dict:
    """HTMLファイルからタイトル・本文・画像URLを抽出"""
    BS = get_bs4()
    if BS is None:
        return {"title": html_path.stem, "text": "", "image_url": "", "html": ""}

    html = html_path.read_text(encoding="utf-8")
    soup = BS(html, "html.parser")

    title_tag = soup.select_one(".header h1") or soup.find("h1") or soup.find("title")
    title = title_tag.get_text(strip=True) if title_tag else html_path.stem

    img_tag = soup.select_one(".article-image img") or soup.find("img")
    image_url = img_tag.get("src", "") if img_tag else ""

    # 本文テキスト
    for tag in soup(["script", "style"]):
        tag.decompose()
    body_text = soup.get_text(separator="\n", strip=True)

    return {"title": title, "text": body_text[:3000], "image_url": image_url, "html": html}


def get_all_images_from_html(html_path: Path) -> list[str]:
    """HTML内の全img srcを取得"""
    BS = get_bs4()
    if BS is None:
        return []
    soup = BS(html_path.read_text(encoding="utf-8"), "html.parser")
    return [img.get("src", "") for img in soup.find_all("img") if img.get("src")]


def load_image_from_url_or_path(url_or_path: str, base_dir: Path) -> "PILImage.Image | None":
    """URLまたはローカルパスから画像をロード"""
    if not WP_IMPORT_OK:
        return None
    try:
        if url_or_path.startswith("http"):
            req = urllib.request.Request(url_or_path, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = resp.read()
            return PILImage.open(io.BytesIO(data))
        else:
            p = Path(url_or_path)
            if not p.is_absolute():
                p = base_dir / url_or_path
            if p.exists():
                return PILImage.open(str(p))
    except Exception:
        pass
    return None


# ─── メインGUIアプリ ───

class ArticleCuratorApp:
    def __init__(self, initial_folder: Path | None = None):
        self.root = Tk()
        self.root.title("記事キュレーター - 採否判定＆アイキャッチ画像確定")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

        # 状態変数
        self.folder: Path | None = None
        self.html_files: list[Path] = []
        self.current_index: int = 0
        self.pipeline_state: dict = {}
        self.current_article: dict = {}
        self.current_image: "PILImage.Image | None" = None  # 表示中画像
        self.generated_image: "PILImage.Image | None" = None  # 生成・選択済み画像
        self.wp_client: "WordPressClient | None" = None

        # 画像選択
        self.image_option = StringVar(value="keep")

        self._build_ui()

        if initial_folder and initial_folder.is_dir():
            self.root.after(100, lambda: self._load_folder(initial_folder))

    def _build_ui(self):
        root = self.root
        # ===== タイトルバー =====
        title_frame = Frame(root, bg="#1565C0", pady=8)
        title_frame.pack(fill=X)
        Label(title_frame, text="記事キュレーター  採否判定 & アイキャッチ画像確定",
              font=("Yu Gothic UI", 14, "bold"), fg="white", bg="#1565C0").pack()

        # ===== フォルダ選択バー =====
        folder_frame = Frame(root, bg="#e3f2fd", pady=4, padx=10)
        folder_frame.pack(fill=X)
        Button(folder_frame, text="フォルダを開く", font=("Yu Gothic UI", 10),
               bg="#1565C0", fg="white", command=self._select_folder).pack(side=LEFT, padx=(0, 8))
        self.folder_label = Label(folder_frame, text="フォルダを選択してください",
                                   font=("Yu Gothic UI", 10), bg="#e3f2fd", fg="#555")
        self.folder_label.pack(side=LEFT)
        self.progress_label = Label(folder_frame, text="",
                                     font=("Yu Gothic UI", 10, "bold"), bg="#e3f2fd")
        self.progress_label.pack(side=RIGHT)

        # ===== メインコンテンツ（左右分割） =====
        main_frame = Frame(root, bg="#f5f5f5")
        main_frame.pack(fill=BOTH, expand=True, padx=8, pady=6)

        # ----- 左パネル（400px固定）: 画像プレビュー + 画像選択 -----
        self.left_panel = Frame(main_frame, bg="#f5f5f5", width=400)
        self.left_panel.pack(side=LEFT, fill=Y, padx=(0, 6))
        self.left_panel.pack_propagate(False)

        # 画像プレビューキャンバス
        img_lf = ttk.LabelFrame(self.left_panel, text="アイキャッチ画像プレビュー", padding=4)
        img_lf.pack(fill=X, pady=(0, 6))
        self.img_canvas = Canvas(img_lf, width=380, height=213, bg="#cccccc")
        self.img_canvas.pack()
        self.img_label_text = Label(img_lf, text="画像なし", font=("Yu Gothic UI", 9), fg="#888")
        self.img_label_text.pack()

        # 画像選択オプション
        opt_lf = ttk.LabelFrame(self.left_panel, text="アイキャッチ画像の選択", padding=8)
        opt_lf.pack(fill=X, pady=(0, 6))

        options = [
            ("keep", "そのまま使用（記事内画像）"),
            ("library", "WPライブラリから選択"),
            ("ai_generate", "AI新規生成（Imagen）"),
            ("chart_generate", "グラフ読取＆新規生成"),
        ]
        for val, label in options:
            rb = Radiobutton(opt_lf, text=label, variable=self.image_option, value=val,
                             font=("Yu Gothic UI", 10), bg="#f5f5f5",
                             command=self._on_image_option_change)
            rb.pack(anchor=W, pady=1)

        self.btn_preview_image = Button(self.left_panel, text="画像をプレビュー",
                                         font=("Yu Gothic UI", 9), bg="#0288D1", fg="white",
                                         command=self._preview_image, state=DISABLED)
        self.btn_preview_image.pack(fill=X, pady=2)

        # 採否ボタン
        action_lf = ttk.LabelFrame(self.left_panel, text="採否決定", padding=8)
        action_lf.pack(fill=X, pady=(0, 6))

        self.btn_adopt = Button(action_lf, text="採用して次へ ▶",
                                 font=("Yu Gothic UI", 11, "bold"),
                                 bg="#43A047", fg="white", height=2,
                                 command=self._adopt, state=DISABLED)
        self.btn_adopt.pack(fill=X, pady=2)

        self.btn_reject = Button(action_lf, text="不採用にして次へ ▶",
                                  font=("Yu Gothic UI", 11, "bold"),
                                  bg="#E53935", fg="white", height=2,
                                  command=self._reject, state=DISABLED)
        self.btn_reject.pack(fill=X, pady=2)

        self.btn_prev = Button(action_lf, text="◀ 前の記事へ戻る",
                                font=("Yu Gothic UI", 10),
                                bg="#546E7A", fg="white",
                                command=self._go_prev, state=DISABLED)
        self.btn_prev.pack(fill=X, pady=2)

        self.btn_finish = Button(action_lf, text="完了（キュレーション終了）",
                                  font=("Yu Gothic UI", 10),
                                  bg="#5C6BC0", fg="white",
                                  command=self._finish, state=DISABLED)
        self.btn_finish.pack(fill=X, pady=2)

        # ステータス
        self.status_label = Label(self.left_panel, text="フォルダを選択してください",
                                   font=("Yu Gothic UI", 9), fg="#333", bg="#f5f5f5",
                                   wraplength=380, justify=LEFT)
        self.status_label.pack(fill=X, pady=4)

        # ----- 右パネル（伸縮）: 記事番号・タイトル・本文 -----
        right_panel = Frame(main_frame, bg="#f5f5f5")
        right_panel.pack(side=LEFT, fill=BOTH, expand=True)

        self.article_no_label = Label(right_panel, text="",
                                       font=("Yu Gothic UI", 10, "bold"), bg="#f5f5f5", fg="#1565C0")
        self.article_no_label.pack(anchor=W)

        self.title_label = Label(right_panel, text="", font=("Yu Gothic UI", 14, "bold"),
                                  bg="#f5f5f5", fg="#1a1a1a", wraplength=770, justify=LEFT)
        self.title_label.pack(anchor=W, pady=(2, 6))

        text_frame = Frame(right_panel)
        text_frame.pack(fill=BOTH, expand=True)
        self.article_text = Text(text_frame, wrap=WORD, font=("Yu Gothic UI", 10),
                                  state=DISABLED, bg="#fafafa")
        scroll = Scrollbar(text_frame, orient=VERTICAL, command=self.article_text.yview)
        self.article_text.configure(yscrollcommand=scroll.set)
        self.article_text.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)

    # ─── フォルダ選択 ───

    def _select_folder(self):
        folder = filedialog.askdirectory(
            initialdir=str(OUTPUT_DIR) if OUTPUT_DIR.exists() else str(BASE_DIR),
            title="調査アウトプットフォルダを選択",
        )
        if folder:
            self._load_folder(Path(folder))

    def _load_folder(self, folder: Path):
        self.folder = folder
        self.folder_label.config(text=folder.name, fg="#1a1a1a")

        # 数字から始まるHTMLファイルを収集（不採用_から始まるものは対象外）
        all_html = sorted(f for f in folder.glob("*.html")
                          if re.match(r'^\d+', f.name) and not f.name.startswith("不採用_"))
        # 不採用ファイルも含めた全HTMLを確認（状態復元用）
        rejected_html = sorted(f for f in folder.glob("不採用_*.html"))
        # 元の名前（不採用_を除いた名前）を保持
        self.html_files = all_html

        # パイプライン状態読み込み
        self.pipeline_state = load_pipeline_state(folder)
        curation = self.pipeline_state["stages"]["curation"]

        # collection ステータスを更新
        total_articles = len(all_html) + len(rejected_html)
        if total_articles > 0:
            self.pipeline_state["stages"]["collection"]["status"] = "completed"
            self.pipeline_state["stages"]["collection"]["article_count"] = total_articles

        if not all_html and not rejected_html:
            messagebox.showwarning("警告", "対象HTMLファイルが見つかりません\n（NN_*.html形式のファイルが必要です）")
            return

        # 再開: pendingの最初の記事から
        if curation.get("status") in ("in_progress",):
            # 未処理の記事を探す
            start_idx = 0
            for i, f in enumerate(self.html_files):
                status = curation.get("articles", {}).get(f.name, "pending")
                if status == "pending":
                    start_idx = i
                    break
            else:
                start_idx = len(self.html_files)
            self.current_index = start_idx
        else:
            self.current_index = 0

        # 採用数・不採用数をカウント
        adopted = sum(1 for v in curation.get("articles", {}).values() if v == "adopted")
        rejected = sum(1 for v in curation.get("articles", {}).values() if v == "rejected")
        curation["adopted_count"] = adopted
        curation["rejected_count"] = rejected
        curation["status"] = "in_progress"
        save_pipeline_state(folder, self.pipeline_state)

        # WPクライアント初期化
        if _wp_env and WP_IMPORT_OK:
            try:
                self.wp_client = WordPressClient(
                    _wp_env["url"], _wp_env["username"], _wp_env["password"])
            except Exception:
                self.wp_client = None

        self.btn_finish.config(state=NORMAL)
        self._show_article()

    # ─── 記事表示 ───

    def _show_article(self):
        if not self.html_files:
            self.status_label.config(text="HTMLファイルがありません")
            return

        if self.current_index >= len(self.html_files):
            self.progress_label.config(text=f"完了 {len(self.html_files)}件")
            self.article_no_label.config(text="全記事を処理しました")
            self.title_label.config(text="")
            self._set_article_text("")
            self.btn_adopt.config(state=DISABLED)
            self.btn_reject.config(state=DISABLED)
            self.status_label.config(text="全記事の判定が完了しました。「完了」ボタンを押してください。")
            return

        filepath = self.html_files[self.current_index]
        total = len(self.html_files)
        idx = self.current_index + 1
        self.progress_label.config(text=f"{idx} / {total}")

        # 採否状態を確認
        curation = self.pipeline_state["stages"]["curation"]
        current_status = curation.get("articles", {}).get(filepath.name, "pending")
        status_text = {"pending": "未判定", "adopted": "採用済", "rejected": "不採用済"}.get(current_status, "")

        # HTMLパース
        self.current_article = parse_html_content(filepath)
        self.article_no_label.config(
            text=f"記事 {idx}/{total}  [{filepath.name}]  状態: {status_text}")
        self.title_label.config(text=self.current_article["title"])
        self._set_article_text(self.current_article["text"])

        # アイキャッチ候補画像を表示
        self.generated_image = None
        self.image_option.set("keep")
        self._load_and_show_article_image(filepath)

        # ボタン有効化
        self.btn_adopt.config(state=NORMAL)
        self.btn_reject.config(state=NORMAL)
        self.btn_prev.config(state=NORMAL if self.current_index > 0 else DISABLED)
        self.btn_preview_image.config(state=NORMAL)
        self.status_label.config(text=f"画像オプションを選択して「採用して次へ」または「不採用にして次へ」を押してください")

    def _set_article_text(self, text: str):
        self.article_text.config(state=NORMAL)
        self.article_text.delete("1.0", END)
        self.article_text.insert("1.0", text)
        self.article_text.config(state=DISABLED)
        self.article_text.yview_moveto(0)

    def _load_and_show_article_image(self, filepath: Path):
        """記事内の最初の画像を読み込んでキャンバスに表示"""
        if not WP_IMPORT_OK:
            self.img_label_text.config(text="PIL未インストール")
            return
        image_url = self.current_article.get("image_url", "")
        if not image_url:
            self.img_canvas.delete("all")
            self.img_label_text.config(text="画像なし")
            self.current_image = None
            return
        threading.Thread(
            target=self._load_image_bg, args=(image_url, filepath.parent), daemon=True
        ).start()

    def _load_image_bg(self, url_or_path: str, base_dir: Path):
        img = load_image_from_url_or_path(url_or_path, base_dir)
        self.root.after(0, lambda: self._display_image(img))

    def _display_image(self, img: "PILImage.Image | None"):
        self.img_canvas.delete("all")
        if img is None:
            self.img_label_text.config(text="画像読込失敗")
            self.current_image = None
            return
        self.current_image = img
        # 380x213にフィット
        img_copy = img.copy()
        img_copy.thumbnail((380, 213), PILImage.LANCZOS)
        self._tk_image = ImageTk.PhotoImage(img_copy)
        self.img_canvas.create_image(190, 106, image=self._tk_image)
        self.img_label_text.config(text=f"元サイズ: {img.width}x{img.height}")

    # ─── 画像オプション変更 ───

    def _on_image_option_change(self):
        option = self.image_option.get()
        if option == "library":
            self._pick_from_media_library()
        elif option == "ai_generate":
            self._generate_ai_image()
        elif option == "chart_generate":
            self._generate_chart_image()
        # "keep" はそのまま

    def _preview_image(self):
        """現在の画像オプションでプレビューを更新"""
        option = self.image_option.get()
        if option == "keep":
            filepath = self.html_files[self.current_index]
            self._load_and_show_article_image(filepath)
        elif self.generated_image is not None:
            self._display_image(self.generated_image)

    # ─── WPライブラリから選択 ───

    def _pick_from_media_library(self):
        if not WP_IMPORT_OK or not self.wp_client:
            messagebox.showwarning("未設定", "WP認証情報が設定されていないため、WPライブラリを使用できません")
            self.image_option.set("keep")
            return

        dialog = MediaLibraryDialog(self.root, self.wp_client)
        self.root.wait_window(dialog.top)

        if dialog.selected_image is not None:
            self.generated_image = dialog.selected_image
            self._display_image(self.generated_image)
            self.status_label.config(text="WPライブラリから画像を選択しました")
        else:
            self.image_option.set("keep")

    # ─── AI画像生成 ───

    def _generate_ai_image(self):
        if not WP_IMPORT_OK:
            messagebox.showwarning("エラー", "必要なライブラリがインストールされていません")
            self.image_option.set("keep")
            return
        if not GEMINI_API_KEY:
            messagebox.showwarning("APIキー未設定", "GEMINI_API_KEY が設定されていません")
            self.image_option.set("keep")
            return

        self.status_label.config(text="AI画像を生成中...")
        self.btn_adopt.config(state=DISABLED)
        self.btn_reject.config(state=DISABLED)
        threading.Thread(target=self._do_generate_ai_image, daemon=True).start()

    def _do_generate_ai_image(self):
        try:
            import tempfile
            article = self.current_article
            tmp_path = Path(tempfile.mktemp(suffix=".png"))
            result = generate_image_with_gemini(
                GEMINI_API_KEY,
                article["title"],
                article["text"][:500],
                tmp_path,
            )
            img = PILImage.open(str(result))
            self.generated_image = img
            self.root.after(0, lambda: self._display_image(img))
            self.root.after(0, lambda: self.status_label.config(text="AI画像の生成が完了しました"))
        except Exception as e:
            self.root.after(0, lambda err=str(e): messagebox.showerror("生成失敗", f"AI画像生成に失敗しました:\n{err}"))
            self.root.after(0, lambda: self.image_option.set("keep"))
        finally:
            self.root.after(0, lambda: self.btn_adopt.config(state=NORMAL))
            self.root.after(0, lambda: self.btn_reject.config(state=NORMAL))

    # ─── グラフ読取＆新規生成 ───

    def _generate_chart_image(self):
        if not WP_IMPORT_OK:
            messagebox.showwarning("エラー", "必要なライブラリがインストールされていません")
            self.image_option.set("keep")
            return
        if not GEMINI_API_KEY:
            messagebox.showwarning("APIキー未設定", "GEMINI_API_KEY が設定されていません")
            self.image_option.set("keep")
            return

        self.status_label.config(text="グラフを解析中...")
        self.btn_adopt.config(state=DISABLED)
        self.btn_reject.config(state=DISABLED)
        threading.Thread(target=self._do_generate_chart_image, daemon=True).start()

    def _do_generate_chart_image(self):
        try:
            if not self.html_files:
                return
            filepath = self.html_files[self.current_index]
            img_urls = get_all_images_from_html(filepath)

            chart_data = None
            for url in img_urls:
                img = load_image_from_url_or_path(url, filepath.parent)
                if img is None:
                    continue
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                img_bytes = buf.getvalue()
                if detect_chart_in_image(img_bytes, GEMINI_API_KEY):
                    chart_data = extract_chart_data(img_bytes, GEMINI_API_KEY)
                    if chart_data:
                        break

            article = self.current_article
            import tempfile
            tmp_path = Path(tempfile.mktemp(suffix=".png"))

            if chart_data:
                result = generate_image_with_chart(
                    GEMINI_API_KEY, article["title"], article["text"][:500],
                    chart_data, tmp_path)
            else:
                self.root.after(0, lambda: self.status_label.config(
                    text="グラフが見つからないため通常AI生成にフォールバック"))
                result = generate_image_with_gemini(
                    GEMINI_API_KEY, article["title"], article["text"][:500], tmp_path)

            img = PILImage.open(str(result))
            self.generated_image = img
            self.root.after(0, lambda: self._display_image(img))
            self.root.after(0, lambda: self.status_label.config(text="グラフ解析＆新規画像の生成が完了しました"))
        except Exception as e:
            self.root.after(0, lambda err=str(e): messagebox.showerror("生成失敗", f"グラフ画像生成に失敗しました:\n{err}"))
            self.root.after(0, lambda: self.image_option.set("keep"))
        finally:
            self.root.after(0, lambda: self.btn_adopt.config(state=NORMAL))
            self.root.after(0, lambda: self.btn_reject.config(state=NORMAL))

    # ─── 採用処理 ───

    def _adopt(self):
        if not self.html_files or self.current_index >= len(self.html_files):
            return
        filepath = self.html_files[self.current_index]

        # アイキャッチ画像を確定・保存
        eyecatch_path = self._save_eyecatch_image(filepath)

        # 状態更新
        curation = self.pipeline_state["stages"]["curation"]
        curation.setdefault("articles", {})[filepath.name] = "adopted"
        curation["adopted_count"] = sum(1 for v in curation["articles"].values() if v == "adopted")
        curation["rejected_count"] = sum(1 for v in curation["articles"].values() if v == "rejected")
        save_pipeline_state(self.folder, self.pipeline_state)

        info = f"採用: {filepath.name}"
        if eyecatch_path:
            info += f"\nアイキャッチ: {eyecatch_path.name}"
        self.status_label.config(text=info)
        self.current_index += 1
        self._show_article()

    def _save_eyecatch_image(self, html_path: Path) -> "Path | None":
        """アイキャッチ画像を {stem}_eyecatch.png として保存"""
        option = self.image_option.get()

        if option == "keep":
            # 記事内画像をそのまま使用
            image_url = self.current_article.get("image_url", "")
            if not image_url:
                return None
            img = load_image_from_url_or_path(image_url, html_path.parent)
            if img is None:
                return None
        elif option in ("library", "ai_generate", "chart_generate"):
            img = self.generated_image
            if img is None:
                return None
        else:
            return None

        # 保存先パス
        eyecatch_path = html_path.parent / f"{html_path.stem}_eyecatch.png"
        try:
            img_resized = img.resize((TARGET_WIDTH, TARGET_HEIGHT), PILImage.LANCZOS)
            img_resized.save(str(eyecatch_path), "PNG")
            return eyecatch_path
        except Exception as e:
            self.status_label.config(text=f"画像保存失敗: {e}")
            return None

    # ─── 不採用処理 ───

    def _reject(self):
        if not self.html_files or self.current_index >= len(self.html_files):
            return
        filepath = self.html_files[self.current_index]

        # ファイル名を「不採用_NN_*.html」にリネーム
        new_name = f"不採用_{filepath.name}"
        new_path = filepath.parent / new_name
        try:
            filepath.rename(new_path)
        except Exception as e:
            messagebox.showerror("エラー", f"リネームに失敗しました: {e}")
            return

        # 同名PNG/JPGもリネーム
        for ext in [".png", ".jpg", ".jpeg"]:
            img_path = filepath.parent / f"{filepath.stem}{ext}"
            if img_path.exists():
                try:
                    img_path.rename(filepath.parent / f"不採用_{filepath.stem}{ext}")
                except Exception:
                    pass

        # html_filesリストから削除
        self.html_files[self.current_index] = new_path  # 参照を更新（インデックスのずれを防ぐ）
        # 実際には不採用ファイルはリストから除外
        self.html_files.pop(self.current_index)

        # 状態更新（元のファイル名で記録）
        curation = self.pipeline_state["stages"]["curation"]
        curation.setdefault("articles", {})[filepath.name] = "rejected"
        curation["adopted_count"] = sum(1 for v in curation["articles"].values() if v == "adopted")
        curation["rejected_count"] = sum(1 for v in curation["articles"].values() if v == "rejected")
        save_pipeline_state(self.folder, self.pipeline_state)

        self.status_label.config(text=f"不採用: {filepath.name} → {new_name}")
        # インデックスは変えない（次の記事が詰まってくる）
        self._show_article()

    # ─── 前の記事へ ───

    def _go_prev(self):
        if self.current_index > 0:
            self.current_index -= 1
            # 現在のhtml_filesに不採用が含まれている可能性を考慮
            self._show_article()

    # ─── 完了 ───

    def _finish(self):
        if not self.folder:
            self.root.destroy()
            return

        curation = self.pipeline_state["stages"]["curation"]
        pending = sum(1 for v in curation.get("articles", {}).values() if v == "pending")
        unprocessed = len(self.html_files) - self.current_index

        if unprocessed > 0:
            if not messagebox.askyesno("完了確認",
                    f"未処理の記事が {unprocessed} 件あります。\n"
                    "このまま終了しますか？（再起動時に続きから再開できます）"):
                return

        curation["status"] = "completed"
        curation["completed_at"] = datetime.now().isoformat(timespec="seconds")
        adopted = sum(1 for v in curation.get("articles", {}).values() if v == "adopted")
        rejected = sum(1 for v in curation.get("articles", {}).values() if v == "rejected")
        curation["adopted_count"] = adopted
        curation["rejected_count"] = rejected
        save_pipeline_state(self.folder, self.pipeline_state)

        messagebox.showinfo("完了",
                f"キュレーション完了\n採用: {adopted}件\n不採用: {rejected}件\n"
                f"状態を _pipeline_state.json に保存しました")
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ─── WPメディアライブラリ選択ダイアログ ───

class MediaLibraryDialog:
    def __init__(self, parent, wp_client: "WordPressClient"):
        self.top = tk.Toplevel(parent)
        self.top.title("WPメディアライブラリから画像を選択")
        self.top.geometry("880x640")
        self.top.resizable(True, True)
        self.top.grab_set()

        self.wp_client = wp_client
        self.selected_image = None
        self._thumbnails = []
        self._page = 1
        self._per_page = 20

        self._build_ui()
        threading.Thread(target=self._load_media, daemon=True).start()

    def _build_ui(self):
        top = self.top

        # 検索バー
        search_frame = Frame(top, pady=6, padx=8)
        search_frame.pack(fill=X)
        Label(search_frame, text="検索:", font=("Yu Gothic UI", 10)).pack(side=LEFT)
        self.search_var = StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30,
                                  font=("Yu Gothic UI", 10))
        search_entry.pack(side=LEFT, padx=4)
        Button(search_frame, text="検索", font=("Yu Gothic UI", 10),
               bg="#1565C0", fg="white",
               command=lambda: threading.Thread(target=self._load_media, daemon=True).start()
               ).pack(side=LEFT)

        self.status_label = Label(top, text="読み込み中...", font=("Yu Gothic UI", 9), fg="#555")
        self.status_label.pack(anchor=W, padx=8)

        # スクロール可能なサムネイルグリッド
        canvas_frame = Frame(top)
        canvas_frame.pack(fill=BOTH, expand=True, padx=8, pady=4)

        self.canvas = Canvas(canvas_frame, bg="#eeeeee")
        scroll_y = Scrollbar(canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        scroll_x = Scrollbar(canvas_frame, orient=HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.pack(side=BOTTOM, fill=X)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self.grid_frame = Frame(self.canvas, bg="#eeeeee")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>",
                              lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # ページナビ
        nav_frame = Frame(top, pady=4)
        nav_frame.pack(fill=X, padx=8)
        Button(nav_frame, text="◀ 前のページ", command=self._prev_page,
               font=("Yu Gothic UI", 10)).pack(side=LEFT, padx=4)
        Button(nav_frame, text="次のページ ▶", command=self._next_page,
               font=("Yu Gothic UI", 10)).pack(side=LEFT, padx=4)
        Button(nav_frame, text="キャンセル", command=self.top.destroy,
               font=("Yu Gothic UI", 10), fg="red").pack(side=RIGHT, padx=4)

    def _load_media(self):
        try:
            keyword = self.search_var.get().strip() if hasattr(self, "search_var") else ""
            params = {"per_page": self._per_page, "page": self._page, "media_type": "image"}
            if keyword:
                params["search"] = keyword
            import requests as _req
            resp = _req.get(
                f"{self.wp_client.api_url}/media",
                auth=self.wp_client.auth,
                params=params,
                timeout=15,
            )
            if resp.status_code != 200:
                self.top.after(0, lambda: self.status_label.config(text=f"読込失敗: HTTP {resp.status_code}"))
                return
            items = resp.json()
            self.top.after(0, lambda: self._show_thumbnails(items))
        except Exception as e:
            self.top.after(0, lambda err=str(e): self.status_label.config(text=f"エラー: {err[:80]}"))

    def _show_thumbnails(self, items: list):
        # グリッドをクリア
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self._thumbnails.clear()

        if not items:
            Label(self.grid_frame, text="画像が見つかりません",
                  font=("Yu Gothic UI", 11), bg="#eeeeee").pack(padx=20, pady=20)
            self.status_label.config(text="画像なし")
            return

        self.status_label.config(text=f"{len(items)}件表示")
        cols = 4
        for i, item in enumerate(items):
            row, col = divmod(i, cols)
            thumb_url = (item.get("media_details", {}).get("sizes", {})
                         .get("thumbnail", {}).get("source_url", "")
                         or item.get("source_url", ""))
            cell = Frame(self.grid_frame, bg="#eeeeee", padx=4, pady=4)
            cell.grid(row=row, column=col, sticky="nw")

            btn = Button(cell, text="選択", font=("Yu Gothic UI", 9), bg="#1565C0", fg="white")
            btn.pack()

            # サムネイルを非同期ロード
            full_url = item.get("source_url", "")
            threading.Thread(
                target=self._load_thumbnail,
                args=(thumb_url, full_url, btn, item.get("title", {}).get("rendered", "")),
                daemon=True
            ).start()

    def _load_thumbnail(self, thumb_url: str, full_url: str, btn: Button, title: str):
        try:
            img = load_image_from_url_or_path(thumb_url, BASE_DIR)
            if img is None:
                img = load_image_from_url_or_path(full_url, BASE_DIR)
            if img:
                img.thumbnail((160, 90), PILImage.LANCZOS)
                tk_img = ImageTk.PhotoImage(img)
                self._thumbnails.append(tk_img)

                def _setup(btn=btn, tk_img=tk_img, full_url=full_url):
                    btn.config(image=tk_img, compound="top",
                               command=lambda: self._select(full_url))
                    btn.image = tk_img

                self.top.after(0, _setup)
            else:
                self.top.after(0, lambda: btn.config(
                    command=lambda: self._select(full_url), text=title[:20] or "選択"))
        except Exception:
            pass

    def _select(self, url: str):
        try:
            img = load_image_from_url_or_path(url, BASE_DIR)
            self.selected_image = img
        except Exception:
            self.selected_image = None
        self.top.destroy()

    def _prev_page(self):
        if self._page > 1:
            self._page -= 1
            threading.Thread(target=self._load_media, daemon=True).start()

    def _next_page(self):
        self._page += 1
        threading.Thread(target=self._load_media, daemon=True).start()


# ─── tkinter インポートを補完 ───
import tkinter as tk


if __name__ == "__main__":
    initial = None
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        if p.is_dir():
            initial = p

    app = ArticleCuratorApp(initial_folder=initial)
    app.run()
