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
        extract_chart_data,
        render_chart_image,
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

TARGET_WIDTH, TARGET_HEIGHT = 1280, 720


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


def _fit_cover_pil(img: "PILImage.Image", w: int, h: int) -> "PILImage.Image":
    """w×h に固定サイズでセンタークロップリサイズ（縦長・横長を問わず均一な矩形に）"""
    img_w, img_h = img.size
    scale = max(w / img_w, h / img_h)
    new_w = max(w, int(img_w * scale))
    new_h = max(h, int(img_h * scale))
    img_resized = img.resize((new_w, new_h), PILImage.LANCZOS)
    left = (new_w - w) // 2
    top  = (new_h - h) // 2
    return img_resized.crop((left, top, left + w, top + h))


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

        # ===== 採否ボタンエリアを先に BOTTOM に固定（常に表示） =====
        bottom_frame = Frame(self.left_panel, bg="#f5f5f5")
        bottom_frame.pack(side=BOTTOM, fill=X)

        # ステータス
        self.status_label = Label(bottom_frame, text="フォルダを選択してください",
                                   font=("Yu Gothic UI", 9), fg="#333", bg="#f5f5f5",
                                   wraplength=380, justify=LEFT)
        self.status_label.pack(fill=X, pady=(2, 0))

        # 採否ボタン（下部固定エリア内）
        action_lf = ttk.LabelFrame(bottom_frame, text="採否決定", padding=6)
        action_lf.pack(fill=X, pady=(0, 4))

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

        btn_sub_frame = Frame(action_lf, bg="#f5f5f5")
        btn_sub_frame.pack(fill=X, pady=2)

        self.btn_prev = Button(btn_sub_frame, text="◀ 前へ戻る",
                                font=("Yu Gothic UI", 10),
                                bg="#546E7A", fg="white",
                                command=self._go_prev, state=DISABLED)
        self.btn_prev.pack(side=LEFT, fill=X, expand=True, padx=(0, 2))

        self.btn_finish = Button(btn_sub_frame, text="完了",
                                  font=("Yu Gothic UI", 10),
                                  bg="#5C6BC0", fg="white",
                                  command=self._finish, state=DISABLED)
        self.btn_finish.pack(side=LEFT, fill=X, expand=True, padx=(2, 0))

        # ===== 上部スクロール可能エリア: 画像プレビュー + 画像選択 =====
        top_canvas = Canvas(self.left_panel, bg="#f5f5f5", highlightthickness=0)
        top_scrollbar = ttk.Scrollbar(self.left_panel, orient=VERTICAL, command=top_canvas.yview)
        top_canvas.configure(yscrollcommand=top_scrollbar.set)
        top_scrollbar.pack(side=RIGHT, fill=Y)
        top_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        top_inner = Frame(top_canvas, bg="#f5f5f5")
        top_canvas_window = top_canvas.create_window((0, 0), window=top_inner, anchor="nw")

        def _on_top_inner_configure(e):
            top_canvas.configure(scrollregion=top_canvas.bbox("all"))
        def _on_top_canvas_configure(e):
            top_canvas.itemconfig(top_canvas_window, width=e.width)
        top_inner.bind("<Configure>", _on_top_inner_configure)
        top_canvas.bind("<Configure>", _on_top_canvas_configure)

        # マウスホイールでスクロール
        def _on_mousewheel(e):
            top_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        top_canvas.bind("<MouseWheel>", _on_mousewheel)
        top_inner.bind("<MouseWheel>", _on_mousewheel)

        # 画像プレビューキャンバス（上部エリア内）
        img_lf = ttk.LabelFrame(top_inner, text="アイキャッチ画像プレビュー", padding=4)
        img_lf.pack(fill=X, pady=(0, 4))
        self.img_canvas = Canvas(img_lf, width=370, height=190, bg="#cccccc")
        self.img_canvas.pack()
        self.img_label_text = Label(img_lf, text="画像なし", font=("Yu Gothic UI", 9), fg="#888")
        self.img_label_text.pack()

        # 画像選択オプション（上部エリア内）
        opt_lf = ttk.LabelFrame(top_inner, text="アイキャッチ画像の選択", padding=6)
        opt_lf.pack(fill=X, pady=(0, 4))

        options = [
            ("keep", "そのまま使用（記事内画像）"),
            ("library", "WPライブラリから選択"),
            ("ai_generate", "AI新規生成（Imagen）"),
            ("chart_generate", "グラフ読取＆再デザイン描画"),
        ]
        for val, label in options:
            rb = Radiobutton(opt_lf, text=label, variable=self.image_option, value=val,
                             font=("Yu Gothic UI", 10), bg="#f5f5f5",
                             command=self._on_image_option_change)
            rb.pack(anchor=W, pady=1)
            rb.bind("<MouseWheel>", _on_mousewheel)

        self.btn_preview_image = Button(top_inner, text="画像をプレビュー",
                                         font=("Yu Gothic UI", 9), bg="#0288D1", fg="white",
                                         command=self._preview_image, state=DISABLED)
        self.btn_preview_image.pack(fill=X, pady=(0, 4))

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
            # ★ 画像なしは目立つ警告表示（採用ブロックと対応）
            self.img_canvas.create_rectangle(0, 0, 370, 190, fill="#fff3e0", outline="#e65100", width=2)
            self.img_canvas.create_text(
                185, 80,
                text="⚠ この記事には画像がありません",
                fill="#c62828", font=("Yu Gothic UI", 10, "bold"),
                width=350, justify="center",
            )
            self.img_canvas.create_text(
                185, 120,
                text="採用するには\n「AI新規生成」または「WPライブラリから選択」\nを選んでください",
                fill="#e65100", font=("Yu Gothic UI", 9),
                width=350, justify="center",
            )
            self.img_label_text.config(text="⚠ 画像なし — 採用には別途画像の選択が必要です", fg="#c62828")
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
        # 370x190にフィット
        img_copy = img.copy()
        img_copy.thumbnail((370, 190), PILImage.LANCZOS)
        self._tk_image = ImageTk.PhotoImage(img_copy)
        self.img_canvas.create_image(185, 95, image=self._tk_image)
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

        self.status_label.config(text="記事内のグラフを検索中...")
        self.btn_adopt.config(state=DISABLED)
        self.btn_reject.config(state=DISABLED)
        threading.Thread(target=self._do_generate_chart_image, daemon=True).start()

    def _do_generate_chart_image(self):
        """記事内のグラフ画像からデータを読み取り、同じ数値で新デザインのグラフを描画する。
        AI画像生成は使用しない（著作権フリーのmatplotlib再描画のみ）。

        ※ detect_chart_in_image（検出ステップ）は廃止。
           extract_chart_data を直接呼び出し、有効なデータが返ってくれば成功とする。
           これにより誤検出（Falseの見逃し）を排除する。
        """
        try:
            if not self.html_files:
                return
            filepath = self.html_files[self.current_index]
            img_urls = get_all_images_from_html(filepath)

            chart_data = None
            tried_count = 0
            for url in img_urls:
                img_src = load_image_from_url_or_path(url, filepath.parent)
                if img_src is None:
                    continue
                tried_count += 1
                buf = io.BytesIO()
                img_src.save(buf, format="PNG")
                img_bytes = buf.getvalue()

                self.root.after(0, lambda n=tried_count: self.status_label.config(
                    text=f"画像{n}枚目からグラフデータを抽出中（Gemini解析）..."))

                # 検出ステップをスキップして直接データ抽出を試みる
                try:
                    candidate = extract_chart_data(img_bytes, GEMINI_API_KEY)
                except Exception as ex:
                    # extract_chart_data が例外を投げた場合（モデルエラーなど）はUIに表示
                    err_msg = str(ex)
                    self.root.after(0, lambda m=err_msg: messagebox.showerror(
                        "API エラー",
                        f"グラフデータ抽出でエラーが発生しました:\n{m}\n\n"
                        "GEMINI_API_KEY が設定されているか確認してください。"))
                    self.root.after(0, lambda: self.image_option.set("keep"))
                    return

                # 有効なデータ（1点以上）があれば採用
                if candidate:
                    has_data = any(
                        len(s.get("data", [])) > 0
                        for s in candidate.get("series", [])
                    )
                    if has_data:
                        chart_data = candidate
                        break

            # どの画像からもグラフデータが取れなかった場合
            if not chart_data:
                msg = (f"記事内の画像（{tried_count}枚）からグラフデータを抽出できませんでした。\n"
                       "・画像が複合インフォグラフィックの場合は正確に抽出できないことがあります\n"
                       "・「そのまま使用」または「AI新規生成」を選択してください。")
                self.root.after(0, lambda: messagebox.showwarning("データ抽出失敗", msg))
                self.root.after(0, lambda: self.image_option.set("keep"))
                return

            # グラフデータの概要をログ
            chart_title = chart_data.get("title", "")
            chart_type  = chart_data.get("chart_type", "bar")
            series_list = chart_data.get("series", [])
            data_count  = sum(len(s.get("data", [])) for s in series_list)
            self.root.after(0, lambda: self.status_label.config(
                text=f"データ抽出完了: {chart_type}グラフ / {len(series_list)}系列 / {data_count}点 → 新グラフを描画中..."))

            # matplotlibで同じ数値・新デザインのグラフを描画（1280×720）
            chart_img = render_chart_image(chart_data, width=1280, height=720)

            if chart_img is None:
                self.root.after(0, lambda: messagebox.showerror(
                    "描画失敗",
                    "グラフの再描画に失敗しました。\n"
                    "matplotlib / numpy がインストールされているか確認してください。"))
                self.root.after(0, lambda: self.image_option.set("keep"))
                return

            self.generated_image = chart_img
            self.root.after(0, lambda: self._display_image(chart_img))

            # 完了ステータス
            summary = f"再描画完了: {chart_type}グラフ"
            if chart_title:
                summary += f"「{chart_title}」"
            summary += f" / {len(series_list)}系列 / {data_count}データポイント（著作権フリー）"
            self.root.after(0, lambda msg=summary: self.status_label.config(text=msg))

        except Exception as e:
            self.root.after(0, lambda err=str(e): messagebox.showerror(
                "エラー", f"グラフ再描画に失敗しました:\n{err}"))
            self.root.after(0, lambda: self.image_option.set("keep"))
        finally:
            self.root.after(0, lambda: self.btn_adopt.config(state=NORMAL))
            self.root.after(0, lambda: self.btn_reject.config(state=NORMAL))

    # ─── 採用処理 ───

    def _adopt(self):
        if not self.html_files or self.current_index >= len(self.html_files):
            return
        filepath = self.html_files[self.current_index]

        # ── 図なし採用チェック（画像がない状態で採用されることを防ぐ）──
        option = self.image_option.get()
        if option == "keep":
            if not self.current_article.get("image_url", ""):
                messagebox.showwarning(
                    "採用不可 — 画像なし",
                    "この記事にはアイキャッチ画像がありません。\n\n"
                    "採用するには以下のいずれかを選択してください：\n"
                    "  ① 「AI新規生成（Imagen）」で新しい画像を作成\n"
                    "  ② 「WPライブラリから選択」で既存画像を選ぶ\n\n"
                    "画像を用意できない場合は「不採用にして次へ」を押してください。")
                return
        elif option in ("library", "ai_generate", "chart_generate"):
            if self.generated_image is None:
                messagebox.showwarning(
                    "採用不可 — 画像未選択",
                    "画像がまだ生成・選択されていません。\n"
                    "「画像をプレビュー」ボタンで画像を生成・選択してから\n"
                    "もう一度「採用して次へ」を押してください。")
                return

        # アイキャッチ画像を確定・保存
        eyecatch_path = self._save_eyecatch_image(filepath)

        # HTMLのタイトル直下（概要の上）の画像をアイキャッチで置き換え
        if eyecatch_path:
            self._embed_eyecatch_in_html(filepath, eyecatch_path)

        # 状態更新
        curation = self.pipeline_state["stages"]["curation"]
        curation.setdefault("articles", {})[filepath.name] = "adopted"
        curation["adopted_count"] = sum(1 for v in curation["articles"].values() if v == "adopted")
        curation["rejected_count"] = sum(1 for v in curation["articles"].values() if v == "rejected")
        self._auto_complete_if_done(curation)
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

    def _embed_eyecatch_in_html(self, html_path: Path, eyecatch_path: Path) -> bool:
        """HTMLのタイトルと概要の間にアイキャッチ画像を埋め込む（既存画像は置き換え）。

        ・.article-image img が存在すれば src をローカルパスに差し替え
        ・.article-image が存在しなければ .header の直後に挿入
        ・いずれの場合も onerror 属性は除去し、ローカルファイル参照に統一
        """
        BS = get_bs4()
        if BS is None:
            return False
        try:
            html_text = html_path.read_text(encoding="utf-8")
            soup = BS(html_text, "html.parser")

            # eyecatch の相対パス（HTML と同じフォルダ）
            rel_src = eyecatch_path.name

            existing_block = soup.select_one(".article-image")

            if existing_block:
                # --- 既存の .article-image ブロックを上書き ---
                img_tag = existing_block.find("img")
                if img_tag:
                    img_tag["src"] = rel_src
                    img_tag["alt"] = html_path.stem
                    img_tag.attrs.pop("onerror", None)   # onerrorは不要
                else:
                    # imgタグがないブロックには新規追加
                    new_img = soup.new_tag(
                        "img", src=rel_src,
                        alt=html_path.stem,
                        style="max-width:100%;height:auto;max-height:400px;object-fit:cover;"
                    )
                    existing_block.clear()
                    existing_block.append(new_img)
            else:
                # --- .article-image ブロックが存在しない → .header 直後に挿入 ---
                header = soup.select_one(".header")
                if header is None:
                    return False

                new_block = soup.new_tag("div", attrs={"class": "article-image"})
                new_img = soup.new_tag(
                    "img", src=rel_src,
                    alt=html_path.stem,
                    style="max-width:100%;height:auto;max-height:400px;object-fit:cover;"
                )
                new_block.append(new_img)
                header.insert_after(new_block)

            html_path.write_text(str(soup), encoding="utf-8")
            return True

        except Exception as e:
            self.status_label.config(text=f"HTML画像埋め込み失敗: {e}")
            return False

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
        self._auto_complete_if_done(curation)
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

    def _auto_complete_if_done(self, curation: dict):
        """全記事に採用/不採用の判定が付いたら自動的に completed に設定する"""
        articles = curation.get("articles", {})
        if not articles:
            return
        all_decided = all(v in ("adopted", "rejected") for v in articles.values())
        # html_files に含まれる全ファイルが判定済みかも確認
        total = len(self.html_files) if hasattr(self, "html_files") else 0
        if all_decided and len(articles) >= total and curation.get("status") != "completed":
            curation["status"] = "completed"
            curation["completed_at"] = datetime.now().isoformat(timespec="seconds")

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
    # サムネイルサイズ（幅×高さ）とセル幅（padding込み）
    THUMB_W, THUMB_H = 170, 96
    CELL_W = THUMB_W + 16   # 186px

    def __init__(self, parent, wp_client: "WordPressClient"):
        self.top = tk.Toplevel(parent)
        self.top.title("WPメディアライブラリから画像を選択")
        self.top.resizable(True, True)
        self.top.grab_set()

        # 最大化で開く
        self.top.state("zoomed")

        self.wp_client = wp_client
        self.selected_image = None
        self._thumbnails = []
        self._page = 1
        self._per_page = 40          # 1ページあたり件数を増加
        self._current_items: list = []  # 最後に取得したアイテム（リサイズ再描画用）
        self._resize_job = None         # リサイズdebounce用

        self._build_ui()
        # 最大化後のウィンドウサイズが確定してから読み込む
        self.top.after(150, lambda: threading.Thread(target=self._load_media, daemon=True).start())

    def _build_ui(self):
        top = self.top

        # ===== 検索バー =====
        search_frame = Frame(top, bg="#f0f0f0", pady=6, padx=10)
        search_frame.pack(fill=X)

        Label(search_frame, text="キーワード検索:", font=("Yu Gothic UI", 10),
              bg="#f0f0f0").pack(side=LEFT)
        self.search_var = StringVar()
        self._search_entry = ttk.Entry(search_frame, textvariable=self.search_var,
                                       width=40, font=("Yu Gothic UI", 10))
        self._search_entry.pack(side=LEFT, padx=(4, 6))

        def _do_search():
            self._page = 1          # 検索時はページを先頭に戻す
            threading.Thread(target=self._load_media, daemon=True).start()

        self._search_entry.bind("<Return>", lambda e: _do_search())
        Button(search_frame, text="  検索  ", font=("Yu Gothic UI", 10),
               bg="#1565C0", fg="white", command=_do_search).pack(side=LEFT)

        Button(search_frame, text="クリア", font=("Yu Gothic UI", 9),
               command=lambda: (self.search_var.set(""), _do_search())
               ).pack(side=LEFT, padx=4)

        self.status_label = Label(top, text="読み込み中...",
                                   font=("Yu Gothic UI", 9), fg="#555", anchor=W)
        self.status_label.pack(fill=X, padx=10, pady=(0, 2))

        # ===== サムネイルグリッド（縦スクロールのみ） =====
        canvas_frame = Frame(top)
        canvas_frame.pack(fill=BOTH, expand=True, padx=8, pady=(0, 4))

        self.canvas = Canvas(canvas_frame, bg="#e8e8e8", highlightthickness=0)
        scroll_y = Scrollbar(canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll_y.set)

        scroll_y.pack(side=RIGHT, fill=Y)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self.grid_frame = Frame(self.canvas, bg="#e8e8e8")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")

        # grid_frameのサイズ変化 → スクロール領域を更新
        self.grid_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        # canvasの幅変化 → grid_frameをcanvas幅に合わせる & 列数を再計算
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # マウスホイール
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(
            int(-1 * (e.delta / 120)), "units"))

        # ===== ページナビ（下部） =====
        nav_frame = Frame(top, bg="#f0f0f0", pady=5, padx=10)
        nav_frame.pack(fill=X, side=BOTTOM)

        self.btn_prev_page = Button(nav_frame, text="◀ 前のページ",
                                     font=("Yu Gothic UI", 10), command=self._prev_page)
        self.btn_prev_page.pack(side=LEFT, padx=4)

        self.btn_next_page = Button(nav_frame, text="次のページ ▶",
                                     font=("Yu Gothic UI", 10), command=self._next_page)
        self.btn_next_page.pack(side=LEFT, padx=4)

        self.page_label = Label(nav_frame, text="", font=("Yu Gothic UI", 9),
                                 bg="#f0f0f0", fg="#555")
        self.page_label.pack(side=LEFT, padx=10)

        Button(nav_frame, text="キャンセル", font=("Yu Gothic UI", 10),
               fg="red", command=self.top.destroy).pack(side=RIGHT, padx=4)

    # ── リサイズ時の再レイアウト ──

    def _on_canvas_resize(self, event):
        """canvasリサイズ時: grid_frameを同幅に、列数を再計算して再描画"""
        new_w = event.width
        self.canvas.itemconfig(self.canvas_window, width=new_w)
        # debounce: 300ms後に再描画
        if self._resize_job:
            self.top.after_cancel(self._resize_job)
        self._resize_job = self.top.after(300, self._rerender_grid)

    def _get_cols(self) -> int:
        """現在のcanvas幅から適切な列数を計算"""
        w = self.canvas.winfo_width()
        if w < 100:
            w = self.top.winfo_width() - 40
        return max(2, w // self.CELL_W)

    def _rerender_grid(self):
        """現在のアイテムを新しい列数で再描画"""
        if self._current_items:
            self._show_thumbnails(self._current_items)

    # ── メディア読み込み ──

    def _load_media(self):
        try:
            self.top.after(0, lambda: self.status_label.config(text="読み込み中..."))
            keyword = self.search_var.get().strip()
            params = {
                "per_page": self._per_page,
                "page": self._page,
                "media_type": "image",
                "orderby": "date",
                "order": "desc",
            }
            if keyword:
                # WordPress REST API: search はタイトル・説明・ファイル名を対象
                params["search"] = keyword

            import requests as _req
            resp = _req.get(
                f"{self.wp_client.api_url}/media",
                auth=self.wp_client.auth,
                params=params,
                timeout=20,
            )
            if resp.status_code != 200:
                self.top.after(0, lambda c=resp.status_code: self.status_label.config(
                    text=f"読込失敗: HTTP {c}"))
                return

            items = resp.json()
            # X-WP-Total ヘッダーから総件数を取得
            total = int(resp.headers.get("X-WP-Total", len(items)))
            total_pages = int(resp.headers.get("X-WP-TotalPages", 1))

            self.top.after(0, lambda: self._on_media_loaded(items, total, total_pages))
        except Exception as e:
            self.top.after(0, lambda err=str(e): self.status_label.config(
                text=f"エラー: {err[:100]}"))

    def _on_media_loaded(self, items: list, total: int, total_pages: int):
        self._current_items = items
        keyword = self.search_var.get().strip()
        kw_info = f'「{keyword}」の検索結果: ' if keyword else ""
        self.status_label.config(
            text=f"{kw_info}全{total}件中 {(self._page-1)*self._per_page+1}〜"
                 f"{min(self._page*self._per_page, total)}件表示")
        self.page_label.config(text=f"ページ {self._page} / {total_pages}")
        self.btn_prev_page.config(state=NORMAL if self._page > 1 else DISABLED)
        self.btn_next_page.config(state=NORMAL if self._page < total_pages else DISABLED)
        self._show_thumbnails(items)

    # ── サムネイル表示 ──

    def _show_thumbnails(self, items: list):
        # 既存ウィジェットをクリア
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self._thumbnails.clear()

        if not items:
            Label(self.grid_frame, text="画像が見つかりません",
                  font=("Yu Gothic UI", 11), bg="#e8e8e8", pady=30).pack()
            return

        cols = self._get_cols()
        tw, th = self.THUMB_W, self.THUMB_H

        for i, item in enumerate(items):
            row, col = divmod(i, cols)
            # medium サイズを優先（thumbnail より高解像度）
            sizes = item.get("media_details", {}).get("sizes", {})
            thumb_url = (sizes.get("medium", {}).get("source_url", "")
                         or sizes.get("thumbnail", {}).get("source_url", "")
                         or item.get("source_url", ""))
            full_url   = item.get("source_url", "")
            img_title  = item.get("title", {}).get("rendered", "")

            cell = Frame(self.grid_frame, bg="#e8e8e8", padx=3, pady=3)
            cell.grid(row=row, column=col, sticky="nw", padx=2, pady=2)

            # ─ 固定サイズ Canvas（サムネイル表示エリア） ─
            # Canvas を使うことで縦長・横長に関わらず常に tw×th の矩形で表示
            thumb_canvas = Canvas(cell, width=tw, height=th,
                                   bg="#aaaaaa", cursor="hand2",
                                   highlightthickness=1,
                                   highlightbackground="#888888")
            thumb_canvas.pack()
            # 読込中テキスト
            thumb_canvas.create_text(tw // 2, th // 2, text="読込中...",
                                      fill="#555555", font=("Yu Gothic UI", 8))

            # ファイル名ラベル
            short_name = (img_title[:24] + "…") if len(img_title) > 24 else img_title
            Label(cell, text=short_name, font=("Yu Gothic UI", 7),
                  bg="#e8e8e8", fg="#555", wraplength=tw, anchor=W).pack(fill=X)

            # 選択ボタン
            select_btn = Button(cell, text="選択", font=("Yu Gothic UI", 8),
                                 bg="#1565C0", fg="white", pady=1, state=DISABLED)
            select_btn.pack(fill=X)

            threading.Thread(
                target=self._load_thumbnail,
                args=(thumb_url, full_url, thumb_canvas, select_btn, img_title),
                daemon=True,
            ).start()

    def _load_thumbnail(self, thumb_url: str, full_url: str,
                        thumb_canvas: "Canvas", btn: "Button", title: str):
        try:
            tw, th = self.THUMB_W, self.THUMB_H
            img = load_image_from_url_or_path(thumb_url, BASE_DIR)
            if img is None:
                img = load_image_from_url_or_path(full_url, BASE_DIR)

            if img:
                # センタークロップで必ず tw×th の固定矩形に
                img_fitted = _fit_cover_pil(img, tw, th)
                tk_img = ImageTk.PhotoImage(img_fitted)
                self._thumbnails.append(tk_img)  # GC防止

                def _setup(c=thumb_canvas, b=btn, ti=tk_img, fu=full_url):
                    c.delete("all")
                    c.create_image(0, 0, anchor="nw", image=ti)
                    c.image = ti  # GC防止
                    c.bind("<Button-1>", lambda e, u=fu: self._select(u))
                    b.config(state=NORMAL, command=lambda u=fu: self._select(u))

                self.top.after(0, _setup)
            else:
                # 画像取得失敗時はテキスト表示
                short = (title[:20] + "…") if len(title) > 20 else (title or "---")
                def _fail(c=thumb_canvas, b=btn, fu=full_url, s=short):
                    c.delete("all")
                    c.create_text(tw // 2, th // 2, text=s,
                                   fill="#555", font=("Yu Gothic UI", 8),
                                   width=tw - 8)
                    b.config(state=NORMAL, command=lambda u=fu: self._select(u))
                self.top.after(0, _fail)
        except Exception:
            pass

    def _select(self, url: str):
        try:
            self.top.after(0, lambda: self.status_label.config(text="画像を取得中..."))
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
