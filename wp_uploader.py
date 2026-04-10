"""WordPress記事自動アップロードツール

調査アウトプットフォルダのHTMLファイルをWordPressに自動投稿する。
- フォルダ選択 → Chrome表示 → 閲覧完了 → アップロード判定
- 既存画像 or Gemini AI生成画像をアイキャッチに設定
- 概要・詳細を緑帯ヘッダー付きで投稿
"""

import io
import os
import re
import shutil
import subprocess
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from pathlib import Path
from datetime import datetime
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from PIL import Image as PILImage, ImageTk
from google import genai
from google.genai import types

# .envファイル読み込み
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# WordPress専用envファイル読み込み
# フォーマット: 1行目=サイトURL, 2行目=ユーザー名, 3行目=アプリパスワード
WP_ENV_PATH = BASE_DIR / "Python_Auto_Uploader.env"
_wp_env: dict[str, str] = {}
if WP_ENV_PATH.exists():
    lines = [
        ln.strip()
        for ln in WP_ENV_PATH.read_text(encoding="utf-8").splitlines()
        if ln.strip() and not ln.strip().startswith("#")
    ]
    if len(lines) >= 3:
        _wp_env["WP_URL"] = lines[0]
        _wp_env["WP_USERNAME"] = lines[1]
        _wp_env["WP_APP_PASSWORD"] = lines[2]
    elif len(lines) == 1:
        _wp_env["WP_APP_PASSWORD"] = lines[0]

# =====================================================================
# 定数
# =====================================================================
OUTPUT_DIR = BASE_DIR / "調査アウトプット"

# WordPress設定（Python_Auto_Uploader.env → .env → 空文字）
WP_URL = _wp_env.get("WP_URL", os.getenv("WP_URL", ""))
WP_USERNAME = _wp_env.get("WP_USERNAME", os.getenv("WP_USERNAME", ""))
WP_APP_PASSWORD = _wp_env.get("WP_APP_PASSWORD", os.getenv("WP_APP_PASSWORD", ""))

# Gemini API設定
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
IMAGE_MODEL = "imagen-4.0-generate-001"

# 画像サイズ
TARGET_WIDTH = 1920
TARGET_HEIGHT = 1080

# =====================================================================
# カテゴリーマッピング（フォルダ名のキーワード → WPカテゴリー名）
# =====================================================================
CATEGORY_MAP: list[tuple[list[str], str]] = [
    (["全固体電池"],                     "全固体電池"),
    (["ペロブスカイト", "太陽電池"],      "ペロブスカイト太陽電池"),
    (["水素"],                           "水素エネルギー"),
    (["蓄電", "蓄エネ"],                 "次世代蓄電"),
    (["半導体", "PLP", "後工程"],        "半導体後工程"),
    (["AI", "機械学習", "ディープラーニング"], "AI・機械学習"),
    (["量子コンピュータ", "量子計算"],    "量子コンピュータ"),
    (["光通信", "フォトニクス"],          "光通信・フォトニクス"),
    (["細胞培養", "細胞支持体", "バイオリアクター"], "細胞培養"),
    (["iPS", "再生医療"],                "iPS細胞・再生医療"),
    (["創薬", "DDS", "ドラッグデリバリー"], "創薬・DDS"),
    (["バイオセンサー"],                  "バイオセンサー"),
    (["機能性材料"],                      "機能性材料"),
    (["高分子", "樹脂", "ポリマー"],      "高分子・樹脂"),
    (["ナノテク", "ナノ材料"],            "ナノテクノロジー"),
    (["接着", "封止"],                    "接着・封止材"),
]


def detect_category(folder_name: str) -> str | None:
    """フォルダ名からカテゴリーを自動判定する"""
    for keywords, category in CATEGORY_MAP:
        for kw in keywords:
            if kw.lower() in folder_name.lower():
                return category
    return None


# =====================================================================
# HTML記事パーサー
# =====================================================================
def parse_article_html(filepath: Path) -> dict:
    """HTMLファイルから記事データを抽出する"""
    html = filepath.read_text(encoding="utf-8")
    soup = BeautifulSoup(html, "html.parser")

    # タイトル
    title_tag = soup.select_one(".header h1")
    title = title_tag.get_text(strip=True) if title_tag else (soup.title.string or "無題")

    # メタ情報（公開日、出典、国）
    meta_spans = soup.select(".header .meta span")
    publish_date = ""
    source_name = ""
    country = ""
    if len(meta_spans) >= 1:
        publish_date = meta_spans[0].get_text(strip=True).replace("公開日 ", "")
    if len(meta_spans) >= 2:
        source_name = meta_spans[1].get_text(strip=True)
    if len(meta_spans) >= 3:
        country = meta_spans[2].get_text(strip=True)

    # 画像URL
    img_tag = soup.select_one(".article-image img")
    image_url = img_tag.get("src", "") if img_tag else ""

    # 概要（HTMLタグを保持）
    overview_div = soup.select_one(".overview")
    summary = overview_div.decode_contents() if overview_div else ""

    # 詳細（HTMLタグを保持）
    detail_div = soup.select_one(".detail")
    detail = detail_div.decode_contents() if detail_div else ""

    # 元記事URL
    source_link = soup.select_one(".source a")
    source_url = source_link.get("href", "") if source_link else ""

    return {
        "title": title,
        "publish_date": publish_date,
        "source_name": source_name,
        "country": country,
        "image_url": image_url,
        "summary": summary,
        "detail": detail,
        "source_url": source_url,
        "filepath": filepath,
    }


# =====================================================================
# WordPress投稿コンテンツのフォーマット（緑帯ヘッダー付き）
# =====================================================================
def format_wp_content(summary: str, detail: str, source_url: str = "") -> str:
    """WordPress投稿用HTMLコンテンツを生成する"""
    content = (
        '<div style="background-color:#4CAF50;color:white;padding:12px 20px;'
        'font-size:1.2em;font-weight:bold;border-radius:4px 4px 0 0;'
        'margin-top:20px;">概要</div>\n'
        '<div style="background-color:#e8f5e9;padding:15px 20px;'
        'border-radius:0 0 4px 4px;margin-bottom:30px;line-height:1.8;">\n'
        f'{summary}\n</div>\n\n'
        '<div style="background-color:#4CAF50;color:white;padding:12px 20px;'
        'font-size:1.2em;font-weight:bold;border-radius:4px 4px 0 0;'
        'margin-top:20px;">詳細</div>\n'
        '<div style="padding:15px 20px;border:1px solid #e0e0e0;border-top:none;'
        'border-radius:0 0 4px 4px;margin-bottom:30px;line-height:1.8;">\n'
        f'{detail}\n</div>\n'
    )

    if source_url:
        content += (
            '\n<p style="text-align:right;font-size:0.9em;color:#666;">'
            f'元記事: <a href="{source_url}" target="_blank" rel="noopener">'
            f'{source_url}</a></p>\n'
        )

    return content


# =====================================================================
# WordPressクライアント
# =====================================================================
class WordPressClient:
    """WordPress REST APIクライアント"""

    def __init__(self, url: str, username: str, app_password: str):
        self.base_url = url.rstrip("/")
        self.api_url = f"{self.base_url}/wp-json/wp/v2"
        self.auth = (username, app_password)

    def test_connection(self) -> bool:
        """接続テスト"""
        resp = requests.get(
            f"{self.api_url}/users/me",
            auth=self.auth,
            timeout=10,
        )
        return resp.status_code == 200

    def upload_media(
        self, image_data: bytes, filename: str, mime_type: str = "image/png"
    ) -> int:
        """メディアファイルをアップロードし、メディアIDを返す

        マルチパート形式でアップロードし、日本語ファイル名はASCII安全名に変換する。
        """
        # 日本語ファイル名をASCII安全な名前に変換
        ext = Path(filename).suffix or ".png"
        # ファイル種別に応じたプレフィックス
        if ext.lower() == ".pdf":
            prefix = "report"
        elif ext.lower() == ".mp3":
            prefix = "podcast"
        else:
            prefix = "article"
        safe_name = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"

        # PDF/MP3はサイズが大きいのでタイムアウトを長めに設定
        upload_timeout = 180 if ext.lower() in (".pdf", ".mp3") else 60

        resp = requests.post(
            f"{self.api_url}/media",
            auth=self.auth,
            files={
                "file": (safe_name, image_data, mime_type),
            },
            timeout=upload_timeout,
        )
        if resp.status_code in (200, 201):
            media_id = resp.json().get("id")
            if media_id:
                return media_id
        raise Exception(
            f"メディアアップロード失敗: {resp.status_code} {resp.text[:300]}"
        )

    def get_or_create_category(self, category_name: str) -> int | None:
        """カテゴリーをWordPressから検索し、なければ新規作成してIDを返す"""
        # 既存カテゴリーを検索
        resp = requests.get(
            f"{self.api_url}/categories",
            auth=self.auth,
            params={"search": category_name, "per_page": 100},
            timeout=10,
        )
        if resp.status_code == 200:
            for cat in resp.json():
                if cat.get("name") == category_name:
                    return cat["id"]

        # 見つからなければ新規作成
        resp = requests.post(
            f"{self.api_url}/categories",
            auth=self.auth,
            json={"name": category_name},
            timeout=10,
        )
        if resp.status_code in (200, 201):
            return resp.json().get("id")
        return None

    def search_post_by_title(self, title: str, log_func=None) -> dict | None:
        """タイトルで投稿を検索し、最初にマッチした投稿を返す

        context=edit で raw コンテンツも取得する。
        各ステータス（publish / draft / private）を個別に検索する。
        完全一致 → 部分一致（タイトルが検索語を含む）の順で探す。
        """
        def _log(msg):
            if log_func:
                log_func(msg)

        def _normalize(text: str) -> str:
            """比較用にタイトルを正規化"""
            text = BeautifulSoup(text, "html.parser").get_text()
            # 全角スペース・半角スペース・改行を除去
            text = re.sub(r'[\s\u3000]+', '', text)
            # WeeklyReport ↔ ウィークリーレポート を統一
            text = re.sub(r'(?i)weekly\s*report', 'ウィークリーレポート', text)
            return text.strip()

        search_normalized = _normalize(title)
        _log(f"  [検索] タイトル: 「{title}」 (正規化: 「{search_normalized}」)")

        # 検索キーワードを複数用意
        # WordPress検索APIはキーワードベースなので、
        # タイトル表記の揺れ（英語/日本語、接着/接着剤 等）に対応するため
        # 複数の検索語で網羅的に候補を収集する
        category_part = title.replace("ウィークリーレポート", "").strip()
        search_terms = [
            title,                          # 接着・封止材ウィークリーレポート
            category_part,                  # 接着・封止材
            f"{category_part}WeeklyReport", # 接着・封止材WeeklyReport
            "WeeklyReport",                 # WeeklyReport（全WR投稿をヒット）
            "ウィークリーレポート",            # ウィークリーレポート（同上）
        ]
        # CATEGORY_MAP のキーワードも追加（「接着」「封止」等の個別キーワード）
        for keywords, cat_name in CATEGORY_MAP:
            if cat_name == category_part:
                search_terms.extend(keywords)
                break
        search_terms = list(dict.fromkeys(search_terms))  # 重複除去

        all_candidates = []  # (post, post_normalized, status)
        seen_ids = set()

        for search_term in search_terms:
            for status in ("publish", "draft", "private"):
                try:
                    resp = requests.get(
                        f"{self.api_url}/posts",
                        auth=self.auth,
                        params={
                            "search": search_term,
                            "per_page": 20,
                            "status": status,
                            "context": "edit",
                        },
                        timeout=15,
                    )
                    if resp.status_code != 200:
                        _log(f"  [検索] 「{search_term}」 status={status}"
                             f" → HTTP {resp.status_code}")
                        continue

                    posts = resp.json()
                    if posts:
                        _log(f"  [検索] 「{search_term}」 status={status}"
                             f" → {len(posts)}件ヒット")

                    for post in posts:
                        post_id = post.get("id")
                        if post_id in seen_ids:
                            continue
                        seen_ids.add(post_id)

                        raw_title = post.get("title", {})
                        if isinstance(raw_title, dict):
                            post_title_text = raw_title.get(
                                "raw", raw_title.get("rendered", "")
                            )
                        else:
                            post_title_text = str(raw_title)

                        post_normalized = _normalize(post_title_text)
                        _log(f"    候補: 「{post_title_text.strip()[:50]}」"
                             f" (正規化: 「{post_normalized}」) [ID:{post_id}]")

                        all_candidates.append((post, post_normalized, status))

                except Exception as e:
                    _log(f"  [検索] 「{search_term}」 status={status} → 例外: {e}")
                    continue

        if not all_candidates:
            _log(f"  [検索] 候補が0件です")
            return None

        # ステータス優先度: publish > draft > private
        STATUS_PRIORITY = {"publish": 0, "draft": 1, "private": 2}

        def _best_match(matches: list) -> dict:
            """複数マッチから公開投稿を優先して返す"""
            matches.sort(key=lambda x: STATUS_PRIORITY.get(x[2], 9))
            return matches[0][0]

        # パス1: 完全一致
        exact = [(p, pn, s) for p, pn, s in all_candidates
                 if pn == search_normalized]
        if exact:
            best = _best_match(exact)
            _log(f"  [検索] ✓ 完全一致: ID={best.get('id')}")
            return best

        # パス2: 部分一致
        partial = [(p, pn, s) for p, pn, s in all_candidates
                   if search_normalized in pn or pn in search_normalized]
        if partial:
            best = _best_match(partial)
            _log(f"  [検索] ✓ 部分一致: ID={best.get('id')}")
            return best

        # パス3: キーワード一致（カテゴリ名の主要キーワード + 「ウィークリーレポート」を含む）
        category_part = search_normalized.replace("ウィークリーレポート", "")
        keywords = [kw for kw in re.split(r'[・/]', category_part) if len(kw) >= 2]
        kw_matches = []
        for post, post_normalized, status in all_candidates:
            has_report = "ウィークリーレポート" in post_normalized
            keyword_match = any(kw in post_normalized for kw in keywords) if keywords else False
            if has_report and keyword_match:
                kw_matches.append((post, post_normalized, status))

        if kw_matches:
            best = _best_match(kw_matches)
            _log(f"  [検索] ✓ キーワード一致: ID={best.get('id')}"
                 f" (キーワード: {keywords})")
            return best

        _log(f"  [検索] ✗ 一致する投稿が見つかりませんでした")
        return None

    def get_post_raw_content(self, post_id: int) -> str:
        """投稿のrawコンテンツを取得する"""
        resp = requests.get(
            f"{self.api_url}/posts/{post_id}",
            auth=self.auth,
            params={"context": "edit"},
            timeout=15,
        )
        if resp.status_code == 200:
            data = resp.json()
            content_obj = data.get("content", {})
            if isinstance(content_obj, dict):
                # raw があればそれを、なければ rendered を使用
                return content_obj.get("raw") or content_obj.get("rendered", "")
            return str(content_obj)
        raise Exception(
            f"投稿コンテンツ取得失敗: {resp.status_code} {resp.text[:200]}"
        )

    def update_post(self, post_id: int, content: str,
                    bump_to_top: bool = False) -> dict:
        """既存投稿のコンテンツを更新する

        Args:
            post_id: 投稿ID
            content: 新しいコンテンツ
            bump_to_top: Trueの場合、投稿日時を現在時刻に更新して
                         最新投稿としてトップに表示させる
        """
        payload = {"content": content}
        if bump_to_top:
            # sticky=true で投稿一覧のトップに固定表示
            payload["sticky"] = True
            # 投稿日時も現在に更新（日本時間）
            payload["date"] = datetime.now().strftime(
                "%Y-%m-%dT%H:%M:%S"
            )
        resp = requests.post(
            f"{self.api_url}/posts/{post_id}",
            auth=self.auth,
            json=payload,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            return resp.json()
        raise Exception(
            f"投稿更新失敗: {resp.status_code} {resp.text[:300]}"
        )

    def create_post(
        self,
        title: str,
        content: str,
        featured_media_id: int | None = None,
        category_ids: list[int] | None = None,
        status: str = "draft",
    ) -> dict:
        """記事を投稿する"""
        data = {
            "title": title,
            "content": content,
            "status": status,
        }
        if featured_media_id:
            data["featured_media"] = featured_media_id
        if category_ids:
            data["categories"] = category_ids

        resp = requests.post(
            f"{self.api_url}/posts",
            auth=self.auth,
            json=data,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            return resp.json()
        raise Exception(
            f"記事投稿失敗: {resp.status_code} {resp.text[:300]}"
        )


# =====================================================================
# Gemini画像生成（NanoBanana2）
# =====================================================================
def generate_image_with_gemini(
    api_key: str,
    title: str,
    summary_text: str,
    output_path: Path,
) -> Path:
    """Imagen 4.0 APIで記事に関連するアイキャッチ画像を生成する

    生成後、1920×1080にリサイズして保存する。
    """
    client = genai.Client(api_key=api_key)

    # 概要からHTMLタグを除去してプロンプトに使用
    clean_summary = BeautifulSoup(summary_text, "html.parser").get_text()[:400]

    prompt = (
        "Create a professional, high-quality photorealistic illustration "
        "suitable as a blog featured image. "
        f"Topic: {title}. "
        f"Context: {clean_summary}. "
        "Style: Modern, clean, cinematic lighting, wide landscape composition. "
        "IMPORTANT: Do NOT include any text, letters, words, sentences, labels, "
        "captions, watermarks, logos, or written characters of any language in the image. "
        "The image must be purely visual with zero text elements."
    )

    response = client.models.generate_images(
        model=IMAGE_MODEL,
        prompt=prompt,
        config=types.GenerateImagesConfig(
            number_of_images=1,
            aspect_ratio="16:9",
        ),
    )

    if not response.generated_images:
        raise Exception("画像が生成されませんでした（安全フィルターの可能性あり）")

    # 生成画像を取得 → PILで1920×1080にリサイズして保存
    gen_img = response.generated_images[0].image
    pil_img = PILImage.open(io.BytesIO(gen_img.image_bytes))
    pil_img = pil_img.resize((TARGET_WIDTH, TARGET_HEIGHT), PILImage.LANCZOS)
    pil_img.save(str(output_path), "PNG")

    return output_path


# =====================================================================
# メインGUIアプリケーション
# =====================================================================
class WPUploaderApp(tk.Tk):
    """WordPress記事アップローダー"""

    def __init__(self):
        super().__init__()

        self.title("WordPress 記事アップローダー")
        self.geometry("780x650")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        # 状態
        self.html_files: list[Path] = []
        self.current_index = 0
        self.current_article: dict = {}
        self.wp_client: WordPressClient | None = None
        self.detected_category: str | None = None  # フォルダ名から判定
        self.selected_folder: Path | None = None   # 選択中のフォルダ
        self._chrome_process: subprocess.Popen | None = None

        self._build_ui()
        self._init_wp_client()

    # ---------------------------------------------------------------
    # UI構築
    # ---------------------------------------------------------------
    def _build_ui(self):
        style = ttk.Style()
        style.configure("Green.TButton", font=("Meiryo", 10, "bold"))

        # --- WordPress設定 ---
        wp_frame = ttk.LabelFrame(self, text="WordPress設定", padding=8)
        wp_frame.pack(fill="x", padx=10, pady=(10, 4))

        row0 = ttk.Frame(wp_frame)
        row0.pack(fill="x")
        ttk.Label(row0, text="サイトURL:").pack(side="left")
        self.wp_url_var = tk.StringVar(value=WP_URL)
        ttk.Entry(row0, textvariable=self.wp_url_var, width=35).pack(
            side="left", padx=(4, 12)
        )
        ttk.Label(row0, text="ユーザー名:").pack(side="left")
        self.wp_user_var = tk.StringVar(value=WP_USERNAME)
        ttk.Entry(row0, textvariable=self.wp_user_var, width=18).pack(
            side="left", padx=(4, 0)
        )

        row1 = ttk.Frame(wp_frame)
        row1.pack(fill="x", pady=(4, 0))
        ttk.Label(row1, text="アプリパスワード:").pack(side="left")
        self.wp_pass_var = tk.StringVar(value=WP_APP_PASSWORD)
        ttk.Entry(row1, textvariable=self.wp_pass_var, width=35, show="*").pack(
            side="left", padx=(4, 12)
        )
        ttk.Button(row1, text="接続テスト", command=self._test_wp_connection).pack(
            side="left"
        )
        self.wp_status_label = ttk.Label(row1, text="", foreground="gray")
        self.wp_status_label.pack(side="left", padx=8)

        # --- フォルダ選択 ---
        folder_frame = ttk.LabelFrame(self, text="フォルダ選択", padding=8)
        folder_frame.pack(fill="x", padx=10, pady=4)

        ttk.Button(folder_frame, text="フォルダを選択", command=self._select_folder).pack(
            side="left"
        )
        self.folder_label = ttk.Label(folder_frame, text="未選択", foreground="gray")
        self.folder_label.pack(side="left", padx=10)
        self.file_count_label = ttk.Label(folder_frame, text="")
        self.file_count_label.pack(side="right")

        # --- 記事情報 ---
        article_frame = ttk.LabelFrame(self, text="現在の記事", padding=8)
        article_frame.pack(fill="x", padx=10, pady=4)

        self.progress_label = ttk.Label(
            article_frame, text="記事: -/-", font=("Meiryo", 10)
        )
        self.progress_label.pack(anchor="w")

        self.title_label = ttk.Label(
            article_frame, text="", font=("Meiryo", 11, "bold"), wraplength=720
        )
        self.title_label.pack(anchor="w", pady=(4, 0))

        self.meta_label = ttk.Label(article_frame, text="", foreground="#666")
        self.meta_label.pack(anchor="w")

        # --- アクションボタン ---
        btn_frame = ttk.Frame(self, padding=(10, 6))
        btn_frame.pack(fill="x")

        self.complete_btn = tk.Button(
            btn_frame,
            text="  閲覧完了  ",
            command=self._on_reading_complete,
            state="disabled",
            bg="#4CAF50",
            fg="white",
            font=("Meiryo", 11, "bold"),
            relief="raised",
            padx=16,
            pady=6,
        )
        self.complete_btn.pack(side="left", padx=5)

        self.skip_btn = ttk.Button(
            btn_frame,
            text="スキップ（次の記事へ）",
            command=self._next_article,
            state="disabled",
        )
        self.skip_btn.pack(side="left", padx=5)

        self.skip_all_btn = tk.Button(
            btn_frame,
            text="すべての記事をスキップ",
            command=self._skip_all_articles,
            state="disabled",
            bg="#FF9800",
            fg="white",
            font=("Meiryo", 10),
            relief="raised",
            padx=10,
            pady=4,
        )
        self.skip_all_btn.pack(side="left", padx=15)

        # --- ログ ---
        log_frame = ttk.LabelFrame(self, text="ログ", padding=4)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=10, font=("Consolas", 9), wrap="word"
        )
        self.log_text.pack(fill="both", expand=True)

    # ---------------------------------------------------------------
    # WordPress接続
    # ---------------------------------------------------------------
    def _init_wp_client(self):
        """初期設定からWPクライアントを生成"""
        url = self.wp_url_var.get().strip()
        user = self.wp_user_var.get().strip()
        pw = self.wp_pass_var.get().strip()
        if url and user and pw:
            self.wp_client = WordPressClient(url, user, pw)

    def _build_wp_client(self) -> bool:
        """UIの入力値からWPクライアントを再構築"""
        url = self.wp_url_var.get().strip()
        user = self.wp_user_var.get().strip()
        pw = self.wp_pass_var.get().strip()
        if not (url and user and pw):
            messagebox.showerror("エラー", "WordPress設定（URL・ユーザー名・アプリパスワード）をすべて入力してください")
            return False
        self.wp_client = WordPressClient(url, user, pw)
        return True

    def _test_wp_connection(self):
        """接続テスト"""
        if not self._build_wp_client():
            return
        self.wp_status_label.config(text="接続中...", foreground="gray")
        self._log("WordPress接続テスト中...")

        def task():
            try:
                ok = self.wp_client.test_connection()
                self.after(0, lambda: self._on_wp_test_result(ok))
            except Exception as e:
                self.after(0, lambda: self._on_wp_test_result(False, str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _on_wp_test_result(self, success: bool, error: str = ""):
        if success:
            self.wp_status_label.config(text="接続OK", foreground="green")
            self._log("✓ WordPress接続成功")
        else:
            self.wp_status_label.config(text="接続失敗", foreground="red")
            self._log(f"✗ WordPress接続失敗: {error}")
            messagebox.showerror("接続失敗", f"WordPressに接続できませんでした\n{error}")

    # ---------------------------------------------------------------
    # フォルダ選択 & 記事一覧
    # ---------------------------------------------------------------
    def _select_folder(self):
        folder = filedialog.askdirectory(
            initialdir=str(OUTPUT_DIR),
            title="調査アウトプットフォルダを選択",
        )
        if not folder:
            return

        folder_path = Path(folder)
        self.selected_folder = folder_path
        self.folder_label.config(text=folder_path.name, foreground="black")

        # HTMLファイル取得（概要ファイルを除外、ソート済み）
        self.html_files = sorted(
            f
            for f in folder_path.glob("*.html")
            if not f.name.endswith("_概要.html")
        )

        count = len(self.html_files)
        self.file_count_label.config(text=f"{count}件のHTMLファイル")

        if count == 0:
            messagebox.showwarning("警告", "HTMLファイルが見つかりません")
            return

        # フォルダ名からカテゴリー自動判定
        self.detected_category = detect_category(folder_path.name)
        cat_msg = f"カテゴリー: {self.detected_category}" if self.detected_category else "カテゴリー: 未判定"
        self._log(f"フォルダ選択: {folder_path.name}（{count}件）| {cat_msg}")
        self.current_index = 0
        self._show_current_article()

    # ---------------------------------------------------------------
    # 記事表示
    # ---------------------------------------------------------------
    def _show_current_article(self):
        if self.current_index >= len(self.html_files):
            self._log("=" * 40)
            self._log("すべての記事の処理が完了しました")
            self.progress_label.config(text="完了")
            self.title_label.config(text="すべての記事を処理しました")
            self.meta_label.config(text="")
            self.complete_btn.config(state="disabled")
            self.skip_btn.config(state="disabled")
            self.skip_all_btn.config(state="disabled")
            # PDFウィークリーレポートのアップロード処理へ
            self._handle_pdf_upload()
            return

        filepath = self.html_files[self.current_index]
        self.current_article = parse_article_html(filepath)

        total = len(self.html_files)
        idx = self.current_index + 1

        self.progress_label.config(text=f"記事: {idx}/{total}")
        self.title_label.config(text=self.current_article["title"])
        self.meta_label.config(
            text=(
                f'{self.current_article["publish_date"]} | '
                f'{self.current_article["source_name"]} | '
                f'{self.current_article["country"]}'
            )
        )

        # Chromeでファイルを開く（独立ウィンドウで起動し、後でプロセスごと閉じる）
        file_url = filepath.as_uri()
        self._open_chrome(file_url)
        self._log(f"[{idx}/{total}] ブラウザで表示: {filepath.name}")

        self.complete_btn.config(state="normal")
        self.skip_btn.config(state="normal")
        self.skip_all_btn.config(state="normal")

    # ---------------------------------------------------------------
    # Chrome制御（一時プロファイルで独立プロセスとして起動）
    # ---------------------------------------------------------------
    def _open_chrome(self, url: str):
        """Chromeを一時プロファイルで独立起動する（確実に閉じられる）"""
        self._close_chrome()  # 前のがあれば閉じる

        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(
                r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
            ),
        ]
        chrome_exe = None
        for p in chrome_paths:
            if os.path.exists(p):
                chrome_exe = p
                break

        if chrome_exe:
            # 一時プロファイルを作成してChromeを独立プロセスで起動
            # これにより既存Chromeに統合されず、terminate()で確実に閉じられる
            self._chrome_tmp_dir = tempfile.mkdtemp(prefix="chrome_preview_")
            self._chrome_process = subprocess.Popen(
                [
                    chrome_exe,
                    f"--user-data-dir={self._chrome_tmp_dir}",
                    "--no-first-run",
                    "--no-default-browser-check",
                    url,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        else:
            import webbrowser
            webbrowser.open(url)

    def _close_chrome(self):
        """起動したChromeプロセスを終了してウィンドウを閉じる"""
        if self._chrome_process and self._chrome_process.poll() is None:
            try:
                self._chrome_process.terminate()
                self._chrome_process.wait(timeout=5)
            except Exception:
                try:
                    self._chrome_process.kill()
                    self._chrome_process.wait(timeout=3)
                except Exception:
                    pass
        self._chrome_process = None

        # 一時プロファイルを削除
        tmp_dir = getattr(self, "_chrome_tmp_dir", None)
        if tmp_dir and os.path.exists(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass
            self._chrome_tmp_dir = None

    # ---------------------------------------------------------------
    # 閲覧完了 → アップロード判定
    # ---------------------------------------------------------------
    def _on_reading_complete(self):
        self._close_chrome()
        self.complete_btn.config(state="disabled")
        self.skip_btn.config(state="disabled")

        # アップロードするか確認
        upload = messagebox.askyesno(
            "アップロード確認",
            f"この記事をWordPressにアップロードしますか？\n\n"
            f"「{self.current_article['title']}」",
        )

        if not upload:
            self._log("→ アップしない（スキップ）")
            self._next_article()
            return

        # WP接続確認
        if not self._build_wp_client():
            self._enable_buttons()
            return

        # 写真の選択
        self._ask_photo_choice()

    def _ask_photo_choice(self):
        """アイキャッチ画像の選択"""
        has_image = bool(self.current_article.get("image_url"))

        if has_image:
            use_existing = messagebox.askyesno(
                "画像選択",
                "記事の既存画像をアイキャッチ画像として使用しますか？\n\n"
                "「はい」→ 既存画像をそのまま使用\n"
                "「いいえ」→ AIで新規画像を生成",
            )
        else:
            messagebox.showinfo(
                "画像",
                "記事に画像がないため、AIで画像を生成します。",
            )
            use_existing = False

        if use_existing:
            self._upload_with_existing_image()
        else:
            self._generate_ai_image()

    # ---------------------------------------------------------------
    # 既存画像でアップロード
    # ---------------------------------------------------------------
    def _upload_with_existing_image(self):
        self._log("既存画像をダウンロード中...")

        def task():
            try:
                img_url = self.current_article["image_url"]
                resp = requests.get(
                    img_url,
                    timeout=15,
                    headers={"User-Agent": "Mozilla/5.0"},
                )
                resp.raise_for_status()

                image_data = resp.content
                content_type = resp.headers.get("Content-Type", "image/jpeg")
                if "jpeg" in content_type or "jpg" in content_type:
                    ext = "jpg"
                    mime = "image/jpeg"
                elif "png" in content_type:
                    ext = "png"
                    mime = "image/png"
                elif "webp" in content_type:
                    ext = "webp"
                    mime = "image/webp"
                else:
                    ext = "jpg"
                    mime = "image/jpeg"

                filename = self.current_article["filepath"].stem + f".{ext}"
                self.after(0, lambda: self._do_wp_upload(image_data, filename, mime))
            except Exception as e:
                self.after(
                    0,
                    lambda: self._on_error(f"画像ダウンロード失敗: {e}"),
                )

        threading.Thread(target=task, daemon=True).start()

    # ---------------------------------------------------------------
    # AI画像生成
    # ---------------------------------------------------------------
    def _generate_ai_image(self):
        if not GEMINI_API_KEY:
            messagebox.showerror(
                "エラー",
                "GEMINI_API_KEYが.envに設定されていません",
            )
            self._enable_buttons()
            return

        html_path = self.current_article["filepath"]
        output_path = html_path.with_suffix(".png")

        self._log("AI画像を生成中（しばらくお待ちください）...")

        def task():
            try:
                generate_image_with_gemini(
                    GEMINI_API_KEY,
                    self.current_article["title"],
                    self.current_article["summary"],
                    output_path,
                )
                self.after(0, lambda: self._confirm_generated_image(output_path))
            except Exception as e:
                self.after(0, lambda: self._on_error(f"画像生成失敗: {e}"))

        threading.Thread(target=task, daemon=True).start()

    def _confirm_generated_image(self, image_path: Path):
        """生成画像のプレビュー確認ダイアログ"""
        self._log(f"画像生成完了: {image_path.name}")

        preview = tk.Toplevel(self)
        preview.title("生成画像の確認")
        preview.geometry("820x560")
        preview.grab_set()
        preview.focus_set()

        # 画像表示
        pil_img = PILImage.open(str(image_path))
        display_img = pil_img.copy()
        display_img.thumbnail((790, 450))
        photo = ImageTk.PhotoImage(display_img)

        img_label = ttk.Label(preview, image=photo)
        img_label.image = photo  # 参照保持
        img_label.pack(pady=(10, 5))

        ttk.Label(
            preview,
            text=f"{image_path.name}  ({TARGET_WIDTH}×{TARGET_HEIGHT}px)",
            foreground="gray",
        ).pack()

        btn_frame = ttk.Frame(preview)
        btn_frame.pack(pady=12)

        def on_accept():
            preview.destroy()
            image_data = image_path.read_bytes()
            self._do_wp_upload(image_data, image_path.name, "image/png")

        def on_regenerate():
            preview.destroy()
            self._generate_ai_image()

        def on_cancel():
            preview.destroy()
            self._log("→ 画像キャンセル、記事スキップ")
            self._next_article()

        tk.Button(
            btn_frame,
            text="  この画像を使用  ",
            command=on_accept,
            bg="#4CAF50",
            fg="white",
            font=("Meiryo", 10, "bold"),
            padx=12,
            pady=4,
        ).pack(side="left", padx=8)

        ttk.Button(btn_frame, text="再生成", command=on_regenerate).pack(
            side="left", padx=8
        )
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel).pack(
            side="left", padx=8
        )

    # ---------------------------------------------------------------
    # WordPressアップロード実行
    # ---------------------------------------------------------------
    def _do_wp_upload(self, image_data: bytes, image_filename: str, mime_type: str):
        """メディアアップロード → 記事投稿"""
        self._log("WordPressにアップロード中...")

        def task():
            try:
                # メディアアップロード
                media_id = self.wp_client.upload_media(
                    image_data, image_filename, mime_type
                )
                self._log_safe(f"  メディアアップロード完了 (ID: {media_id})")

                # 投稿コンテンツ作成（緑帯ヘッダー付き）
                content = format_wp_content(
                    self.current_article["summary"],
                    self.current_article["detail"],
                    self.current_article.get("source_url", ""),
                )

                # カテゴリーID取得
                category_ids = None
                if self.detected_category:
                    cat_id = self.wp_client.get_or_create_category(
                        self.detected_category
                    )
                    if cat_id:
                        category_ids = [cat_id]
                        self._log_safe(
                            f"  カテゴリー設定: {self.detected_category} (ID: {cat_id})"
                        )

                # 記事投稿（下書き）
                result = self.wp_client.create_post(
                    title=self.current_article["title"],
                    content=content,
                    featured_media_id=media_id,
                    category_ids=category_ids,
                    status="draft",
                )

                post_id = result.get("id", "?")
                post_link = result.get("link", "")
                self.after(0, lambda: self._on_upload_success(post_id, post_link))
            except Exception as e:
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _on_upload_success(self, post_id, post_link):
        self._log(f"✓ 投稿完了（下書き） ID: {post_id}")
        if post_link:
            self._log(f"  URL: {post_link}")
        self._next_article()

    # ---------------------------------------------------------------
    # PDF・ポッドキャスト → WPアップロード
    # ---------------------------------------------------------------
    def _handle_pdf_upload(self):
        """全記事処理後、フォルダ内のPDFとポッドキャストMP3を処理する"""
        if not self.selected_folder:
            self._finish_all()
            return

        # 「ウィークリーレポート」を含むPDFを検索
        pdf_files = sorted(
            f
            for f in self.selected_folder.glob("*.pdf")
            if "ウィークリーレポート" in f.name
        )

        # ポッドキャストMP3を検索
        mp3_files = sorted(
            f
            for f in self.selected_folder.glob("*.mp3")
            if "ポッドキャスト" in f.name
        )

        upload_files = pdf_files + mp3_files

        if not upload_files:
            self._log("PDF・ポッドキャストファイルが見つかりませんでした")
            self._finish_all()
            return

        if not self.detected_category:
            self._log("カテゴリー未判定のため、ファイルアップロードをスキップします")
            self._finish_all()
            return

        # WPクライアント確認
        if not self.wp_client and not self._build_wp_client():
            self._finish_all()
            return

        # 検出ファイルをログ出力
        for f in upload_files:
            file_type = "PDF" if f.suffix.lower() == ".pdf" else "ポッドキャスト"
            self._log(f"{file_type}ファイル検出: {f.name}")

        # 確認ダイアログ
        file_names = "\n".join(f"  ・{f.name}" for f in upload_files)
        proceed = messagebox.askyesno(
            "ファイルアップロード確認",
            f"以下のファイルをウィークリーレポート投稿に追加しますか？\n\n"
            f"{file_names}\n\n"
            f"対象投稿: 「{self.detected_category}ウィークリーレポート」",
        )

        if not proceed:
            self._log("→ ファイルアップロードをスキップ")
            self._finish_all()
            return

        # バックグラウンドで実行
        self._log("ファイルアップロード処理を開始...")
        self.title_label.config(text="PDF・ポッドキャストをアップロード中...")

        def task():
            try:
                self._upload_files_to_weekly_report(upload_files)
                self.after(0, self._finish_all)
            except Exception as e:
                self.after(0, lambda: self._on_error(f"ファイルアップロード失敗: {e}"))

        threading.Thread(target=task, daemon=True).start()

    def _upload_files_to_weekly_report(self, files: list[Path]):
        """PDF/MP3をまとめてアップロードし、ウィークリーレポート投稿を構成する。

        構成:
        - 最新のPDF/MP3 → 直接表示（ファイル名が見える状態）
        - 過去のPDF/MP3 → 「過去のレポート」アコーディオン内に折りたたみ
        """

        # --- 1. 全ファイルをメディアライブラリにアップロード ---
        new_blocks = []  # 今回アップロードしたブロック
        for file_path in files:
            ext = file_path.suffix.lower()
            is_mp3 = ext == ".mp3"
            file_type_label = "ポッドキャスト" if is_mp3 else "PDF"
            mime_type = "audio/mpeg" if is_mp3 else "application/pdf"

            self._log_safe(f"  {file_type_label}アップロード中: {file_path.name}")
            self._log_safe(f"  ファイルサイズ: {file_path.stat().st_size / 1024:.0f} KB")
            file_data = file_path.read_bytes()

            try:
                media_id = self.wp_client.upload_media(file_data, file_path.name, mime_type)
            except Exception as e:
                self._log_safe(f"  ✗ メディアアップロード失敗: {e}")
                continue
            self._log_safe(f"  メディアアップロード完了 (ID: {media_id})")

            resp = requests.get(
                f"{self.wp_client.api_url}/media/{media_id}",
                auth=self.wp_client.auth, timeout=15,
            )
            if resp.status_code != 200:
                self._log_safe(f"  ✗ メディア情報取得失敗: {resp.status_code}")
                continue
            file_url = resp.json().get("source_url", "")
            if not file_url:
                self._log_safe(f"  ✗ {file_type_label}のURLを取得できませんでした")
                continue
            self._log_safe(f"  {file_type_label} URL: {file_url}")

            display_label = self._make_file_display_label(file_path.name)

            if is_mp3:
                block = (
                    f'<!-- wp:paragraph -->\n'
                    f'<p><strong>🎙 {display_label}</strong></p>\n'
                    f'<!-- /wp:paragraph -->\n'
                    f'<!-- wp:audio {{"id":{media_id}}} -->\n'
                    f'<figure class="wp-block-audio">'
                    f'<audio controls src="{file_url}"></audio></figure>\n'
                    f'<!-- /wp:audio -->\n'
                    f'<!-- wp:file {{"id":{media_id},"href":"{file_url}"}} -->\n'
                    f'<div class="wp-block-file">'
                    f'<a id="wp-block-file--media-{media_id}" '
                    f'href="{file_url}">{file_path.name}</a>'
                    f'<a href="{file_url}" '
                    f'class="wp-block-file__button wp-element-button" '
                    f'download aria-describedby="wp-block-file--media-{media_id}">'
                    f'ダウンロード</a></div>\n'
                    f'<!-- /wp:file -->\n'
                )
            else:
                block = (
                    f'<!-- wp:paragraph -->\n'
                    f'<p><strong>📄 {display_label}</strong></p>\n'
                    f'<!-- /wp:paragraph -->\n'
                    f'<!-- wp:file {{"id":{media_id},"href":"{file_url}"}} -->\n'
                    f'<div class="wp-block-file">'
                    f'<a id="wp-block-file--media-{media_id}" '
                    f'href="{file_url}">{display_label}</a>'
                    f'<a href="{file_url}" '
                    f'class="wp-block-file__button wp-element-button" '
                    f'download aria-describedby="wp-block-file--media-{media_id}">'
                    f'ダウンロード</a></div>\n'
                    f'<!-- /wp:file -->\n'
                )
            new_blocks.append(block)

        if not new_blocks:
            self._log_safe("  アップロードできるファイルがありませんでした")
            return

        # --- 2. ウィークリーレポート投稿を検索 ---
        report_title = f"{self.detected_category}ウィークリーレポート"
        self._log_safe(f"  投稿を検索: 「{report_title}」")
        post = self.wp_client.search_post_by_title(
            report_title, log_func=self._log_safe
        )

        if post:
            post_id = post["id"]
            try:
                existing_content = self.wp_client.get_post_raw_content(post_id)
            except Exception as e:
                self._log_safe(f"  ⚠ rawコンテンツ取得失敗、空として続行: {e}")
                existing_content = ""
            self._log_safe(f"  既存コンテンツ取得完了 ({len(existing_content)}文字)")

            # --- 3. 既存コンテンツを解析して再構成 ---
            new_content = self._restructure_content(
                existing_content, new_blocks
            )

            self.wp_client.update_post(
                post_id, new_content, bump_to_top=True
            )
            self._log_safe(
                f"  ✓ 投稿を更新・トップに移動しました (ID: {post_id})"
            )
        else:
            self._log_safe(
                f"  投稿「{report_title}」が見つかりません。新規作成します。"
            )
            try:
                cat_id = self.wp_client.get_or_create_category(
                    self.detected_category
                )
                content = "\n".join(new_blocks)
                new_post = self.wp_client.create_post(
                    title=report_title,
                    content=content,
                    category_ids=[cat_id] if cat_id else None,
                    status="draft",
                )
                new_id = new_post.get("id", "?")
                self._log_safe(
                    f"  ✓ 新規投稿を下書きで作成しました (ID: {new_id})"
                )
            except Exception as e:
                self._log_safe(f"  ✗ 新規投稿の作成に失敗: {e}")

    def _restructure_content(
        self, existing_content: str, new_blocks: list[str]
    ) -> str:
        """投稿コンテンツを再構成する。

        - 既存の「今週のレポート」セクション（アコーディオン外のブロック）を
          「過去のレポート」アコーディオン内に移動
        - 新しいブロックを先頭に配置
        - 既存の「過去のレポート」アコーディオンはそのまま維持（中に追記）
        """
        ARCHIVE_SUMMARY = "過去のレポート・ポッドキャスト"

        # 既存コンテンツから各パートを分離
        # 1. 「過去のレポート」アコーディオンの中身を抽出
        archive_pattern = (
            r'<!-- wp:details -->\s*'
            r'<details class="wp-block-details">\s*'
            r'<summary>' + re.escape(ARCHIVE_SUMMARY) + r'</summary>\s*'
            r'(.*?)'
            r'</details>\s*'
            r'<!-- /wp:details -->'
        )
        archive_match = re.search(archive_pattern, existing_content, re.DOTALL)

        if archive_match:
            # 既存のアーカイブアコーディオンの中身
            archive_inner = archive_match.group(1).strip()
            # アーカイブ以外のコンテンツ（＝前回の「最新」部分）
            non_archive = (
                existing_content[:archive_match.start()].strip()
                + "\n"
                + existing_content[archive_match.end():].strip()
            ).strip()
        else:
            archive_inner = ""
            non_archive = existing_content.strip()

        # 2. 前回の「最新」部分をアコーディオンアイテムに変換
        #    （空のパラグラフや空のSwell accordion は除外）
        old_latest_items = self._extract_file_blocks(non_archive)

        # 3. アーカイブに追記する中身を構築
        #    前回の最新 → 個別アコーディオンに変換してアーカイブに追加
        new_archive_items = ""
        for item_label, item_content in old_latest_items:
            new_archive_items += (
                f'<!-- wp:details -->\n'
                f'<details class="wp-block-details">'
                f'<summary>{item_label}</summary>\n'
                f'{item_content}'
                f'</details>\n'
                f'<!-- /wp:details -->\n'
            )

        # アーカイブ全体: 新しい→古い の順
        full_archive_inner = new_archive_items + archive_inner

        # 4. 最終コンテンツを組み立て
        #    最新ブロック + 過去アコーディオン
        parts = []

        # 最新のPDF/MP3ブロック（直接表示）
        for block in new_blocks:
            parts.append(block)

        # 過去のレポートアコーディオン（中身がある場合のみ）
        if full_archive_inner.strip():
            parts.append(
                f'<!-- wp:details -->\n'
                f'<details class="wp-block-details">'
                f'<summary>{ARCHIVE_SUMMARY}</summary>\n'
                f'{full_archive_inner}'
                f'</details>\n'
                f'<!-- /wp:details -->\n'
            )

        return "\n".join(parts)

    def _extract_file_blocks(
        self, content: str
    ) -> list[tuple[str, str]]:
        """コンテンツからファイルブロック（PDF/MP3）を抽出する。

        Returns:
            [(ラベル, ブロックHTML), ...] のリスト
        """
        items = []
        if not content.strip():
            return items

        # wp:details ブロックを抽出（既存のアコーディオン項目）
        details_pattern = (
            r'<!-- wp:details -->\s*'
            r'<details class="wp-block-details">\s*'
            r'<summary>(.*?)</summary>\s*'
            r'(.*?)'
            r'</details>\s*'
            r'<!-- /wp:details -->'
        )
        for m in re.finditer(details_pattern, content, re.DOTALL):
            label = m.group(1).strip()
            inner = m.group(2).strip()
            if inner and label:  # 空でなければ追加
                items.append((label, inner))

        # wp:file ブロックをアコーディオン外から直接抽出
        # (前回の最新表示分: paragraphヘッダー + file/audioブロック)
        remaining = re.sub(details_pattern, '', content, flags=re.DOTALL)

        # 📄/🎙 ヘッダー付きのファイルブロック群を抽出
        header_pattern = (
            r'<!-- wp:paragraph -->\s*'
            r'<p><strong>[📄🎙]\s*(.*?)</strong></p>\s*'
            r'<!-- /wp:paragraph -->'
        )
        headers = list(re.finditer(header_pattern, remaining, re.DOTALL))

        for hi, hm in enumerate(headers):
            label = hm.group(1).strip()
            start = hm.end()
            # 次のヘッダーまで、またはコンテンツ末尾
            end = headers[hi + 1].start() if hi + 1 < len(headers) else len(remaining)
            block_content = remaining[start:end].strip()

            # 空でないfile/audioブロックが含まれていれば追加
            if '<!-- wp:file' in block_content or '<!-- wp:audio' in block_content:
                items.append((label, block_content))

        return items

    def _make_file_display_label(self, filename: str) -> str:
        """ファイル名から表示ラベルを生成する

        例: 全固体電池調査ウィークリーレポート20260322.pdf
          → ウィークリーレポート 2026年3月22日（PDF）をダウンロード
        例: 全固体電池調査20260322ポッドキャスト.mp3
          → ポッドキャスト 2026年3月22日（MP3）を再生
        """
        ext = Path(filename).suffix.lower()
        is_mp3 = ext == ".mp3"

        # ファイル名から日付部分(YYYYMMDD)を抽出
        match = re.search(r"(\d{4})(\d{2})(\d{2})", filename)
        if match:
            year = match.group(1)
            month = str(int(match.group(2)))  # 先頭ゼロ除去
            day = str(int(match.group(3)))
            date_str = f"{year}年{month}月{day}日"
        else:
            now = datetime.now()
            date_str = f"{now.year}年{now.month}月{now.day}日"

        if is_mp3:
            return f"ポッドキャスト {date_str}（MP3）を再生"
        else:
            return f"ウィークリーレポート {date_str}（PDF）をダウンロード"

    def _finish_all(self):
        """全処理完了"""
        self._log("=" * 40)
        self._log("すべての処理が完了しました")
        self.title_label.config(text="すべての処理が完了しました")
        messagebox.showinfo("完了", "すべての処理が完了しました")

    # ---------------------------------------------------------------
    # ヘルパー
    # ---------------------------------------------------------------
    def _next_article(self):
        self._close_chrome()
        self.current_index += 1
        self._show_current_article()

    def _skip_all_articles(self):
        """すべての記事をスキップしてPDF/ポッドキャストアップロードへ直行"""
        self._close_chrome()
        remaining = len(self.html_files) - self.current_index
        if remaining <= 0:
            return
        proceed = messagebox.askyesno(
            "すべてスキップ",
            f"残り{remaining}件の記事をすべてスキップして、\n"
            f"PDF・ポッドキャストのアップロードに進みますか？",
        )
        if not proceed:
            return
        self._log(f"→ 残り{remaining}件の記事をすべてスキップ")
        self.current_index = len(self.html_files)
        self.progress_label.config(text="完了")
        self.title_label.config(text="記事をスキップしました")
        self.meta_label.config(text="")
        self.complete_btn.config(state="disabled")
        self.skip_btn.config(state="disabled")
        self.skip_all_btn.config(state="disabled")
        # PDF/ポッドキャストアップロード処理へ
        self._handle_pdf_upload()

    def _enable_buttons(self):
        self.complete_btn.config(state="normal")
        self.skip_btn.config(state="normal")
        self.skip_all_btn.config(state="normal")

    def _on_error(self, message: str):
        self._log(f"✗ エラー: {message}")
        messagebox.showerror("エラー", message)
        self._enable_buttons()

    def _log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")

    def _log_safe(self, message: str):
        """別スレッドからのログ追加"""
        self.after(0, lambda: self._log(message))


# =====================================================================
# エントリーポイント
# =====================================================================
def main():
    app = WPUploaderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
