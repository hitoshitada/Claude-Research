"""WordPress記事自動アップロードツール

調査アウトプットフォルダのHTMLファイルをWordPressに自動投稿する。
- フォルダ選択 → Chrome表示 → 閲覧完了 → アップロード判定
- article_curator.py で確定した _eyecatch.png/jpg をアイキャッチとして使用
- 不採用_で始まるファイルは自動スキップ
- 概要・詳細を緑帯ヘッダー付きで投稿
- sys.argv[1] でフォルダを指定可能（pipeline_launcher.py 連携）
"""

import sys
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
# 記事内容カテゴリー分類（企業動向 / 市場動向 / 新技術・技術紹介）
# =====================================================================
_CONTENT_CATEGORY_RULES: list[tuple[str, list[str]]] = [
    ("企業動向", [
        # 企業・組織名
        "企業", "会社", "メーカー", "スタートアップ", "グループ",
        # 事業活動
        "製品発表", "新製品", "量産", "量産化", "パイロット", "パイロットライン",
        "工場", "生産ライン", "生産拠点", "稼働", "稼働開始",
        # 投資・資本
        "投資", "資金調達", "融資", "出資",
        # 提携・M&A
        "提携", "買収", "合弁", "パートナーシップ", "共同開発", "協業",
        # 戦略・計画
        "戦略", "ロードマップ", "計画", "目標", "方針",
        # 上場・業績
        "上場", "株式", "売上", "収益", "業績", "利益",
        # 商業展開
        "商業化", "商業生産", "サンプル出荷", "サンプル提供", "販売開始",
    ]),
    ("市場動向", [
        # 市場・需給
        "市場", "需要", "供給", "価格", "コスト",
        # 成長・予測
        "成長率", "CAGR", "予測", "見通し", "展望", "予想",
        # 調査・レポート
        "調査", "調査レポート", "市場調査", "リサーチ",
        # 競争・シェア
        "シェア", "競争", "競合", "市場占有率",
        # 業界・トレンド
        "業界", "トレンド", "動向", "普及",
        # 貿易・地域
        "輸出", "輸入", "貿易", "サプライチェーン",
        "グローバル", "アジア", "欧州", "北米", "中国市場", "米国市場",
        # 規制・政策
        "規制", "政策", "補助金", "政府", "法規制", "標準化", "規格",
        # 金額・規模
        "億ドル", "兆円", "億円", "市場規模",
    ]),
    ("新技術・技術紹介", [
        # 研究・学術
        "研究", "論文", "学術", "大学", "研究所", "研究機関", "教授", "博士",
        # 発見・開発
        "発見", "新手法", "ブレークスルー", "新技術", "革新",
        # 性能・特性
        "性能向上", "特性改善", "高性能", "高効率",
        # 技術解析
        "メカニズム", "解析", "分析", "評価", "実証",
        # 製造・合成
        "合成", "製法", "プロセス", "構造", "微細構造",
        # 試作・実験
        "プロトタイプ", "試作", "実験", "検証", "テスト",
        # 特許・イノベーション
        "特許", "イノベーション", "先端技術", "次世代技術",
        # 技術分野キーワード
        "ナノ", "量子", "固体電解質", "電解質", "正極", "負極",
    ]),
]


def detect_content_category(article: dict) -> str | None:
    """記事のタイトル・概要・詳細テキストをスコアリングして
    「企業動向」「市場動向」「新技術・技術紹介」の最も適切なカテゴリーを返す。
    スコアが同点の場合は None を返す（判定不能）。
    """
    # HTML除去してテキスト取得（タイトル2倍重み付き）
    title_text = article.get("title", "")
    summary_text = BeautifulSoup(article.get("summary", ""), "html.parser").get_text()
    detail_text  = BeautifulSoup(article.get("detail",  ""), "html.parser").get_text()
    full_text = title_text * 2 + " " + summary_text + " " + detail_text

    scores: dict[str, int] = {}
    for category_name, keywords in _CONTENT_CATEGORY_RULES:
        score = sum(full_text.count(kw) for kw in keywords)
        scores[category_name] = score

    # 最高スコアのカテゴリーを返す（0点なら None）
    best_cat = max(scores, key=lambda c: scores[c])
    if scores[best_cat] == 0:
        return None
    return best_cat


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

    def get_or_create_category(
        self, category_name: str, parent_id: int | None = None
    ) -> int | None:
        """カテゴリーをWordPressから検索し、なければ新規作成してIDを返す。

        parent_id が指定された場合は、そのカテゴリーの子として検索・作成する。
        これにより「半導体後工程 > 企業動向」「全固体電池 > 市場動向」など
        親カテゴリーに紐付いた子カテゴリーが正しく設定される。
        """
        # 検索パラメータ（parent指定で子カテゴリーに絞り込む）
        params: dict = {"search": category_name, "per_page": 100}
        if parent_id is not None:
            params["parent"] = parent_id

        resp = requests.get(
            f"{self.api_url}/categories",
            auth=self.auth,
            params=params,
            timeout=10,
        )
        if resp.status_code == 200:
            for cat in resp.json():
                if cat.get("name") == category_name:
                    # parent_id 指定時は親が一致するものだけ返す
                    if parent_id is None or cat.get("parent") == parent_id:
                        return cat["id"]

        # 見つからなければ新規作成
        create_data: dict = {"name": category_name}
        if parent_id is not None:
            create_data["parent"] = parent_id

        resp = requests.post(
            f"{self.api_url}/categories",
            auth=self.auth,
            json=create_data,
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
        """既存投稿のコンテンツを更新する。

        コンテンツ更新と日付/sticky 更新を**別々のAPIコール**で行う。
        同一リクエストに混在させると WordPress の content 保存フックが
        date を元に戻すことがあるため、必ず2ステップで実行する。
        """
        # ── Step 1: コンテンツを publish で更新 ──
        resp = requests.post(
            f"{self.api_url}/posts/{post_id}",
            auth=self.auth,
            json={"content": content, "status": "publish"},
            timeout=60,
        )
        if resp.status_code not in (200, 201):
            raise Exception(
                f"投稿コンテンツ更新失敗: {resp.status_code} {resp.text[:300]}"
            )
        result = resp.json()

        # ── Step 2: 日付・sticky を別コールで更新 ──
        if bump_to_top:
            result = self.bump_to_top(post_id)

        return result

    def bump_to_top(self, post_id: int) -> dict:
        """投稿日時を現在時刻に更新し先頭に移動させる（単独APIコール）。

        date（ローカル JST）と date_gmt（UTC）の両方を明示的に設定し、
        sticky=True で公開状態にする。
        """
        from datetime import timezone as _tz, timedelta as _td
        now_utc = datetime.now(_tz.utc)
        now_utc_str = now_utc.strftime("%Y-%m-%dT%H:%M:%S")
        # WordPress の date フィールドはサイトのローカル日時。
        # 日本語サイト = JST (UTC+9) に合わせて明示的に設定する。
        now_jst = now_utc + _td(hours=9)
        now_jst_str = now_jst.strftime("%Y-%m-%dT%H:%M:%S")

        payload = {
            "date":     now_jst_str,   # ローカル日時（JST）
            "date_gmt": now_utc_str,   # UTC 日時
            "status":   "publish",
            "sticky":   True,
        }
        resp = requests.post(
            f"{self.api_url}/posts/{post_id}",
            auth=self.auth,
            json=payload,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            return resp.json()
        raise Exception(
            f"投稿トップ移動失敗: {resp.status_code} {resp.text[:300]}"
        )

    def search_media_by_keyword(self, keyword: str) -> int | None:
        """メディアライブラリからキーワードに一致する画像のIDを返す。

        WP REST API の /wp-json/wp/v2/media?search=keyword でタイトル・
        alt text・ファイル名を横断検索し、最初にヒットした画像の ID を返す。
        見つからない場合は None を返す。
        """
        try:
            resp = requests.get(
                f"{self.api_url}/media",
                auth=self.auth,
                params={"search": keyword, "media_type": "image", "per_page": 10},
                timeout=15,
            )
            if resp.status_code != 200:
                return None
            items = resp.json()
            if items:
                return items[0]["id"]
        except Exception:
            pass
        return None

    def create_post(
        self,
        title: str,
        content: str,
        featured_media_id: int | None = None,
        category_ids: list[int] | None = None,
        status: str = "draft",
        sticky: bool = False,
        date_gmt: str | None = None,
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
        if sticky:
            data["sticky"] = True
        if date_gmt:
            data["date_gmt"] = date_gmt
            # date（ローカル JST）も明示設定して WordPress が正しく日付を解釈できるようにする
            from datetime import timezone as _tz, timedelta as _td
            try:
                from datetime import datetime as _dt
                utc_dt = _dt.strptime(date_gmt, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=_tz.utc)
                jst_dt = utc_dt + _td(hours=9)
                data["date"] = jst_dt.strftime("%Y-%m-%dT%H:%M:%S")
            except Exception:
                pass  # date_gmt パース失敗時は date なし（WordPress が自動補完）

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

    # 概要からHTMLタグを除去
    clean_summary = BeautifulSoup(summary_text, "html.parser").get_text()

    # ── タイトルから「文として読める英語の長文」要素を視覚的なキーワードに変換 ──
    # 数値・金額・パーセント・年号などを除去（画像内テキストの原因になる）
    import re as _re
    title_visual = _re.sub(r'\$[\d,\.]+\s*[BbMmKk](?:illion|illion)?', '', title)
    title_visual = _re.sub(r'[\d,\.]+\s*%', '', title_visual)
    title_visual = _re.sub(r'\b20\d\d\b', '', title_visual)        # 西暦年
    title_visual = _re.sub(r'\b(?:is set to|will reach|reach|in|by|for|of|the)\b',
                           '', title_visual, flags=_re.IGNORECASE)
    title_visual = _re.sub(r'\s{2,}', ' ', title_visual).strip(' ,.')

    # 概要から視覚的イメージに使えるキーワードだけ抽出（先頭100字）
    summary_short = _re.sub(r'\s+', ' ', clean_summary).strip()[:100]

    # ── プロンプト設計のポイント ──
    # 1) 「NO TEXT」を冒頭に置く（Imagenは先頭の指示を優先）
    # 2) "Topic:" "Context:" のようなラベル表記を使わない（そのまま画像に描かれる）
    # 3) 視覚的なシーン描写のみにする
    # 4) プロンプト内の "NO TEXT" 表現でテキスト要素を排除（negative_prompt は API 非対応）
    prompt = (
        "NO TEXT. NO WORDS. NO LETTERS. NO NUMBERS. ZERO TEXT ELEMENTS. "
        "A professional photorealistic wide-format image for a technology news blog. "
        "Purely visual scene — no written characters, no captions, no labels, "
        "no watermarks, no overlay text of any kind whatsoever. "
        f"Visually depict the scene related to: {title_visual}. "
        f"Visual atmosphere inspired by: {summary_short}. "
        "Composition: wide 16:9, cinematic lighting, modern high-tech environment, "
        "sharp focus, detailed. Completely text-free image."
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
# グラフ検出・再生成モジュール
# =====================================================================

def _img_mime(image_data: bytes) -> str:
    """バイト列からMIMEタイプを簡易判定する"""
    if image_data[:4] == b'\x89PNG':
        return "image/png"
    if image_data[:4] in (b'RIFF',) or image_data[8:12] == b'WEBP':
        return "image/webp"
    return "image/jpeg"


def detect_chart_in_image(image_data: bytes, api_key: str) -> bool:
    """Gemini Visionで画像内にグラフ/チャートが含まれるか検出する"""
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=[
                types.Part.from_bytes(data=image_data, mime_type=_img_mime(image_data)),
                types.Part.from_text(
                    "Does this image contain any data visualization chart or graph "
                    "(bar chart, line chart, pie chart, scatter plot, area chart, etc.)? "
                    "Answer only 'yes' or 'no'."
                ),
            ],
        )
        return "yes" in response.text.lower()
    except Exception:
        return False


def extract_chart_data(image_data: bytes, api_key: str) -> dict | None:
    """Gemini Visionでグラフデータを JSON 形式で抽出する。

    Returns:
        {
          "chart_type": "bar"|"line"|"pie"|"area"|"scatter"|"other",
          "title": str, "x_label": str, "y_label": str, "unit": str,
          "series": [{"name": str, "data": [{"label": str, "value": float}, ...]}]
        }
        抽出失敗時は None。
    """
    # 試行するモデルのリスト（上から順に試す）
    _CHART_MODELS = [
        "gemini-2.5-flash-preview-04-17",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
    ]
    import json as _json

    last_err = None
    for _model in _CHART_MODELS:
        try:
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model=_model,
                contents=[
                    types.Part.from_bytes(data=image_data, mime_type=_img_mime(image_data)),
                    types.Part.from_text(
                        "This image contains one or more charts/graphs (may be a compound infographic).\n"
                        "Focus on the LARGEST or MOST PROMINENT chart in the image.\n"
                        "Extract all numerical data from that chart.\n"
                        "Return ONLY a valid JSON object with this structure:\n"
                        '{"chart_type":"bar|line|pie|area|scatter|other",'
                        '"title":"chart title or empty","x_label":"","y_label":"",'
                        '"unit":"unit of values (%, billion USD, etc.) or empty",'
                        '"series":[{"name":"series name or empty",'
                        '"data":[{"label":"category or x-value","value":123.4}]}]}'
                        "\nBe precise with all numeric values. No markdown, no explanation.\n"
                        "If the image has NO chart at all, return: "
                        '{"chart_type":"other","title":"","x_label":"","y_label":"","unit":"","series":[]}'
                    ),
                ],
                config={"response_mime_type": "application/json"},
            )
            text = re.sub(r'^```(?:json)?\s*', '', response.text.strip())
            text = re.sub(r'\s*```$', '', text)
            result = _json.loads(text)
            return result
        except Exception as e:
            last_err = e
            continue  # 次のモデルを試す

    # 全モデルが失敗した場合は最後のエラーを付けて None を返す
    raise RuntimeError(f"extract_chart_data failed with all models. Last error: {last_err}")


def render_chart_image(chart_data: dict,
                        width: int = 900, height: int = 500) -> "PILImage.Image | None":
    """matplotlibでグラフを描画し PIL Image を返す（著作権フリーの新グラフ）"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import numpy as np
        import io as _io

        chart_type = chart_data.get("chart_type", "bar").lower()
        title      = chart_data.get("title", "")
        x_label    = chart_data.get("x_label", "")
        y_label    = chart_data.get("y_label", "")
        unit       = chart_data.get("unit", "")
        series     = chart_data.get("series", [])

        if not series or not series[0].get("data"):
            return None

        dpi = 100
        fig, ax = plt.subplots(figsize=(width / dpi, height / dpi))
        fig.patch.set_facecolor('#0d1b2a')
        ax.set_facecolor('#1b2838')

        palette = ['#4fc3f7', '#81c784', '#ffb74d', '#f48fb1',
                   '#ce93d8', '#80deea', '#a5d6a7', '#fff176']

        def _vals(s):
            return [float(d.get("value", 0)) for d in s.get("data", [])]

        def _labs(s):
            return [str(d.get("label", "")) for d in s.get("data", [])]

        if chart_type in ("bar",):
            labels = _labs(series[0])
            if len(series) > 1:
                n = len(labels)
                bw = 0.8 / len(series)
                x = np.arange(n)
                for si, s in enumerate(series):
                    ax.bar(x + si * bw, _vals(s), bw,
                           label=s.get("name", ""),
                           color=palette[si % len(palette)],
                           alpha=0.85, edgecolor="#ffffff20")
                ax.set_xticks(x + bw * (len(series) - 1) / 2)
                ax.set_xticklabels(labels, rotation=20, ha="right",
                                   fontsize=8, color="#cccccc")
                ax.legend(fontsize=8, facecolor="#1b2838", labelcolor="white")
            else:
                vals = _vals(series[0])
                bars = ax.bar(range(len(labels)), vals,
                              color=palette[:len(labels)], alpha=0.85,
                              edgecolor="#ffffff20")
                for bar, v in zip(bars, vals):
                    ax.text(bar.get_x() + bar.get_width() / 2,
                            bar.get_height(), f"{v:g}",
                            ha="center", va="bottom", fontsize=7, color="#e0e0e0")
                ax.set_xticks(range(len(labels)))
                ax.set_xticklabels(labels, rotation=20, ha="right",
                                   fontsize=8, color="#cccccc")

        elif chart_type in ("line", "area"):
            for si, s in enumerate(series):
                labs = _labs(s)
                vals = _vals(s)
                col  = palette[si % len(palette)]
                ax.plot(range(len(labs)), vals, color=col, linewidth=2.5,
                        marker="o", markersize=5, label=s.get("name", ""), zorder=3)
                if chart_type == "area":
                    ax.fill_between(range(len(labs)), vals, alpha=0.2, color=col)
            if len(series) > 1:
                ax.legend(fontsize=8, facecolor="#1b2838", labelcolor="white")
            labs_first = _labs(series[0])
            ax.set_xticks(range(len(labs_first)))
            ax.set_xticklabels(labs_first, rotation=20, ha="right",
                               fontsize=8, color="#cccccc")

        elif chart_type == "pie":
            first = series[0].get("data", [])
            pl = [str(d.get("label", "")) for d in first]
            pv = [max(float(d.get("value", 0)), 0) for d in first]
            wedges, texts, autotexts = ax.pie(
                pv, labels=pl, colors=palette[:len(pl)],
                autopct="%1.1f%%", startangle=90,
                textprops={"color": "#e0e0e0", "fontsize": 8},
                wedgeprops={"linewidth": 0.5, "edgecolor": "#0d1b2a"},
            )
            for at in autotexts:
                at.set_color("#ffffff"); at.set_fontsize(7)
        else:
            # フォールバック: 棒グラフ
            labs = _labs(series[0])
            vals = _vals(series[0])
            ax.bar(range(len(labs)), vals, color=palette[:len(labs)], alpha=0.85)
            ax.set_xticks(range(len(labs)))
            ax.set_xticklabels(labs, rotation=20, ha="right", fontsize=8, color="#cccccc")

        # 軸・グリッド装飾（円グラフ以外）
        if chart_type != "pie":
            ax.tick_params(axis="y", colors="#aaaaaa", labelsize=8)
            ax.tick_params(axis="x", colors="#aaaaaa")
            for spine in ("bottom", "left"):
                ax.spines[spine].set_color("#444444")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.yaxis.grid(True, alpha=0.2, color="#888888", linestyle="--")
            ax.set_axisbelow(True)
            full_y = f"{y_label} ({unit})" if unit else y_label
            if full_y:
                ax.set_ylabel(full_y, color="#bbbbbb", fontsize=8)
            if x_label:
                ax.set_xlabel(x_label, color="#bbbbbb", fontsize=8)

        if title:
            ax.set_title(title, color="#e0e0e0", fontsize=10,
                         fontweight="bold", pad=10)

        plt.tight_layout(pad=1.2)
        buf = _io.BytesIO()
        plt.savefig(buf, format="png", dpi=dpi,
                    facecolor=fig.get_facecolor(), bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return PILImage.open(buf).copy()

    except Exception:
        return None


def generate_image_with_chart(
    api_key: str,
    title: str,
    summary_text: str,
    chart_data: dict,
    output_path: Path,
) -> Path:
    """Imagen背景 + matplotlibグラフを合成したアイキャッチ画像を生成する。

    1. Imagen で右下スペースを確保した背景を生成
    2. chart_data から matplotlib で新グラフを描画
    3. 背景の右下にグラフを合成して 1920×1080 で保存
    失敗時は通常の generate_image_with_gemini にフォールバック。
    """
    client = genai.Client(api_key=api_key)

    # タイトルから数値・金額などを除去してビジュアル用に整形
    import re as _re
    title_visual = _re.sub(r'\$[\d,\.]+\s*[BbMm](?:illion)?', '', title)
    title_visual = _re.sub(r'[\d,\.]+\s*%', '', title_visual)
    title_visual = _re.sub(r'\b20\d\d\b', '', title_visual)
    title_visual = _re.sub(r'\s{2,}', ' ', title_visual).strip(' ,.')
    summary_short = _re.sub(r'\s+', ' ',
        BeautifulSoup(summary_text, "html.parser").get_text()).strip()[:100]

    # 右下を暗く空けたロンプト（グラフ配置スペース確保）
    prompt = (
        "NO TEXT. NO WORDS. NO LETTERS. NO NUMBERS. ZERO TEXT ELEMENTS. "
        "Professional wide-format technology background image for a data-driven news article. "
        "The lower-right quarter of the image should be relatively dark and clear "
        "— reserved for a data chart that will be overlaid. "
        f"Visual theme: {title_visual}. "
        f"Atmosphere: {summary_short}. "
        "Style: cinematic dark blue tech, modern, 16:9. Completely text-free."
    )
    # Step1: 背景画像生成
    resp = client.models.generate_images(
        model=IMAGE_MODEL,
        prompt=prompt,
        config=types.GenerateImagesConfig(
            number_of_images=1,
            aspect_ratio="16:9",
        ),
    )
    if not resp.generated_images:
        # 背景生成失敗 → 通常生成にフォールバック
        return generate_image_with_gemini(api_key, title, summary_text, output_path)

    bg_pil = PILImage.open(
        io.BytesIO(resp.generated_images[0].image.image_bytes)
    ).resize((TARGET_WIDTH, TARGET_HEIGHT), PILImage.LANCZOS).convert("RGBA")

    # Step2: matplotlibでグラフ描画
    chart_pil = render_chart_image(chart_data, width=860, height=480)
    if chart_pil is None:
        # グラフ描画失敗 → 背景だけ保存
        bg_pil.convert("RGB").save(str(output_path), "PNG")
        return output_path

    # Step3: 右下にグラフを合成
    chart_pil = chart_pil.convert("RGBA")
    margin = 36
    cw = int(TARGET_WIDTH * 0.48)
    ch = int(cw * chart_pil.height / chart_pil.width)
    chart_resized = chart_pil.resize((cw, ch), PILImage.LANCZOS)

    cx = TARGET_WIDTH - cw - margin
    cy = TARGET_HEIGHT - ch - margin

    # グラフの背後に半透明パネル
    panel = PILImage.new("RGBA", (cw + 16, ch + 16), (0, 0, 0, 190))
    bg_pil.alpha_composite(panel, (cx - 8, cy - 8))
    bg_pil.alpha_composite(chart_resized, (cx, cy))

    bg_pil.convert("RGB").save(str(output_path), "PNG")
    return output_path




# =====================================================================
# メインGUIアプリケーション（新版）
# =====================================================================
class WPUploaderApp(tk.Tk):
    """WordPress アップローダー（新版）

    3つの独立したアップロードボタンを提供:
    1. 採用記事を一括アップロード（Chromeプレビューなし・確認なし）
    2. WeeklyReport + ポッドキャストをアップロード（新規投稿・先頭移動）
    3. 全てまとめてアップロード（1→2を順に実行）
    """

    def __init__(self):
        super().__init__()
        self.title("WordPress アップローダー")
        self.geometry("960x720")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        # 状態
        self.wp_client: WordPressClient | None = None
        self.selected_folder: Path | None = None
        self.detected_category: str | None = None
        self._uploading = False

        self._build_ui()
        self._init_wp_client()

    # ---------------------------------------------------------------
    # UI構築
    # ---------------------------------------------------------------
    def _build_ui(self):
        # --- WordPress設定 ---
        wp_frame = ttk.LabelFrame(self, text="WordPress設定", padding=8)
        wp_frame.pack(fill="x", padx=10, pady=(10, 4))

        row0 = ttk.Frame(wp_frame)
        row0.pack(fill="x")
        ttk.Label(row0, text="サイトURL:").pack(side="left")
        self.wp_url_var = tk.StringVar(value=WP_URL)
        ttk.Entry(row0, textvariable=self.wp_url_var, width=35).pack(
            side="left", padx=(4, 12))
        ttk.Label(row0, text="ユーザー名:").pack(side="left")
        self.wp_user_var = tk.StringVar(value=WP_USERNAME)
        ttk.Entry(row0, textvariable=self.wp_user_var, width=18).pack(
            side="left", padx=(4, 0))

        row1 = ttk.Frame(wp_frame)
        row1.pack(fill="x", pady=(4, 0))
        ttk.Label(row1, text="アプリパスワード:").pack(side="left")
        self.wp_pass_var = tk.StringVar(value=WP_APP_PASSWORD)
        ttk.Entry(row1, textvariable=self.wp_pass_var, width=35, show="*").pack(
            side="left", padx=(4, 12))
        ttk.Button(row1, text="接続テスト",
                   command=self._test_wp_connection).pack(side="left")
        self.wp_status_label = ttk.Label(row1, text="", foreground="gray")
        self.wp_status_label.pack(side="left", padx=8)

        # --- フォルダ選択 ---
        folder_frame = ttk.LabelFrame(self, text="対象フォルダ", padding=8)
        folder_frame.pack(fill="x", padx=10, pady=4)

        ttk.Button(folder_frame, text="フォルダを選択",
                   command=self._select_folder).pack(side="left")
        self.folder_label = ttk.Label(
            folder_frame, text="未選択", foreground="gray")
        self.folder_label.pack(side="left", padx=10)
        self.folder_info_label = ttk.Label(
            folder_frame, text="", foreground="#1565C0")
        self.folder_info_label.pack(side="right")

        # --- アップロードアクション ---
        action_frame = ttk.LabelFrame(self, text="アップロード操作", padding=10)
        action_frame.pack(fill="x", padx=10, pady=4)

        btn_row = tk.Frame(action_frame, bg="#f0f0f0")
        btn_row.pack(fill="x")

        self.btn_articles = tk.Button(
            btn_row,
            text="📰 採用記事を\n一括アップロード",
            command=self._upload_articles_batch,
            state="disabled",
            bg="#2E7D32", fg="white",
            font=("Meiryo", 11, "bold"),
            relief="raised", padx=12, pady=10,
        )
        self.btn_articles.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_weekly = tk.Button(
            btn_row,
            text="📊 WeeklyReport を\nアップロード",
            command=self._upload_weekly_report,
            state="disabled",
            bg="#1565C0", fg="white",
            font=("Meiryo", 11, "bold"),
            relief="raised", padx=12, pady=10,
        )
        self.btn_weekly.pack(side="left", fill="x", expand=True, padx=4)

        self.btn_all = tk.Button(
            btn_row,
            text="🚀 全てまとめて\nアップロード",
            command=self._upload_all,
            state="disabled",
            bg="#E65100", fg="white",
            font=("Meiryo", 11, "bold"),
            relief="raised", padx=12, pady=10,
        )
        self.btn_all.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # ステータスと進行バー
        status_frame = tk.Frame(action_frame, bg="#f0f0f0")
        status_frame.pack(fill="x", pady=(8, 0))
        self.status_var = tk.StringVar(value="フォルダを選択してください")
        ttk.Label(status_frame, textvariable=self.status_var,
                  font=("Meiryo", 9)).pack(anchor="w")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            status_frame, variable=self.progress_var,
            maximum=100, mode="determinate",
        )
        self.progress_bar.pack(fill="x", pady=(4, 0))

        # --- ログ ---
        log_frame = ttk.LabelFrame(self, text="ログ", padding=4)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=16, font=("Consolas", 9), wrap="word"
        )
        self.log_text.pack(fill="both", expand=True)

    # ---------------------------------------------------------------
    # WordPress接続
    # ---------------------------------------------------------------
    def _init_wp_client(self):
        url  = self.wp_url_var.get().strip()
        user = self.wp_user_var.get().strip()
        pw   = self.wp_pass_var.get().strip()
        if url and user and pw:
            self.wp_client = WordPressClient(url, user, pw)

    def _build_wp_client(self) -> bool:
        url  = self.wp_url_var.get().strip()
        user = self.wp_user_var.get().strip()
        pw   = self.wp_pass_var.get().strip()
        if not (url and user and pw):
            messagebox.showerror(
                "エラー",
                "WordPress設定（URL・ユーザー名・アプリパスワード）をすべて入力してください")
            return False
        self.wp_client = WordPressClient(url, user, pw)
        return True

    def _test_wp_connection(self):
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
            messagebox.showerror("接続失敗",
                                 f"WordPressに接続できませんでした\n{error}")

    # ---------------------------------------------------------------
    # フォルダ選択
    # ---------------------------------------------------------------
    def _select_folder(self):
        folder = filedialog.askdirectory(
            initialdir=str(OUTPUT_DIR),
            title="調査アウトプットフォルダを選択",
        )
        if folder:
            self._load_folder_path(Path(folder))

    def _load_folder_path(self, folder_path: Path):
        self.selected_folder = folder_path
        self.folder_label.config(text=folder_path.name, foreground="black")

        # 採用記事HTMLを収集（不採用_・_概要 を除外）
        html_files = sorted(
            f for f in folder_path.glob("*.html")
            if not f.name.startswith("不採用_")
            and not f.name.endswith("_概要.html")
        )
        # アイキャッチ画像があるものだけをアップロード対象としてカウント
        articles_with_eyecatch = [
            f for f in html_files if self._find_eyecatch_file(f)]

        # PDF・MP3を確認
        pdf_files = [f for f in folder_path.glob("*.pdf")
                     if "ウィークリーレポート" in f.name]
        mp3_files = [f for f in folder_path.glob("*.mp3")
                     if "ポッドキャスト" in f.name
                     and not f.stem.endswith("_original")]

        # カテゴリー自動判定
        self.detected_category = detect_category(folder_path.name)
        cat_msg = self.detected_category or "（カテゴリー未判定）"

        info_parts = [f"記事: {len(articles_with_eyecatch)}件"]
        if pdf_files:
            info_parts.append(f"PDF: {len(pdf_files)}件")
        if mp3_files:
            info_parts.append(f"ポッドキャスト: {len(mp3_files)}件")
        info_parts.append(cat_msg)
        self.folder_info_label.config(text=" | ".join(info_parts))

        self._log(f"フォルダ: {folder_path.name}")
        self._log(f"  採用記事（アイキャッチあり）: {len(articles_with_eyecatch)}件")
        self._log(f"  PDF: {len(pdf_files)}件 / MP3: {len(mp3_files)}件")
        self._log(f"  カテゴリー: {cat_msg}")

        # ボタン有効化
        self.btn_articles.config(state="normal")
        self.btn_weekly.config(state="normal")
        self.btn_all.config(state="normal")
        self.status_var.set("アップロードボタンを押してください")
        self.progress_var.set(0)

    def _find_eyecatch_file(self, html_path: Path) -> "Path | None":
        """対応する _eyecatch.png/jpg/jpeg を返す（なければ None）"""
        for ext in [".png", ".jpg", ".jpeg"]:
            c = html_path.parent / f"{html_path.stem}_eyecatch{ext}"
            if c.exists():
                return c
        return None

    # ---------------------------------------------------------------
    # ① 採用記事を一括アップロード
    # ---------------------------------------------------------------
    def _upload_articles_batch(self):
        if not self.selected_folder:
            messagebox.showwarning("エラー", "フォルダを選択してください")
            return
        if not self._build_wp_client():
            return

        # アップロード対象の記事を収集
        html_files = sorted(
            f for f in self.selected_folder.glob("*.html")
            if not f.name.startswith("不採用_")
            and not f.name.endswith("_概要.html")
        )
        target_articles = [
            (f, self._find_eyecatch_file(f)) for f in html_files]

        if not target_articles:
            messagebox.showwarning("対象なし",
                                   "アップロード対象の記事が見つかりません")
            return

        eyecatch_count   = sum(1 for _, e in target_articles if e)
        no_eyecatch_count = len(target_articles) - eyecatch_count

        msg = f"採用記事 {len(target_articles)} 件をWordPressにアップロードします。\n"
        msg += f"  ・アイキャッチあり（アップロード対象）: {eyecatch_count}件\n"
        if no_eyecatch_count:
            msg += f"  ・アイキャッチなし（スキップ）: {no_eyecatch_count}件\n"
        msg += f"\nカテゴリー: {self.detected_category or '未判定'}\n\n続行しますか？"

        if not messagebox.askyesno("記事アップロード確認", msg):
            return

        self._set_uploading(True)
        self.status_var.set(f"記事をアップロード中... 0/{len(target_articles)}")
        self.progress_var.set(0)

        def task():
            try:
                self._do_upload_articles(target_articles)
            except Exception as e:
                self.after(0, lambda: self._on_error(f"記事アップロード失敗: {e}"))
            finally:
                self.after(0, lambda: self._set_uploading(False))

        threading.Thread(target=task, daemon=True).start()

    def _do_upload_articles(self, target_articles: list):
        """バックグラウンドで記事を一括アップロードする"""
        total         = len(target_articles)
        success_count = 0
        skip_count    = 0

        # 技術カテゴリーIDを事前に取得（全記事で共通）
        parent_cat_id: int | None = None
        if self.detected_category:
            try:
                parent_cat_id = self.wp_client.get_or_create_category(
                    self.detected_category)
                self._log_safe(
                    f"  技術カテゴリー: {self.detected_category}"
                    f" (ID: {parent_cat_id})")
            except Exception as e:
                self._log_safe(f"  ⚠ カテゴリー取得失敗: {e}")

        for i, (html_path, eyecatch_path) in enumerate(target_articles, 1):
            pct = (i - 1) / total * 100
            self.after(0, lambda p=pct, ii=i, t=total: (
                self.progress_var.set(p),
                self.status_var.set(f"記事をアップロード中... {ii}/{t}"),
            ))

            if not eyecatch_path:
                self._log_safe(
                    f"  [{i}/{total}] スキップ（アイキャッチなし）: {html_path.name}")
                skip_count += 1
                continue

            self._log_safe(f"[{i}/{total}] {html_path.name}")

            try:
                article = parse_article_html(html_path)

                # アイキャッチ画像をWPメディアにアップロード
                self._log_safe(f"  アイキャッチ: {eyecatch_path.name}")
                mime_map = {
                    ".png":  "image/png",
                    ".jpg":  "image/jpeg",
                    ".jpeg": "image/jpeg",
                }
                mime     = mime_map.get(eyecatch_path.suffix.lower(), "image/png")
                media_id = self.wp_client.upload_media(
                    eyecatch_path.read_bytes(), eyecatch_path.name, mime)
                self._log_safe(f"  メディア ID: {media_id}")

                # カテゴリーIDを構築
                category_ids: list[int] = []
                if parent_cat_id:
                    category_ids.append(parent_cat_id)

                # 記事内容から企業動向/市場動向/新技術サブカテゴリーを判定
                content_cat = detect_content_category(article)
                if content_cat:
                    try:
                        sub_id = self.wp_client.get_or_create_category(
                            content_cat,
                            parent_id=parent_cat_id,
                        )
                        if sub_id:
                            category_ids.append(sub_id)
                            self._log_safe(
                                f"  カテゴリー: {self.detected_category}"
                                f" > {content_cat}")
                    except Exception:
                        pass
                else:
                    self._log_safe(
                        f"  カテゴリー: {self.detected_category}"
                        f"（サブカテゴリー未判定）")

                # 投稿コンテンツ（緑帯ヘッダー付き）
                content = format_wp_content(
                    article["summary"],
                    article["detail"],
                    article.get("source_url", ""),
                )

                # 投稿（公開）
                result = self.wp_client.create_post(
                    title=article["title"],
                    content=content,
                    featured_media_id=media_id,
                    category_ids=category_ids if category_ids else None,
                    status="publish",
                )
                post_id   = result.get("id", "?")
                post_link = result.get("link", "")
                self._log_safe(f"  ✓ 投稿完了 ID:{post_id}  {post_link}")
                success_count += 1

            except Exception as e:
                self._log_safe(f"  ✗ エラー: {e}")

        # 完了
        self.after(0, lambda: self.progress_var.set(100))
        self.after(0, lambda: self.status_var.set(
            f"記事アップロード完了: {success_count}件成功"
            f" / {skip_count}件スキップ"))
        self._log_safe("=" * 40)
        self._log_safe(
            f"記事アップロード完了: 成功={success_count}, スキップ={skip_count}")

    # ---------------------------------------------------------------
    # ② WeeklyReport をアップロード
    # ---------------------------------------------------------------
    def _upload_weekly_report(self):
        if not self.selected_folder:
            messagebox.showwarning("エラー", "フォルダを選択してください")
            return
        if not self.detected_category:
            messagebox.showwarning(
                "カテゴリー未判定",
                "フォルダ名からカテゴリーを判定できませんでした。\n"
                "フォルダ名にカテゴリーキーワード（例: 全固体電池）を含めてください。")
            return
        if not self._build_wp_client():
            return

        # PDF・MP3を収集
        pdf_files = sorted(
            f for f in self.selected_folder.glob("*.pdf")
            if "ウィークリーレポート" in f.name
        )
        mp3_files = sorted(
            f for f in self.selected_folder.glob("*.mp3")
            if "ポッドキャスト" in f.name and not f.stem.endswith("_original")
        )

        if not pdf_files and not mp3_files:
            messagebox.showwarning(
                "ファイルなし",
                "WeeklyReport PDF または ポッドキャスト MP3 が見つかりません\n\n"
                "確認事項:\n"
                "  ・PDFファイル名に「ウィークリーレポート」を含むこと\n"
                "  ・MP3ファイル名に「ポッドキャスト」を含むこと\n"
                "  ・_original.mp3 は除外されます")
            return

        # ポッドキャストレビュー状態チェック
        if mp3_files:
            review_state_path = (
                self.selected_folder / "_review_work" / "_review_state.json")
            if review_state_path.exists():
                try:
                    import json as _json
                    rs = _json.loads(
                        review_state_path.read_text(encoding="utf-8"))
                    if rs.get("status") not in ("reviewed",):
                        if not messagebox.askyesno(
                            "ポッドキャストレビュー未完了",
                            f"ポッドキャストのレビューが完了していません\n"
                            f"（状態: {rs.get('status', 'unreviewed')}）\n\n"
                            "レビュー未完了のまま続行しますか？",
                            icon="warning",
                        ):
                            self._log("→ ポッドキャストレビュー未完了のためキャンセル")
                            return
                except Exception:
                    pass

        # 確認ダイアログ
        all_files = pdf_files + mp3_files
        file_list = "\n".join(f"  ・{f.name}" for f in all_files)
        if not messagebox.askyesno(
            "WeeklyReportアップロード確認",
            f"以下のファイルをWeeklyReport投稿としてアップロードします:\n\n"
            f"{file_list}\n\n"
            f"続行しますか？",
        ):
            return

        self._set_uploading(True)
        self.status_var.set("WeeklyReportをアップロード中...")
        self.progress_var.set(0)

        def task():
            try:
                self._do_upload_weekly_report(all_files)
            except Exception as e:
                self.after(0, lambda: self._on_error(
                    f"WeeklyReportアップロード失敗: {e}"))
            finally:
                self.after(0, lambda: self._set_uploading(False))

        threading.Thread(target=task, daemon=True).start()

    def _do_upload_weekly_report(self, files: list[Path]):
        """WeeklyReport 新規投稿を作成する（バックグラウンド実行）"""
        import json as _json

        # --- 1. 投稿タイトル用日付を決定 ---
        # ファイル名中の YYYYMMDD から取得、なければ今日の日付
        date_str = None
        for f in files:
            m = re.search(r"(\d{4})(\d{2})(\d{2})", f.name)
            if m:
                date_str = (
                    f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日")
                break
        if not date_str:
            now = datetime.now()
            date_str = f"{now.year}年{now.month}月{now.day}日"

        # 投稿タイトル: "{技術カテゴリー}{YYYY}年{M}月{D}日号"
        # 例: 全固体電池2026年4月16日号
        post_title = f"{self.detected_category}{date_str}号"
        self._log_safe(f"投稿タイトル: {post_title}")

        # --- 2. 各ファイルをWPメディアにアップロードしてブロックを生成 ---
        new_blocks: list[str] = []
        total_files = len(files)

        for fi, file_path in enumerate(files, 1):
            ext    = file_path.suffix.lower()
            is_mp3 = ext == ".mp3"
            label  = "ポッドキャスト" if is_mp3 else "WeeklyReport PDF"
            mime   = "audio/mpeg" if is_mp3 else "application/pdf"

            pct = (fi - 1) / total_files * 75  # 75%まで進める
            self.after(0, lambda p=pct, fn=file_path.name: (
                self.progress_var.set(p),
                self.status_var.set(f"{label}をアップロード中: {fn}"),
            ))
            size_kb = file_path.stat().st_size / 1024
            self._log_safe(
                f"  {label}をアップロード中: {file_path.name}"
                f" ({size_kb:.0f} KB)")

            try:
                media_id = self.wp_client.upload_media(
                    file_path.read_bytes(), file_path.name, mime)
            except Exception as e:
                self._log_safe(f"  ✗ {label}アップロード失敗: {e}")
                continue

            self._log_safe(f"  メディア ID: {media_id}")

            # メディアURLを取得
            try:
                resp = requests.get(
                    f"{self.wp_client.api_url}/media/{media_id}",
                    auth=self.wp_client.auth, timeout=15,
                )
                file_url = (
                    resp.json().get("source_url", "")
                    if resp.status_code == 200 else "")
            except Exception:
                file_url = ""

            if not file_url:
                self._log_safe(f"  ✗ {label}のURLを取得できませんでした")
                continue

            display_label = self._make_file_display_label(file_path.name)

            if is_mp3:
                block = (
                    f'<!-- wp:paragraph -->\n'
                    f'<p><strong>🎙 {display_label}</strong></p>\n'
                    f'<!-- /wp:paragraph -->\n'
                    f'<!-- wp:audio {{"id":{media_id}}} -->\n'
                    f'<figure class="wp-block-audio">'
                    f'<audio controls src="{file_url}"></audio>'
                    f'</figure>\n'
                    f'<!-- /wp:audio -->\n'
                    f'<!-- wp:file {{"id":{media_id},"href":"{file_url}"}} -->\n'
                    f'<div class="wp-block-file">'
                    f'<a id="wp-block-file--media-{media_id}" '
                    f'href="{file_url}">{file_path.name}</a>'
                    f'<a href="{file_url}" '
                    f'class="wp-block-file__button wp-element-button" '
                    f'download aria-describedby='
                    f'"wp-block-file--media-{media_id}">'
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
                    f'download aria-describedby='
                    f'"wp-block-file--media-{media_id}">'
                    f'ダウンロード</a></div>\n'
                    f'<!-- /wp:file -->\n'
                )
            new_blocks.append(block)

        if not new_blocks:
            self._log_safe("  ✗ アップロードできるファイルがありませんでした")
            return

        # --- 3. カテゴリーIDを取得 ---
        self.after(0, lambda: (
            self.status_var.set("カテゴリーを設定中..."),
            self.progress_var.set(80),
        ))
        category_ids:  list[int] = []
        parent_cat_id: int | None = None

        try:
            parent_cat_id = self.wp_client.get_or_create_category(
                self.detected_category)
            if parent_cat_id:
                category_ids.append(parent_cat_id)
                self._log_safe(
                    f"  技術カテゴリー: {self.detected_category}"
                    f" (ID: {parent_cat_id})")
        except Exception as e:
            self._log_safe(f"  ⚠ 技術カテゴリー取得失敗: {e}")

        # 子カテゴリー「最新のトピック（1週間）」
        weekly_cat_name = "最新のトピック（1週間）"
        try:
            weekly_cat_id = self.wp_client.get_or_create_category(
                weekly_cat_name,
                parent_id=parent_cat_id if parent_cat_id else None,
            )
            if weekly_cat_id:
                category_ids.append(weekly_cat_id)
                self._log_safe(
                    f"  サブカテゴリー: {weekly_cat_name}"
                    f" (ID: {weekly_cat_id})")
        except Exception as e:
            self._log_safe(f"  ⚠ サブカテゴリー取得失敗: {e}")

        # --- 4. アイキャッチ画像をメディアライブラリから検索 ---
        eyecatch_id: int | None = None
        for kw in [
            f"{self.detected_category}ウィークリーレポート",
            f"{self.detected_category}WeeklyReport",
            "ウィークリーレポート",
        ]:
            try:
                eyecatch_id = self.wp_client.search_media_by_keyword(kw)
            except Exception:
                pass
            if eyecatch_id:
                self._log_safe(f"  アイキャッチ画像 ID: {eyecatch_id}")
                break
        if not eyecatch_id:
            self._log_safe(
                "  ⚠ アイキャッチ画像が見つかりませんでした（なしで続行）")

        # --- 5. 新規投稿を作成（公開・sticky・先頭移動）---
        self.after(0, lambda: (
            self.status_var.set("投稿を作成中..."),
            self.progress_var.set(90),
        ))
        try:
            from datetime import timezone as _tz
            now_utc = datetime.now(_tz.utc).strftime("%Y-%m-%dT%H:%M:%S")
            content = "\n".join(new_blocks)
            new_post = self.wp_client.create_post(
                title=post_title,
                content=content,
                featured_media_id=eyecatch_id,
                category_ids=category_ids if category_ids else None,
                status="publish",
                sticky=True,
                date_gmt=now_utc,
            )
            new_id = new_post.get("id", "?")
            self._log_safe(
                f"  ✓ 投稿作成 (ID: {new_id})"
                f" status={new_post.get('status', '?')}")

            # 作成後に bump_to_top を別コールで確実に適用
            if isinstance(new_id, int):
                bump = self.wp_client.bump_to_top(new_id)
                self._log_safe(
                    f"  ✓ 先頭移動 sticky={bump.get('sticky', '?')}"
                    f" date={bump.get('date', '?')}")

        except Exception as e:
            self._log_safe(f"  ✗ 投稿作成失敗: {e}")
            raise

        self.after(0, lambda: (
            self.progress_var.set(100),
            self.status_var.set(
                f"WeeklyReport投稿完了: 「{post_title}」"),
        ))
        self._log_safe("=" * 40)
        self._log_safe(f"WeeklyReport投稿完了: 「{post_title}」")

    # ---------------------------------------------------------------
    # ③ 全てまとめてアップロード（記事 → WeeklyReport の順）
    # ---------------------------------------------------------------
    def _upload_all(self):
        if not self.selected_folder:
            messagebox.showwarning("エラー", "フォルダを選択してください")
            return
        if not self._build_wp_client():
            return

        # 実行前の確認
        html_files = sorted(
            f for f in self.selected_folder.glob("*.html")
            if not f.name.startswith("不採用_")
            and not f.name.endswith("_概要.html")
        )
        articles_count = sum(
            1 for f in html_files if self._find_eyecatch_file(f))
        pdf_files = [f for f in self.selected_folder.glob("*.pdf")
                     if "ウィークリーレポート" in f.name]
        mp3_files = [f for f in self.selected_folder.glob("*.mp3")
                     if "ポッドキャスト" in f.name
                     and not f.stem.endswith("_original")]

        msg = (
            "以下を順番にアップロードします:\n\n"
            f"  ① 採用記事（アイキャッチあり）: {articles_count}件\n"
            f"  ② WeeklyReport PDF: {len(pdf_files)}件\n"
            f"  ③ ポッドキャスト MP3: {len(mp3_files)}件\n\n"
            "続行しますか？"
        )
        if not messagebox.askyesno("全件アップロード確認", msg):
            return

        self._set_uploading(True)
        self.status_var.set("全件アップロードを開始します...")
        self.progress_var.set(0)

        def task():
            try:
                # ① 記事アップロード
                target_articles = [
                    (f, self._find_eyecatch_file(f)) for f in html_files]
                if target_articles:
                    self._log_safe("── 記事を一括アップロード ──")
                    self._do_upload_articles(target_articles)

                # ② WeeklyReport アップロード
                all_report_files = pdf_files + mp3_files
                if all_report_files and self.detected_category:
                    self._log_safe("── WeeklyReport をアップロード ──")
                    self._do_upload_weekly_report(all_report_files)
                elif all_report_files and not self.detected_category:
                    self._log_safe(
                        "  ⚠ カテゴリー未判定のためWeeklyReportアップロードをスキップ")

                self.after(0, lambda: (
                    self.progress_var.set(100),
                    self.status_var.set("全件アップロード完了"),
                ))
                self._log_safe("=" * 40)
                self._log_safe("全件アップロード完了")
                self.after(0, lambda: messagebox.showinfo(
                    "完了", "全てのアップロードが完了しました"))

            except Exception as e:
                self.after(0, lambda: self._on_error(
                    f"全件アップロード失敗: {e}"))
            finally:
                self.after(0, lambda: self._set_uploading(False))

        threading.Thread(target=task, daemon=True).start()

    # ---------------------------------------------------------------
    # ヘルパー
    # ---------------------------------------------------------------
    def _set_uploading(self, uploading: bool):
        """アップロード中はボタンをすべて無効化し、完了後に再有効化する"""
        self._uploading = uploading
        state = "disabled" if uploading else "normal"
        self.btn_articles.config(state=state)
        self.btn_weekly.config(state=state)
        self.btn_all.config(state=state)

    def _on_error(self, message: str):
        self._log(f"✗ エラー: {message}")
        messagebox.showerror("エラー", message)

    def _log(self, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {message}\n")
        self.log_text.see("end")

    def _log_safe(self, message: str):
        """別スレッドからのログ追加（after で UI スレッドに転送）"""
        self.after(0, lambda m=message: self._log(m))

    def _make_file_display_label(self, filename: str) -> str:
        """ファイル名から表示ラベルを生成する

        例: 全固体電池調査ウィークリーレポート20260322.pdf
          → ウィークリーレポート 2026年3月22日（PDF）をダウンロード
        例: 全固体電池調査20260322ポッドキャスト.mp3
          → ポッドキャスト 2026年3月22日（MP3）を再生・ダウンロード
        """
        ext    = Path(filename).suffix.lower()
        is_mp3 = ext == ".mp3"
        m = re.search(r"(\d{4})(\d{2})(\d{2})", filename)
        if m:
            date_str = (
                f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日")
        else:
            now = datetime.now()
            date_str = f"{now.year}年{now.month}月{now.day}日"
        return (
            f"ポッドキャスト {date_str}（MP3）を再生・ダウンロード"
            if is_mp3
            else f"ウィークリーレポート {date_str}（PDF）をダウンロード"
        )


# =====================================================================
# エントリーポイント
# =====================================================================
def main():
    app = WPUploaderApp()

    # sys.argv[1] があればフォルダ選択をスキップ
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        if p.is_dir():
            app.after(200, lambda: app._load_folder_path(p))

    app.mainloop()


if __name__ == "__main__":
    main()
