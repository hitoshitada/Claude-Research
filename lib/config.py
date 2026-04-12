"""設定・定数・プロンプトテンプレート"""

import os
from pathlib import Path
from datetime import datetime, timedelta

# ベースディレクトリ（プロジェクトルート = lib/ の親）
BASE_DIR = Path(__file__).parent.parent
PROMPTS_DIR = BASE_DIR / "prompts"


def _load_prompt(filename: str) -> str:
    """prompts/フォルダからプロンプトテキストを読み込む"""
    path = PROMPTS_DIR / filename
    return path.read_text(encoding="utf-8").strip()

# 入出力パス
INVESTIGATION_DIR = BASE_DIR / "調査内容ファイル"
OUTPUT_DIR = BASE_DIR / "調査アウトプット"
ENV_FILE = BASE_DIR / ".env"

# API設定
SEARCH_MODEL = "gemini-2.5-flash"       # Search Grounding + 構造化両方で使用
STRUCTURING_MODEL = "gemini-2.5-flash"  # 記事構造化・翻訳用

# 記事設定
TARGET_ARTICLE_COUNT = 30  # 各10件×3クエリ（重複除外後20〜30件）
TARGET_REGIONS = ["United States", "Europe (UK, Germany, France, Netherlands, etc.)", "Taiwan", "South Korea", "Japan"]

# HTTP設定
HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    )
}
HTTP_TIMEOUT = 10
IMAGE_WORKERS = 5


def build_search_queries(topic_content: str) -> list[str]:
    """Search Grounding用の検索クエリを構築する（3つの異なる角度）

    各クエリは過去7日間の記事を10件ずつ検索し、
    合計最大30件の記事カバレッジを目指す。
    プロンプトは prompts/search_queries.md から読み込む。
    """
    today = datetime.now()
    date_from = (today - timedelta(days=7)).strftime("%Y-%m-%d")
    date_to = today.strftime("%Y-%m-%d")

    template = _load_prompt("search_queries.md")
    # コメント行（# で始まる行）を除去
    lines = [l for l in template.split("\n") if not l.strip().startswith("#")]
    template_clean = "\n".join(lines).strip()

    # --- で区切って各クエリテンプレートを取得
    query_templates = [q.strip() for q in template_clean.split("---") if q.strip()]

    queries = []
    for qt in query_templates:
        query = qt.replace("{date_from}", date_from) \
                   .replace("{date_to}", date_to) \
                   .replace("{topic_content}", topic_content)
        queries.append(query)

    return queries


def build_structuring_prompt(topic: str) -> str:
    """マークダウンレポートをJSON構造化+日本語翻訳するプロンプト（外部ファイルから読み込み）"""
    template = _load_prompt("structuring.md")
    return template.replace("{topic}", topic) + "\n"


def build_summary_prompt(topic: str) -> str:
    """ウィークリーサマリー生成プロンプト（外部ファイルから読み込み）"""
    template = _load_prompt("summary.md")
    return template.replace("{topic}", topic) + "\n"
