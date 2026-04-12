"""
テスト: 既存HTMLファイルから記事データを復元し、WeeklyReport を生成する。

使い方:
  python -m tests.test_weekly_report "接着・封止材_20260411"
"""

import json
import os
import re
import sys
from pathlib import Path
from html.parser import HTMLParser
from dataclasses import dataclass, field

# プロジェクトルートをパスに追加
PROJECT_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from dotenv import load_dotenv
load_dotenv(PROJECT_DIR / ".env")

from google import genai
from lib.weekly_report_generator import (
    build_articles_data,
    generate_front_report_json,
    generate_front_report_pdf,
    concatenate_pdfs,
)

OUTPUT_DIR = PROJECT_DIR / "調査アウトプット"


# ===================================================================
# HTML から記事データを復元
# ===================================================================

@dataclass
class Article:
    """pdf_generator の Article 互換"""
    title_ja: str = ""
    url: str = ""
    source_name: str = ""
    country: str = ""
    publish_date: str = ""
    summary_ja: str = ""
    content_ja: str = ""
    image_url: str | None = None


class _SummaryHTMLParser(HTMLParser):
    """概要HTML から全記事のメタデータ + サマリーを抽出する"""

    def __init__(self):
        super().__init__()
        self.articles: list[dict] = []
        self._current: dict = {}
        self._in_title = False
        self._in_meta = False
        self._in_summary = False
        self._meta_spans: list[str] = []
        self._span_data = ""

    def handle_starttag(self, tag, attrs):
        classes = dict(attrs).get("class", "")
        if tag == "div" and "article-item" in classes:
            self._current = {}
        elif tag == "div" and "title" in classes:
            self._in_title = True
        elif tag == "div" and "article-meta" in classes:
            self._in_meta = True
            self._meta_spans = []
        elif tag == "div" and "summary" in classes:
            self._in_summary = True
        elif tag == "a" and self._in_title:
            href = dict(attrs).get("href", "")
            self._current["html_file"] = href
        elif tag == "span" and self._in_meta:
            self._span_data = ""

    def handle_data(self, data):
        if self._in_title:
            self._current["title"] = data.strip()
        elif self._in_meta and self._span_data is not None:
            self._span_data += data
        elif self._in_summary:
            self._current["summary"] = data.strip()

    def handle_endtag(self, tag):
        if tag == "div" and self._in_title:
            self._in_title = False
        elif tag == "span" and self._in_meta:
            self._meta_spans.append(self._span_data.strip())
            self._span_data = ""
        elif tag == "div" and self._in_meta:
            self._in_meta = False
            if len(self._meta_spans) >= 3:
                date_text = self._meta_spans[0].replace("公開日 ", "").strip()
                self._current["date"] = date_text
                self._current["source"] = self._meta_spans[1].strip()
                self._current["country"] = self._meta_spans[2].strip()
        elif tag == "div" and self._in_summary:
            self._in_summary = False
            if self._current.get("title"):
                self.articles.append(dict(self._current))


class _DetailHTMLParser(HTMLParser):
    """個別記事HTMLから詳細テキストとURLを抽出"""

    def __init__(self):
        super().__init__()
        self.detail_text = ""
        self.url = ""
        self._in_detail = False
        self._in_source = False

    def handle_starttag(self, tag, attrs):
        classes = dict(attrs).get("class", "")
        if tag == "div" and "detail" in classes:
            self._in_detail = True
        elif tag == "div" and "source" in classes:
            self._in_source = True
        elif tag == "a" and self._in_source:
            self.url = dict(attrs).get("href", "")

    def handle_data(self, data):
        if self._in_detail:
            self.detail_text += data

    def handle_endtag(self, tag):
        if tag == "div" and self._in_detail:
            self._in_detail = False
        elif tag == "div" and self._in_source:
            self._in_source = False


def parse_articles_from_folder(folder: Path) -> list[Article]:
    """出力フォルダの概要HTML + 個別HTMLから Article リストを復元する。"""
    # 概要HTMLを探す
    summary_files = list(folder.glob("*_概要.html"))
    if not summary_files:
        raise FileNotFoundError(f"概要HTMLが見つかりません: {folder}")

    summary_html = summary_files[0].read_text(encoding="utf-8")
    parser = _SummaryHTMLParser()
    parser.feed(summary_html)

    articles = []
    for item in parser.articles:
        # タイトルから先頭の番号プレフィックスを除去
        title = re.sub(r"^\d+_", "", item.get("title", ""))

        # 日付を YYYY-MM-DD に変換
        date_raw = item.get("date", "")
        m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", date_raw)
        publish_date = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else date_raw

        # 個別HTMLから詳細を取得
        html_file = item.get("html_file", "")
        detail_text = ""
        url = ""
        if html_file:
            detail_path = folder / html_file
            if detail_path.exists():
                detail_html = detail_path.read_text(encoding="utf-8")
                dp = _DetailHTMLParser()
                dp.feed(detail_html)
                detail_text = dp.detail_text.strip()
                url = dp.url

        articles.append(Article(
            title_ja=title,
            url=url,
            source_name=item.get("source", ""),
            country=item.get("country", ""),
            publish_date=publish_date,
            summary_ja=item.get("summary", ""),
            content_ja=detail_text,
        ))

    return articles


# ===================================================================
# メイン
# ===================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python -m tests.test_weekly_report <フォルダ名>")
        print("Example: python -m tests.test_weekly_report 接着・封止材_20260411")
        sys.exit(1)

    folder_name = sys.argv[1]
    folder = OUTPUT_DIR / folder_name
    if not folder.exists():
        print(f"フォルダが見つかりません: {folder}")
        sys.exit(1)

    # トピック名と日付を抽出
    m = re.match(r"(.+?)_(\d{8})$", folder_name)
    if m:
        topic_name = m.group(1)
        date_raw = m.group(2)
        collection_date = f"{date_raw[:4]}年{date_raw[4:6]}月{date_raw[6:8]}日"
    else:
        topic_name = folder_name
        collection_date = "2026年04月11日"

    print(f"=== WeeklyReport テスト ===")
    print(f"フォルダ: {folder}")
    print(f"トピック: {topic_name}")
    print(f"収集日:   {collection_date}")
    print()

    # Step 1: 記事データ復元
    print("[1/4] HTMLファイルから記事データを復元中...")
    articles = parse_articles_from_folder(folder)
    print(f"  → {len(articles)}件の記事を復元")

    # Step 1.5: articles_data.json 生成
    articles_data = build_articles_data(articles, topic_name, collection_date)
    json_path = folder / "articles_data.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(articles_data, f, ensure_ascii=False, indent=2)
    print(f"  → {json_path.name} を保存")

    # Step 2: Gemini 分析
    print("[2/4] Gemini APIで記事分析中...")
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("エラー: GEMINI_API_KEY または GOOGLE_API_KEY が設定されていません")
        sys.exit(1)
    client = genai.Client(api_key=api_key)

    front_report = generate_front_report_json(
        client, articles_data,
        progress_callback=lambda msg: print(f"  → {msg}"),
    )

    front_json_path = folder / "front_report.json"
    with open(front_json_path, "w", encoding="utf-8") as f:
        json.dump(front_report, f, ensure_ascii=False, indent=2)
    print(f"  → {front_json_path.name} を保存")

    # Step 3: フロントレポートPDF生成
    print("[3/4] フロントレポートPDF生成中...")
    front_pdf_path = folder / f"{topic_name}_フロントレポート.pdf"
    generate_front_report_pdf(front_report, front_pdf_path)
    print(f"  → {front_pdf_path.name} を保存")

    # Step 4: PDF連結
    # 既存の記事PDFを探す
    existing_pdfs = list(folder.glob(f"{topic_name}ウィークリーレポート*.pdf"))
    if existing_pdfs:
        existing_pdf = existing_pdfs[0]
        print(f"[4/4] PDF連結中... (既存: {existing_pdf.name})")
        final_path = folder / f"{topic_name}ウィークリーレポート_NEW.pdf"
        concatenate_pdfs(front_pdf_path, existing_pdf, final_path)
        print(f"  → {final_path.name} を保存")
    else:
        print("[4/4] 既存記事PDFなし、フロントレポートのみ")
        final_path = front_pdf_path

    print()
    print(f"完了! 出力: {final_path}")


if __name__ == "__main__":
    main()
