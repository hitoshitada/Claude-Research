"""HTML記事ファイルを生成するモジュール"""

import os
import re
import unicodedata
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
from article_parser import Article
from config import OUTPUT_DIR


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <style>
    * {{
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }}
    body {{
      font-family: 'Segoe UI', 'Meiryo', 'Hiragino Kaku Gothic ProN', sans-serif;
      line-height: 1.8;
      color: #333;
      background-color: #f5f5f5;
      padding: 20px;
    }}
    .container {{
      max-width: 800px;
      margin: 0 auto;
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.1);
      overflow: hidden;
    }}
    .header {{
      background: linear-gradient(135deg, #1a237e, #283593);
      color: white;
      padding: 30px;
    }}
    .header h1 {{
      font-size: 1.5em;
      margin-bottom: 10px;
      line-height: 1.4;
    }}
    .meta {{
      font-size: 0.9em;
      opacity: 0.85;
    }}
    .meta span {{
      margin-right: 15px;
    }}
    .article-image {{
      width: 100%;
      text-align: center;
      background: #eee;
    }}
    .article-image img {{
      max-width: 100%;
      height: auto;
      max-height: 400px;
      object-fit: cover;
    }}
    .content {{
      padding: 30px;
    }}
    .section {{
      margin-bottom: 25px;
    }}
    .section-title {{
      font-size: 1.1em;
      font-weight: bold;
      color: #1a237e;
      border-left: 4px solid #1a237e;
      padding-left: 12px;
      margin-bottom: 12px;
    }}
    .overview {{
      background: #e8eaf6;
      padding: 15px 20px;
      border-radius: 8px;
      font-size: 0.95em;
    }}
    .detail p {{
      margin-bottom: 12px;
      text-align: justify;
    }}
    .detail h3 {{
      color: #283593;
      margin: 18px 0 8px 0;
      font-size: 1.05em;
    }}
    .detail ul {{
      margin: 8px 0 12px 20px;
    }}
    .detail li {{
      margin-bottom: 6px;
    }}
    .source {{
      text-align: right;
      padding: 15px 30px;
      border-top: 1px solid #eee;
      font-size: 0.85em;
    }}
    .source a {{
      color: #1a237e;
      text-decoration: none;
    }}
    .source a:hover {{
      text-decoration: underline;
    }}
    .footer {{
      text-align: center;
      padding: 15px;
      background: #f5f5f5;
      font-size: 0.8em;
      color: #999;
    }}
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>{title}</h1>
      <div class="meta">
        <span>公開日 {publish_date}</span>
        <span>{source_name}</span>
        <span>{country}</span>
      </div>
    </div>

    <div class="article-image">
      <img src="{image_url}" alt="{title}" onerror="this.parentElement.style.display='none'">
    </div>

    <div class="content">
      <div class="section">
        <div class="section-title">概要</div>
        <div class="overview">{summary}</div>
      </div>

      <div class="section">
        <div class="section-title">詳細</div>
        <div class="detail">
{detail}
        </div>
      </div>
    </div>

    <div class="source">
      元記事: <a href="{source_url}" target="_blank">{source_url}</a>
    </div>

    <div class="footer">
      収集日: {collection_date} | 自動記事収集・翻訳システム (Gemini API使用)
    </div>
  </div>
</body>
</html>"""


SUMMARY_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{topic_name} - 記事概要一覧</title>
  <style>
    * {{
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }}
    body {{
      font-family: 'Segoe UI', 'Meiryo', 'Hiragino Kaku Gothic ProN', sans-serif;
      line-height: 1.8;
      color: #333;
      background-color: #f5f5f5;
      padding: 20px;
    }}
    .container {{
      max-width: 1000px;
      margin: 0 auto;
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.1);
      overflow: hidden;
    }}
    .header {{
      background: linear-gradient(135deg, #1a237e, #283593);
      color: white;
      padding: 30px;
    }}
    .header h1 {{
      font-size: 1.6em;
      margin-bottom: 8px;
    }}
    .header .meta {{
      font-size: 0.9em;
      opacity: 0.85;
    }}
    .article-list {{
      padding: 20px 30px;
    }}
    .article-item {{
      border-bottom: 1px solid #e0e0e0;
      padding: 18px 0;
    }}
    .article-item:last-child {{
      border-bottom: none;
    }}
    .article-item .number {{
      display: inline-block;
      background: #1a237e;
      color: white;
      width: 32px;
      height: 32px;
      line-height: 32px;
      text-align: center;
      border-radius: 50%;
      font-size: 0.85em;
      font-weight: bold;
      margin-right: 10px;
      vertical-align: top;
    }}
    .article-item .info {{
      display: inline-block;
      width: calc(100% - 50px);
      vertical-align: top;
    }}
    .article-item .title {{
      font-size: 1.05em;
      font-weight: bold;
      color: #1a237e;
      margin-bottom: 4px;
    }}
    .article-item .title a {{
      color: #1a237e;
      text-decoration: none;
    }}
    .article-item .title a:hover {{
      text-decoration: underline;
    }}
    .article-item .article-meta {{
      font-size: 0.82em;
      color: #888;
      margin-bottom: 6px;
    }}
    .article-item .article-meta span {{
      margin-right: 12px;
    }}
    .article-item .summary {{
      font-size: 0.92em;
      color: #555;
      background: #f8f9ff;
      padding: 10px 14px;
      border-radius: 6px;
      border-left: 3px solid #c5cae9;
    }}
    .footer {{
      text-align: center;
      padding: 15px;
      background: #f5f5f5;
      font-size: 0.8em;
      color: #999;
    }}
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>{topic_name} - 記事概要一覧</h1>
      <div class="meta">収集日: {collection_date} ｜ 全{article_count}件</div>
    </div>
    <div class="article-list">
{article_items}
    </div>
    <div class="footer">
      自動記事収集・翻訳システム (Gemini API使用)
    </div>
  </div>
</body>
</html>"""


SUMMARY_ITEM_TEMPLATE = """      <div class="article-item">
        <span class="number">{number}</span>
        <div class="info">
          <div class="title"><a href="{html_filename}">{title}</a></div>
          <div class="article-meta">
            <span>公開日 {publish_date}</span>
            <span>{source_name}</span>
            <span>{country}</span>
          </div>
          <div class="summary">{summary}</div>
        </div>
      </div>"""


# =====================================================================
# ファイル名生成
# =====================================================================

def _sanitize_filename(text: str, max_length: int = 50) -> str:
    """タイトルをファイル名に使える安全な文字列にする（最大max_length文字）

    - Windowsファイル名で使えない文字を除去
    - 全角・半角を問わず50文字以内に切り詰め
    """
    if not text:
        return "無題"

    # ファイル名に使えない文字を除去/置換
    # Windows禁止文字: \ / : * ? " < > |
    sanitized = re.sub(r'[\\/:*?"<>|]', '', text)
    # 制御文字を除去
    sanitized = re.sub(r'[\x00-\x1f\x7f]', '', sanitized)
    # 先頭・末尾の空白・ドットを除去
    sanitized = sanitized.strip(' .')
    # 連続空白を1つに
    sanitized = re.sub(r'\s+', ' ', sanitized)

    # max_length文字に切り詰め
    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length].rstrip()

    if not sanitized:
        return "無題"

    return sanitized


def make_article_filename(article_number: int, title: str) -> str:
    """記事番号+タイトル50字以内のファイル名を生成する

    例: "01_全固体電池の超高速充電技術が独立検証で確認される.html"

    Args:
        article_number: 記事番号
        title: 記事タイトル（日本語）

    Returns:
        ファイル名文字列（.html拡張子付き）
    """
    safe_title = _sanitize_filename(title, max_length=50)
    return f"{article_number:02d}_{safe_title}.html"


# =====================================================================
# 既存記事番号の検出
# =====================================================================

def get_next_article_number(output_folder: Path) -> int:
    """次の記事番号を取得する（既存ファイルの続き）

    フォルダ内の "NN_*.html" パターンのファイルから最大番号を検出する。

    Args:
        output_folder: 出力フォルダ

    Returns:
        次の記事番号（1始まり）
    """
    existing = list(output_folder.glob("*.html"))
    if not existing:
        return 1

    numbers = []
    for f in existing:
        # "01_タイトル.html" の先頭2桁を取得
        match = re.match(r"^(\d{2})_", f.name)
        if match:
            numbers.append(int(match.group(1)))

    return max(numbers) + 1 if numbers else 1


# =====================================================================
# 記事HTML生成・保存
# =====================================================================

def generate_article_html(
    article: Article,
    collection_date: str | None = None,
) -> str:
    """記事データからHTMLを生成する"""
    if not collection_date:
        collection_date = datetime.now().strftime("%Y年%m月%d日")

    publish_date = article.publish_date
    if publish_date and re.match(r"\d{4}-\d{2}-\d{2}", publish_date):
        try:
            dt = datetime.strptime(publish_date, "%Y-%m-%d")
            publish_date = dt.strftime("%Y年%m月%d日")
        except ValueError:
            pass

    detail = article.content_ja
    if not detail:
        detail = f"<p>{article.summary_ja}</p>"

    image_url = article.image_url or ""

    html = HTML_TEMPLATE.format(
        title=_escape_html(article.title_ja),
        publish_date=publish_date or "日付不明",
        source_name=_escape_html(article.source_name or "出典不明"),
        country=_escape_html(article.country or ""),
        image_url=image_url,
        summary=_escape_html_preserve_tags(article.summary_ja),
        detail=detail,
        source_url=article.url or "#",
        collection_date=collection_date,
    )

    return html


def save_article_html(
    html_content: str,
    output_folder: Path,
    article_number: int,
    title: str,
) -> Path:
    """HTML記事ファイルを保存する

    ファイル名は "NN_タイトル50字以内.html" 形式。

    Args:
        html_content: HTMLテキスト
        output_folder: 出力フォルダ
        article_number: 記事番号
        title: 記事タイトル

    Returns:
        保存されたファイルのPath
    """
    filename = make_article_filename(article_number, title)
    filepath = output_folder / filename

    filepath.write_text(html_content, encoding="utf-8")
    return filepath


# =====================================================================
# 概要一覧HTML生成
# =====================================================================

@dataclass
class SavedArticleInfo:
    """保存済み記事の情報（概要一覧用）"""
    number: int
    filename: str        # 拡張子付きファイル名
    title: str
    source_name: str
    country: str
    publish_date: str
    summary: str


def generate_summary_html(
    topic_name: str,
    saved_articles: list[SavedArticleInfo],
    collection_date: str | None = None,
) -> str:
    """記事概要一覧HTMLを生成する

    Args:
        topic_name: 調査テーマ名
        saved_articles: 保存済み記事情報のリスト
        collection_date: 収集日

    Returns:
        概要一覧のHTMLテキスト
    """
    if not collection_date:
        collection_date = datetime.now().strftime("%Y年%m月%d日")

    # 各記事のHTML断片を生成
    items_html_parts = []
    for info in saved_articles:
        # 発行日フォーマット
        pub_date = info.publish_date
        if pub_date and re.match(r"\d{4}-\d{2}-\d{2}", pub_date):
            try:
                dt = datetime.strptime(pub_date, "%Y-%m-%d")
                pub_date = dt.strftime("%Y年%m月%d日")
            except ValueError:
                pass

        # ファイル名から .html を除いた表示名
        display_name = info.filename
        if display_name.endswith(".html"):
            display_name = display_name[:-5]

        item_html = SUMMARY_ITEM_TEMPLATE.format(
            number=info.number,
            html_filename=info.filename,
            title=_escape_html(display_name),
            publish_date=pub_date or "日付不明",
            source_name=_escape_html(info.source_name or "出典不明"),
            country=_escape_html(info.country or ""),
            summary=_escape_html(info.summary),
        )
        items_html_parts.append(item_html)

    article_items = "\n".join(items_html_parts)

    html = SUMMARY_HTML_TEMPLATE.format(
        topic_name=_escape_html(topic_name),
        collection_date=collection_date,
        article_count=len(saved_articles),
        article_items=article_items,
    )

    return html


def save_summary_html(
    html_content: str,
    output_folder: Path,
    topic_name: str,
) -> Path:
    """概要一覧HTMLを保存する

    ファイル名: "{topic_name}_概要.html"

    Args:
        html_content: 概要一覧HTMLテキスト
        output_folder: 出力フォルダ
        topic_name: 調査テーマ名

    Returns:
        保存されたファイルのPath
    """
    filename = f"{topic_name}_概要.html"
    filepath = output_folder / filename

    filepath.write_text(html_content, encoding="utf-8")
    return filepath


# =====================================================================
# 出力フォルダ
# =====================================================================

def make_output_folder(topic_name: str, date_str: str | None = None) -> Path:
    """出力フォルダを作成する（既存の場合は再利用）"""
    if not date_str:
        date_str = datetime.now().strftime("%Y%m%d")

    folder_name = f"{topic_name}_{date_str}"
    output_path = OUTPUT_DIR / folder_name

    output_path.mkdir(parents=True, exist_ok=True)
    return output_path


# =====================================================================
# ヘルパー
# =====================================================================

def _escape_html(text: str) -> str:
    """HTMLエスケープ"""
    if not text:
        return ""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    return text


def _escape_html_preserve_tags(text: str) -> str:
    """概要部分のHTMLエスケープ（HTMLタグがあればそのまま）"""
    if not text:
        return ""
    if "<" not in text:
        return _escape_html(text)
    return text
