"""Search Groundingレポートから記事を構造化抽出+日本語翻訳するモジュール"""

import json
import re
import requests
from dataclasses import dataclass, field
from typing import Optional
from google import genai
from .config import STRUCTURING_MODEL, build_structuring_prompt, HTTP_HEADERS
from .search_grounding_client import ResearchResult


@dataclass
class Article:
    """記事データ"""
    title_ja: str = ""
    url: str = ""
    source_name: str = ""
    country: str = ""
    publish_date: str = ""
    summary_ja: str = ""
    content_ja: str = ""
    image_url: str | None = None


def resolve_redirect_urls(
    urls: list[str],
    progress_callback=None,
) -> dict[str, str]:
    """リダイレクトURLを実際のURLに解決する

    Args:
        urls: リダイレクトURLのリスト
        progress_callback: 進捗通知用コールバック

    Returns:
        {リダイレクトURL: 実際のURL} のマッピング
    """
    resolved = {}
    total = len(urls)

    for i, url in enumerate(urls):
        if progress_callback and (i % 5 == 0 or i == total - 1):
            progress_callback(f"URL解決中: {i + 1}/{total}")
        try:
            if "grounding-api-redirect" in url or "vertexaisearch" in url:
                resp = requests.head(
                    url, headers=HTTP_HEADERS, allow_redirects=False, timeout=10
                )
                location = resp.headers.get("Location")
                if location:
                    resolved[url] = location
                else:
                    # allow_redirectsで再試行
                    resp2 = requests.head(
                        url, headers=HTTP_HEADERS, allow_redirects=True, timeout=10
                    )
                    resolved[url] = resp2.url
            else:
                resolved[url] = url
        except requests.RequestException as e:
            # 解決できなくてもリダイレクトURLをそのまま使う
            resolved[url] = url

    return resolved


def extract_articles_from_report(
    client: genai.Client,
    research_result: ResearchResult,
    topic: str,
    url_mapping: Optional[dict[str, str]] = None,
) -> list[Article]:
    """Search Groundingレポートから記事を構造化抽出+日本語翻訳する

    Args:
        client: Gemini APIクライアント
        research_result: Search Groundingの結果
        topic: 調査テーマ
        url_mapping: {リダイレクトURL: 実際のURL} マッピング

    Returns:
        Article のリスト
    """
    markdown_report = research_result.text

    # URL マッピング情報を構築
    url_info = ""
    if url_mapping:
        url_lines = []
        for redirect_url, actual_url in url_mapping.items():
            if redirect_url != actual_url:
                url_lines.append(f"  {redirect_url}\n    -> {actual_url}")
            else:
                url_lines.append(f"  {actual_url}")
        url_info = "\n\n## 参照元URL一覧（実際のURL）\n" + "\n".join(url_lines)

    # プロンプト構築
    prompt = build_structuring_prompt(topic)

    # レポートが長すぎる場合は切り詰め
    max_report_length = 80000
    report_text = markdown_report
    if len(report_text) > max_report_length:
        report_text = report_text[:max_report_length] + "\n\n[レポートの残りは省略されました]"

    full_prompt = f"{prompt}\n\n## リサーチレポート\n{report_text}{url_info}"

    # Gemini APIで構造化+翻訳
    try:
        response = client.models.generate_content(
            model=STRUCTURING_MODEL,
            contents=full_prompt,
            config={
                "response_mime_type": "application/json",
            },
        )

        return _parse_response(response.text, url_mapping)

    except json.JSONDecodeError as e:
        # JSONパースに失敗した場合、リトライ
        return _retry_extraction(client, markdown_report, topic, url_mapping)

    except Exception as e:
        # 1回リトライ
        try:
            return _retry_extraction(client, markdown_report, topic, url_mapping)
        except Exception as e2:
            raise RuntimeError(f"記事抽出に失敗: {str(e)} / リトライ: {str(e2)}") from e2


def _parse_response(
    response_text: str,
    url_mapping: Optional[dict[str, str]] = None,
) -> list[Article]:
    """GeminiのJSONレスポンスをパースしてArticleリストに変換する"""
    text = response_text.strip()

    # コードブロックでラップされている場合を処理
    if text.startswith("```"):
        # ```json ... ``` のパターンを抽出
        match = re.search(r"```(?:json)?\s*\n(.*?)\n```", text, re.DOTALL)
        if match:
            text = match.group(1).strip()

    articles_data = json.loads(text)

    # トップレベルが辞書で"articles"キーがある場合
    if isinstance(articles_data, dict):
        if "articles" in articles_data:
            articles_data = articles_data["articles"]
        else:
            # 最初のリスト値を使用
            for val in articles_data.values():
                if isinstance(val, list):
                    articles_data = val
                    break

    if not isinstance(articles_data, list):
        raise ValueError(f"予期しないJSONフォーマット: {type(articles_data)}")

    articles = []
    for item in articles_data:
        if not isinstance(item, dict):
            continue

        article = Article(
            title_ja=item.get("title_ja", "") or item.get("title", ""),
            url=item.get("url", ""),
            source_name=item.get("source_name", "") or item.get("source", ""),
            country=item.get("country", "") or item.get("region", ""),
            publish_date=item.get("publish_date", "") or item.get("date", ""),
            summary_ja=item.get("summary_ja", "") or item.get("summary", ""),
            content_ja=item.get("content_ja", "") or item.get("content", ""),
        )

        # URLマッピングの適用（リダイレクトURLが残っている場合）
        if url_mapping and article.url in url_mapping:
            article.url = url_mapping[article.url]

        articles.append(article)

    return articles


def _retry_extraction(
    client: genai.Client,
    markdown_report: str,
    topic: str,
    url_mapping: Optional[dict[str, str]] = None,
) -> list[Article]:
    """構造化抽出のリトライ（よりシンプルなプロンプトで）"""
    simple_prompt = f"""以下のリサーチレポートから各記事の情報を抽出し、JSON配列として返してください。

各記事のフィールド:
- "title_ja": 日本語タイトル
- "url": 記事URL（不明なら空文字列）
- "source_name": メディア名
- "country": 発信国（日本語）
- "publish_date": 発行日（YYYY-MM-DD）
- "summary_ja": 日本語要約（3文程度）
- "content_ja": 日本語詳細（HTML形式、<p>タグで段落分け、300文字以上）

レポート:
{markdown_report[:50000]}

JSON配列のみ返してください。
"""

    response = client.models.generate_content(
        model=STRUCTURING_MODEL,
        contents=simple_prompt,
        config={
            "response_mime_type": "application/json",
        },
    )

    return _parse_response(response.text, url_mapping)
