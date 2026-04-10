"""Search Grounding API操作モジュール（Google検索付きGemini API）"""

import re
import time
from dataclasses import dataclass, field
from typing import Callable, Optional

from google import genai
from google.genai import types
from config import SEARCH_MODEL, build_search_queries


@dataclass
class ResearchResult:
    """検索結果（下流パイプラインとの互換インターフェース）"""
    text: str = ""
    source_urls: list[str] = field(default_factory=list)
    annotations: list[dict] = field(default_factory=list)


def create_client(api_key: str) -> genai.Client:
    """Gemini APIクライアントを作成"""
    return genai.Client(api_key=api_key)


def _inspect_response_metadata(response, log_callback=None) -> str:
    """レスポンスのグラウンディングメタデータ構造をデバッグ用に検査する"""
    debug_lines = []

    try:
        candidates = getattr(response, "candidates", None)
        debug_lines.append(f"candidates: {'あり' if candidates else 'なし'}")

        if not candidates or len(candidates) == 0:
            debug_lines.append("candidates が空です")
            return " | ".join(debug_lines)

        candidate = candidates[0]
        # candidate の属性を列挙
        candidate_attrs = [a for a in dir(candidate) if not a.startswith("_")]
        debug_lines.append(f"candidate属性: {', '.join(candidate_attrs[:15])}")

        metadata = getattr(candidate, "grounding_metadata", None)
        debug_lines.append(f"grounding_metadata: {'あり' if metadata else 'なし'}")

        if metadata:
            meta_attrs = [a for a in dir(metadata) if not a.startswith("_")]
            debug_lines.append(f"metadata属性: {', '.join(meta_attrs[:15])}")

            # grounding_chunks
            chunks = getattr(metadata, "grounding_chunks", None)
            debug_lines.append(f"grounding_chunks: {type(chunks).__name__}({len(chunks) if chunks else 0})")

            if chunks and len(chunks) > 0:
                chunk0 = chunks[0]
                chunk_attrs = [a for a in dir(chunk0) if not a.startswith("_")]
                debug_lines.append(f"chunk[0]属性: {', '.join(chunk_attrs[:10])}")
                web = getattr(chunk0, "web", None)
                if web:
                    web_attrs = [a for a in dir(web) if not a.startswith("_")]
                    debug_lines.append(f"chunk[0].web属性: {', '.join(web_attrs[:10])}")

            # grounding_supports
            supports = getattr(metadata, "grounding_supports", None)
            if supports:
                debug_lines.append(f"grounding_supports: {len(supports)}件")

            # search_entry_point
            sep = getattr(metadata, "search_entry_point", None)
            if sep:
                debug_lines.append(f"search_entry_point: あり")

            # retrieval_metadata
            rm = getattr(metadata, "retrieval_metadata", None)
            if rm:
                debug_lines.append(f"retrieval_metadata: あり")

    except Exception as e:
        debug_lines.append(f"検査エラー: {str(e)}")

    result = " | ".join(debug_lines)
    if log_callback:
        log_callback(f"[DEBUG] {result}")
    return result


def _extract_grounding_urls(response, log_callback=None) -> list[dict]:
    """レスポンスからGrounding URLとタイトルを抽出する

    複数のメタデータパスを試行し、最も多くのURLを取得する。

    Returns:
        [{"url": "https://...", "title": "記事タイトル"}, ...]
    """
    results = []

    try:
        candidates = getattr(response, "candidates", None)
        if not candidates or len(candidates) == 0:
            if log_callback:
                log_callback("[DEBUG] URL抽出: candidates が空")
            return results

        metadata = getattr(candidates[0], "grounding_metadata", None)
        if not metadata:
            if log_callback:
                log_callback("[DEBUG] URL抽出: grounding_metadata が見つからない")
            return results

        # ---- パス1: grounding_chunks (最も一般的) ----
        chunks = getattr(metadata, "grounding_chunks", None)
        if chunks:
            for chunk in chunks:
                web = getattr(chunk, "web", None)
                if web:
                    uri = getattr(web, "uri", None)
                    title = getattr(web, "title", "")
                else:
                    uri = getattr(chunk, "uri", None)
                    title = getattr(chunk, "title", "")

                if uri:
                    results.append({"url": uri, "title": title or ""})

        # ---- パス2: grounding_supports (SDK v1.x系の一部) ----
        if not results:
            supports = getattr(metadata, "grounding_supports", None)
            if supports:
                for support in supports:
                    # grounding_chunk_indices + 直接URLの場合
                    seg = getattr(support, "segment", None)
                    chunks_ref = getattr(support, "grounding_chunk_indices", None)
                    # web_search_queries にURLが含まれる場合
                    web_queries = getattr(support, "web_search_queries", None)

                    # segment からURLを抽出
                    if seg:
                        seg_text = getattr(seg, "text", "") or str(seg)
                        urls_in_seg = re.findall(r'https?://[^\s<>"\']+', seg_text)
                        for url in urls_in_seg:
                            results.append({"url": url, "title": ""})

        # ---- パス3: retrieval_metadata (Search Grounding v2) ----
        if not results:
            rm = getattr(metadata, "retrieval_metadata", None)
            if rm:
                web_dynamic = getattr(rm, "google_search_dynamic_retrieval_score", None)
                if log_callback:
                    log_callback(f"[DEBUG] retrieval_metadata 検出 (dynamic score: {web_dynamic})")

        if log_callback:
            log_callback(f"[DEBUG] grounding_chunks から {len(results)} 件のURL抽出")

    except Exception as e:
        if log_callback:
            log_callback(f"[DEBUG] URL抽出エラー: {str(e)}")

    return results


def _extract_urls_from_text(text: str) -> list[str]:
    """テキスト本文からURLをフォールバック抽出する"""
    url_pattern = r'https?://(?:www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b[-a-zA-Z0-9()@:%_\+.~#?&//=]*'
    found = re.findall(url_pattern, text)

    # フィルタリング: Google内部URL、画像URL、短すぎるURLを除外
    filtered = []
    seen = set()
    for url in found:
        # Google内部URLや画像URLを除外
        skip_patterns = [
            "google.com/search",
            "googleapis.com",
            "gstatic.com",
            "youtube.com",
            "youtu.be",
            ".jpg", ".jpeg", ".png", ".gif", ".svg",
            ".css", ".js",
        ]
        if any(pat in url.lower() for pat in skip_patterns):
            continue

        # 重複排除
        clean_url = url.rstrip(".,;:)")
        if clean_url not in seen and len(clean_url) > 20:
            seen.add(clean_url)
            filtered.append(clean_url)

    return filtered


def search_articles(
    client: genai.Client,
    topic_content: str,
    progress_callback: Optional[Callable[[str, float], None]] = None,
    stop_check: Optional[Callable[[], bool]] = None,
) -> Optional[ResearchResult]:
    """Search Groundingを使って記事を検索する

    3つの異なる角度（欧米/アジア/技術）で検索を行い、
    結果を統合してResearchResultとして返す。

    Args:
        client: Gemini APIクライアント
        topic_content: 調査テーマの内容
        progress_callback: 進捗コールバック (メッセージ, 経過秒)
        stop_check: 中止チェック関数

    Returns:
        ResearchResult（テキスト+ソースURL）、失敗時はNone
    """
    start_time = time.time()
    queries = build_search_queries(topic_content)
    query_labels = ["欧米記事を検索中", "アジア記事を検索中", "技術・専門記事を検索中"]

    all_text_parts: list[str] = []
    all_urls: list[str] = []
    all_annotations: list[dict] = []
    seen_urls: set[str] = set()

    # Google Searchツールの設定
    search_tool = types.Tool(google_search=types.GoogleSearch())
    config = types.GenerateContentConfig(tools=[search_tool])

    def _log(msg: str):
        if progress_callback:
            progress_callback(msg, time.time() - start_time)

    for i, query in enumerate(queries):
        # 中止チェック
        if stop_check and stop_check():
            _log("調査が中止されました")
            return None

        elapsed = time.time() - start_time
        label = query_labels[i] if i < len(query_labels) else f"検索中 ({i + 1})"

        _log(f"{label}... ({i + 1}/{len(queries)})")

        # Search Grounding付きAPI呼び出し（リトライ付き）
        response = None
        last_error = ""
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=SEARCH_MODEL,
                    contents=query,
                    config=config,
                )
                break
            except Exception as e:
                last_error = str(e)
                if attempt < 2:
                    _log(f"{label} - リトライ ({attempt + 1}/3): {last_error[:100]}")
                    time.sleep(5)
                else:
                    _log(f"{label} - 失敗（スキップ）: {last_error[:100]}")

        if response is None:
            continue

        # テキスト収集
        text = ""
        try:
            text = getattr(response, "text", None) or ""
        except Exception as e:
            # response.text がエラーを投げる場合のフォールバック
            try:
                parts = response.candidates[0].content.parts
                text = "".join(getattr(p, "text", "") for p in parts if getattr(p, "text", None))
            except Exception:
                _log(f"{label} - テキスト取得エラー: {str(e)[:80]}")

        if text:
            all_text_parts.append(f"## 検索{i + 1}: {label}\n\n{text}")

        # レスポンスメタデータのデバッグ検査（初回のみ詳細ログ）
        if i == 0:
            _inspect_response_metadata(response, log_callback=lambda msg: _log(msg))

        # Grounding URLを収集
        url_infos = _extract_grounding_urls(
            response,
            log_callback=lambda msg: _log(msg) if i == 0 else None,
        )
        for info in url_infos:
            url = info["url"]
            if url not in seen_urls:
                seen_urls.add(url)
                all_urls.append(url)
                all_annotations.append({
                    "source": url,
                    "title": info.get("title", ""),
                })

        # テキスト本文からURLをフォールバック抽出
        if text:
            text_urls = _extract_urls_from_text(text)
            for url in text_urls:
                if url not in seen_urls:
                    seen_urls.add(url)
                    all_urls.append(url)
                    all_annotations.append({
                        "source": url,
                        "title": "",
                    })

        _log(f"{label} 完了 - テキスト{len(text)}文字, URL{len(url_infos)}件(grounding) + テキスト内URL")

    elapsed = time.time() - start_time

    # 結果を統合
    if not all_text_parts:
        _log("検索結果が得られませんでした")
        return None

    combined_text = "\n\n---\n\n".join(all_text_parts)

    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    _log(f"検索完了 ({minutes}分{seconds}秒) - {len(all_urls)}件のソースURL検出")

    return ResearchResult(
        text=combined_text,
        source_urls=all_urls,
        annotations=all_annotations,
    )


def validate_api_key(api_key: str) -> tuple[bool, str]:
    """APIキーが有効か確認する"""
    try:
        client = create_client(api_key)
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents="Hello",
        )
        return True, "APIキーは有効です"
    except Exception as e:
        return False, f"APIキーエラー: {str(e)}"
