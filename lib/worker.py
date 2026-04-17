"""バックグラウンドスレッドで記事収集パイプラインを実行するモジュール"""

import json
import threading
import queue
import time
import traceback
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Optional

from google import genai
from .config import INVESTIGATION_DIR, IMAGE_WORKERS
from .search_grounding_client import create_client, search_articles, ResearchResult
from .article_parser import (
    Article,
    resolve_redirect_urls,
    extract_articles_from_report,
)
from .image_extractor import extract_image_url
from .html_generator import (
    generate_article_html,
    make_output_folder,
    get_next_article_number,
    save_article_html,
    generate_summary_html,
    save_summary_html,
    SavedArticleInfo,
)
# generate_combined_pdf / generate_podcast / generate_weekly_report は
# Stage3（content_generator.py）で実行するため、ここではインポートしない


@dataclass
class ProgressMessage:
    """GUI に送る進捗メッセージ"""
    msg_type: str  # "log" | "progress" | "status" | "error" | "done"
    message: str = ""
    current: int = 0
    total: int = 0
    topic: str = ""


class ResearchWorker:
    """記事収集パイプラインのワーカー"""

    def __init__(
        self,
        api_key: str,
        selected_files: list[str],
        message_queue: queue.Queue,
    ):
        self.api_key = api_key
        self.selected_files = selected_files
        self.message_queue = message_queue
        self.stop_event = threading.Event()
        self.thread: Optional[threading.Thread] = None

    def start(self):
        """ワーカースレッドを開始"""
        self.stop_event.clear()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        """ワーカースレッドの停止を要求"""
        self.stop_event.set()

    def is_running(self) -> bool:
        """ワーカーが実行中かどうか"""
        return self.thread is not None and self.thread.is_alive()

    def _send(self, msg_type: str, message: str, current: int = 0, total: int = 0, topic: str = ""):
        """メッセージキューに進捗を送信"""
        self.message_queue.put(ProgressMessage(
            msg_type=msg_type,
            message=message,
            current=current,
            total=total,
            topic=topic,
        ))

    def _run(self):
        """メインパイプライン"""
        try:
            client = create_client(self.api_key)
            total_files = len(self.selected_files)

            for file_idx, filename in enumerate(self.selected_files):
                if self.stop_event.is_set():
                    self._send("log", "処理が中止されました")
                    break

                topic_name = Path(filename).stem  # 拡張子を除いたファイル名

                self._send("log", f"")
                self._send("log", f"{'='*50}")
                self._send("log", f"調査開始 ({file_idx + 1}/{total_files}): {topic_name}")
                self._send("log", f"{'='*50}")

                try:
                    self._process_topic(client, filename, topic_name)
                except Exception as e:
                    tb = traceback.format_exc()
                    self._send("error", f"エラー（{topic_name}）: {str(e)}", topic=topic_name)
                    self._send("log", f"詳細: {tb[-500:]}", topic=topic_name)
                    continue

            self._send("done", "全ての調査が完了しました")

        except Exception as e:
            tb = traceback.format_exc()
            self._send("error", f"致命的エラー: {str(e)}")
            self._send("log", f"詳細: {tb[-500:]}")
            self._send("done", "エラーにより終了しました")

    def _process_topic(self, client: genai.Client, filename: str, topic_name: str):
        """1つのトピックを処理する"""

        # Step 1: テキストファイル読込
        filepath = INVESTIGATION_DIR / filename
        topic_content = filepath.read_text(encoding="utf-8").strip()
        self._send("log", f"調査テーマ: {topic_content}", topic=topic_name)

        # Step 2: Search Groundingで記事検索
        self._send("status", "Search Grounding で記事を検索中...", topic=topic_name)

        def progress_cb(msg: str, elapsed: float):
            self._send("status", msg, topic=topic_name)
            # 重要なメッセージはログにも出力
            if "完了" in msg or "失敗" in msg or "エラー" in msg or "[DEBUG]" in msg:
                self._send("log", msg, topic=topic_name)

        def stop_check() -> bool:
            return self.stop_event.is_set()

        self._send("log", "Search Grounding (3方向検索) を開始...", topic=topic_name)

        research_result = search_articles(
            client, topic_content,
            progress_callback=progress_cb,
            stop_check=stop_check,
        )

        if research_result is None:
            self._send("error", "検索結果を取得できませんでした。上記のログを確認してください。", topic=topic_name)
            return

        self._send("log", f"レポート取得完了 ({len(research_result.text)}文字)", topic=topic_name)
        self._send("log", f"ソースURL: {len(research_result.source_urls)}件検出", topic=topic_name)

        # Step 3: ソースURL解決（リダイレクトがあれば）
        self._send("status", "ソースURLを確認中...", topic=topic_name)
        url_mapping = {}
        if research_result.source_urls:
            def url_progress(msg):
                self._send("status", msg, topic=topic_name)

            url_mapping = resolve_redirect_urls(
                research_result.source_urls,
                progress_callback=url_progress,
            )
            resolved_count = sum(1 for k, v in url_mapping.items() if k != v)
            self._send("log", f"URL確認: {len(url_mapping)}件 ({resolved_count}件リダイレクト解決)", topic=topic_name)

        # Step 4: 記事構造化+翻訳
        self._send("status", "記事を構造化・翻訳中（Gemini 2.5 Flash）...", topic=topic_name)
        self._send("log", "Gemini 2.5 Flashで記事抽出+日本語翻訳中...", topic=topic_name)
        try:
            articles = extract_articles_from_report(
                client, research_result, topic_name, url_mapping
            )
            self._send("log", f"{len(articles)}件の記事を抽出・翻訳完了", topic=topic_name)
        except Exception as e:
            self._send("error", f"記事構造化に失敗: {str(e)}", topic=topic_name)
            raise

        if not articles:
            self._send("error", "記事が見つかりませんでした", topic=topic_name)
            return

        # Step 5: 画像抽出（並列）
        self._send("status", "記事画像を抽出中...", topic=topic_name)
        self._send("log", f"各記事のURLから代表画像を取得中（並列{IMAGE_WORKERS}スレッド）...", topic=topic_name)
        total_articles = len(articles)

        def extract_image_for_article(idx_article):
            idx, article = idx_article
            if self.stop_event.is_set():
                return
            try:
                image_url = extract_image_url(article.url)
                article.image_url = image_url
            except Exception:
                article.image_url = None
            self._send(
                "progress",
                f"画像抽出: {idx + 1}/{total_articles}",
                current=idx + 1,
                total=total_articles,
                topic=topic_name,
            )

        with ThreadPoolExecutor(max_workers=IMAGE_WORKERS) as executor:
            list(executor.map(extract_image_for_article, enumerate(articles)))

        if self.stop_event.is_set():
            return

        image_count = sum(1 for a in articles if a.image_url)
        self._send("log", f"画像取得: {image_count}/{total_articles}件成功", topic=topic_name)

        # Step 6: HTMLファイル生成・保存
        self._send("status", "HTMLファイルを生成中...", topic=topic_name)
        output_folder = make_output_folder(topic_name)
        start_number = get_next_article_number(output_folder)
        collection_date = datetime.now().strftime("%Y年%m月%d日")

        self._send("log", f"出力先: {output_folder}", topic=topic_name)
        self._send("log", f"記事番号: {start_number}番から開始", topic=topic_name)

        saved_count = 0
        saved_articles_info: list[SavedArticleInfo] = []

        for idx, article in enumerate(articles):
            if self.stop_event.is_set():
                break

            article_number = start_number + idx
            try:
                html_content = generate_article_html(article, collection_date)
                filepath = save_article_html(
                    html_content, output_folder, article_number, article.title_ja
                )
                saved_count += 1

                # 概要一覧用の情報を記録
                saved_articles_info.append(SavedArticleInfo(
                    number=article_number,
                    filename=filepath.name,
                    title=article.title_ja,
                    source_name=article.source_name,
                    country=article.country,
                    publish_date=article.publish_date,
                    summary=article.summary_ja,
                ))

                self._send(
                    "progress",
                    f"HTML保存: {idx + 1}/{total_articles} ({filepath.name[:30]}...)",
                    current=idx + 1,
                    total=total_articles,
                    topic=topic_name,
                )
            except Exception as e:
                self._send("error", f"HTML保存失敗(記事{article_number}): {str(e)}", topic=topic_name)

        # Step 7: 概要一覧HTMLを生成・保存
        if saved_articles_info:
            self._send("status", "概要一覧HTMLを生成中...", topic=topic_name)
            try:
                summary_html = generate_summary_html(
                    topic_name, saved_articles_info, collection_date
                )
                summary_path = save_summary_html(summary_html, output_folder, topic_name)
                self._send("log", f"概要一覧を保存: {summary_path.name}", topic=topic_name)
            except Exception as e:
                self._send("error", f"概要一覧の保存に失敗: {str(e)}", topic=topic_name)

        # ※ WeeklyReport・ポッドキャスト生成はここでは行わない
        # → 記事選別（Stage2: article_curator.py）の後に
        #   Stage3（content_generator.py）で採用記事のみを対象に実行する

        self._send("log", f"", topic=topic_name)
        self._send(
            "log",
            f"完了: {saved_count}件のHTMLを保存 -> {output_folder.name}/",
            topic=topic_name,
        )
        self._send(
            "log",
            f"次のステップ: article_curator.py で記事を選別してください",
            topic=topic_name,
        )

    # _generate_weekly_summary は Stage3（content_generator.py）に移行
