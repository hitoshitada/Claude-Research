"""記事収集ツール (Search Grounding) - メインGUIアプリケーション"""

import sys
import os
import queue
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# プロジェクトルートをパスに追加
PROJECT_DIR = Path(__file__).parent
sys.path.insert(0, str(PROJECT_DIR))

from lib.config import INVESTIGATION_DIR, OUTPUT_DIR, ENV_FILE
from lib.worker import ResearchWorker, ProgressMessage


class SearchGroundingApp:
    """記事収集GUIアプリケーション (Search Grounding)"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("記事収集ツール (Search Grounding)")
        self.root.geometry("750x700")
        self.root.minsize(600, 550)

        # 状態
        self.message_queue: queue.Queue = queue.Queue()
        self.worker: ResearchWorker | None = None
        self.check_vars: dict[str, tk.BooleanVar] = {}

        # .envからAPIキーを読込
        self.api_key = self._load_api_key()

        # UI構築
        self._build_ui()

        # ファイルリスト更新
        self._refresh_file_list()

        # キューチェック開始
        self._check_queue()

    def _load_api_key(self) -> str:
        """APIキーを.envファイルから読み込む"""
        load_dotenv(ENV_FILE)
        return os.environ.get("GEMINI_API_KEY", "")

    def _save_api_key(self):
        """APIキーを.envファイルに保存"""
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("警告", "APIキーを入力してください")
            return

        # .envファイルに書き込み
        env_content = f"GEMINI_API_KEY={api_key}\n"
        ENV_FILE.write_text(env_content, encoding="utf-8")
        os.environ["GEMINI_API_KEY"] = api_key
        self.api_key = api_key
        messagebox.showinfo("保存完了", "APIキーを保存しました")

    def _build_ui(self):
        """UIを構築"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === APIキーセクション ===
        api_frame = ttk.LabelFrame(main_frame, text="API設定", padding=8)
        api_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(api_frame, text="Gemini API Key:").pack(side=tk.LEFT, padx=(0, 5))
        self.api_key_var = tk.StringVar(value=self.api_key)
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*", width=50)
        api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.show_key_var = tk.BooleanVar(value=False)
        def toggle_key_visibility():
            api_entry.config(show="" if self.show_key_var.get() else "*")
        ttk.Checkbutton(api_frame, text="表示", variable=self.show_key_var,
                        command=toggle_key_visibility).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(api_frame, text="保存", command=self._save_api_key, width=6).pack(side=tk.LEFT)

        # === ファイルリストセクション ===
        file_frame = ttk.LabelFrame(main_frame, text="調査ファイル一覧", padding=8)
        file_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 8))

        # ツールバー
        toolbar = ttk.Frame(file_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(toolbar, text="更新", command=self._refresh_file_list, width=6).pack(side=tk.RIGHT)
        ttk.Button(toolbar, text="全選択", command=self._select_all, width=8).pack(side=tk.RIGHT, padx=(0, 5))
        ttk.Button(toolbar, text="全解除", command=self._deselect_all, width=8).pack(side=tk.RIGHT, padx=(0, 5))

        # チェックボックスリスト用のスクロールフレーム
        canvas = tk.Canvas(file_frame, height=120)
        scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=canvas.yview)
        self.file_list_frame = ttk.Frame(canvas)

        self.file_list_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.file_list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # === プログレスセクション ===
        progress_frame = ttk.LabelFrame(main_frame, text="進捗", padding=8)
        progress_frame.pack(fill=tk.X, pady=(0, 8))

        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", length=400)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        self.status_var = tk.StringVar(value="待機中")
        ttk.Label(progress_frame, textvariable=self.status_var,
                  foreground="#555").pack(fill=tk.X)

        # === ボタンセクション（先にpackしてログに押し出されないようにする） ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(8, 0))

        self.start_button = ttk.Button(
            button_frame, text="▶ 開始", command=self._start_research, width=12
        )
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_button = ttk.Button(
            button_frame, text="■ 停止", command=self._stop_research, width=12, state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT)

        # 出力フォルダを開くボタン
        ttk.Button(
            button_frame, text="出力フォルダを開く",
            command=self._open_output_folder, width=16
        ).pack(side=tk.RIGHT)

        # コスト情報
        self.cost_var = tk.StringVar(value="")
        ttk.Label(button_frame, textvariable=self.cost_var,
                  foreground="#888").pack(side=tk.RIGHT, padx=10)

        # === ログセクション ===
        log_frame = ttk.LabelFrame(main_frame, text="ログ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=12, font=("Consolas", 9), wrap=tk.WORD, state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _refresh_file_list(self):
        """調査ファイル一覧を更新"""
        # 既存のウィジェットをクリア
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        self.check_vars.clear()

        # ファイル一覧取得
        if not INVESTIGATION_DIR.exists():
            INVESTIGATION_DIR.mkdir(parents=True, exist_ok=True)

        txt_files = sorted(INVESTIGATION_DIR.glob("*.txt"))

        if not txt_files:
            ttk.Label(self.file_list_frame,
                      text="テキストファイルがありません。\n調査内容ファイルフォルダにテキストファイルを追加してください。",
                      foreground="#999").pack(padx=10, pady=10)
            return

        for txt_file in txt_files:
            var = tk.BooleanVar(value=True)
            self.check_vars[txt_file.name] = var
            cb = ttk.Checkbutton(
                self.file_list_frame, text=txt_file.name, variable=var
            )
            cb.pack(anchor=tk.W, padx=5, pady=2)

        # コスト見積もり更新
        self._update_cost_estimate()

    def _select_all(self):
        """全てのファイルを選択"""
        for var in self.check_vars.values():
            var.set(True)
        self._update_cost_estimate()

    def _deselect_all(self):
        """全てのファイルの選択を解除"""
        for var in self.check_vars.values():
            var.set(False)
        self._update_cost_estimate()

    def _update_cost_estimate(self):
        """コスト見積もりを更新"""
        selected_count = sum(1 for v in self.check_vars.values() if v.get())
        if selected_count > 0:
            cost = round(selected_count * 0.5, 1)
            self.cost_var.set(f"推定コスト: ~${cost}")
        else:
            self.cost_var.set("")

    def _get_selected_files(self) -> list[str]:
        """選択されたファイル名のリストを取得"""
        return [name for name, var in self.check_vars.items() if var.get()]

    def _start_research(self):
        """調査を開始"""
        # バリデーション
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("警告", "APIキーを入力してください")
            return

        selected_files = self._get_selected_files()
        if not selected_files:
            messagebox.showwarning("警告", "調査するファイルを選択してください")
            return

        # 確認ダイアログ
        selected_count = len(selected_files)
        cost = round(selected_count * 0.5, 1)
        confirm = messagebox.askyesno(
            "確認",
            f"{selected_count}件の調査を開始します。\n\n"
            f"推定コスト: ~${cost}\n"
            f"推定所要時間: {selected_count * 1}~{selected_count * 3}分\n\n"
            f"実行しますか？"
        )
        if not confirm:
            return

        # 出力フォルダ確認
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        # ログクリア
        self._clear_log()
        self._log("調査を開始します...")

        # UIを更新
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(20)

        # ワーカー開始
        self.worker = ResearchWorker(api_key, selected_files, self.message_queue)
        self.worker.start()

    def _stop_research(self):
        """調査を停止"""
        if self.worker and self.worker.is_running():
            self.worker.stop()
            self._log("停止を要求しました。処理中のタスクが完了するまでお待ちください...")
            self.stop_button.config(state=tk.DISABLED)

    def _check_queue(self):
        """メッセージキューをチェックしてGUIを更新"""
        try:
            while True:
                msg: ProgressMessage = self.message_queue.get_nowait()
                self._handle_message(msg)
        except queue.Empty:
            pass

        # 100msごとにチェック
        self.root.after(100, self._check_queue)

    def _handle_message(self, msg: ProgressMessage):
        """進捗メッセージを処理"""
        if msg.msg_type == "log":
            self._log(msg.message)

        elif msg.msg_type == "status":
            self.status_var.set(msg.message)

        elif msg.msg_type == "progress":
            if msg.total > 0:
                self.progress_bar.stop()
                self.progress_bar.config(mode="determinate")
                self.progress_bar["value"] = (msg.current / msg.total) * 100
            self.status_var.set(msg.message)

        elif msg.msg_type == "error":
            self._log(f"[エラー] {msg.message}")

        elif msg.msg_type == "done":
            self._log(f"\n{msg.message}")
            self.status_var.set(msg.message)
            self.progress_bar.stop()
            self.progress_bar.config(mode="determinate")
            self.progress_bar["value"] = 100
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)

    def _log(self, message: str):
        """ログにメッセージを追加"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{timestamp} {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        """ログをクリア"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _open_output_folder(self):
        """出力フォルダをエクスプローラーで開く"""
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(str(OUTPUT_DIR))


def main():
    root = tk.Tk()

    # スタイル設定
    style = ttk.Style()
    style.theme_use("clam")

    app = SearchGroundingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
