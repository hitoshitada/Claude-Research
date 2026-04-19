"""パイプラインランチャー - 5ステージの調査パイプラインを統合管理

各ステージの進捗を表示し、対応するアプリを起動する。
"""
import sys
import os
import json
import subprocess
import threading
from pathlib import Path
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Text, Scrollbar,
    StringVar, messagebox, filedialog,
    BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, N, S,
    WORD, END, DISABLED, NORMAL,
)
from tkinter import ttk

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "調査アウトプット"

POLL_INTERVAL_MS = 1000  # 状態ポーリング間隔（ミリ秒）

# ─── ステージ定義 ───
STAGES = [
    {
        "id": "collection",
        "label": "Stage 1  記事収集",
        "script": "search_grounding_app.py",
        "desc": "記事収集・翻訳",
    },
    {
        "id": "curation",
        "label": "Stage 2  記事選別・画像",
        "script": "article_curator.py",
        "desc": "記事採否判定＆アイキャッチ確定",
    },
    {
        "id": "generation",
        "label": "Stage 3  コンテンツ生成",
        "script": "content_generator.py",
        "desc": "WeeklyReport & ポッドキャスト原稿生成",
    },
    {
        "id": "podcast_review",
        "label": "Stage 4  ポッドキャスト",
        "script": "podcast_reviewer.py",
        "desc": "ポッドキャストレビュー＆修正",
    },
    {
        "id": "upload",
        "label": "Stage 5  WPアップロード",
        "script": "wp_uploader.py",
        "desc": "WordPressへのアップロード",
    },
]

# ─── 状態の色 ───
STATUS_COLORS = {
    "pending":    "#9E9E9E",  # グレー
    "in_progress": "#FF9800", # オレンジ
    "completed":  "#4CAF50",  # 緑
    "error":      "#F44336",  # 赤
    "unreviewed": "#9E9E9E",
    "reviewed":   "#4CAF50",
}

STATUS_LABELS = {
    "pending":     "待機中",
    "in_progress": "進行中",
    "completed":   "完了",
    "error":       "エラー",
    "unreviewed":  "未レビュー",
    "reviewed":    "レビュー済",
}


def load_pipeline_state(folder: Path) -> dict:
    state_file = folder / "_pipeline_state.json"
    if state_file.exists():
        try:
            return json.loads(state_file.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def get_stage_display(stage_id: str, state: dict, folder: Path | None) -> tuple[str, str]:
    """(表示テキスト, ステータスキー) を返す"""
    if not state or not folder:
        return "待機中", "pending"

    stages = state.get("stages", {})
    st = stages.get(stage_id, {})
    status = st.get("status", "pending")

    # ステージ固有の追加情報
    if stage_id == "collection":
        count = st.get("article_count", 0)
        if status == "completed" and count > 0:
            return f"完了 {count}件", status
    elif stage_id == "curation":
        adopted = st.get("adopted_count", 0)
        rejected = st.get("rejected_count", 0)
        total_in_folder = 0
        if folder:
            import re
            total_in_folder = len([f for f in folder.glob("*.html")
                                    if re.match(r"^\d+", f.name)])
            rejected_count = len([f for f in folder.glob("不採用_*.html")])
            total_in_folder += rejected_count
        processed = adopted + rejected
        if status == "completed":
            return f"完了 採用{adopted}/不採用{rejected}", status
        elif status == "in_progress" and total_in_folder > 0:
            return f"進行中 {processed}/{total_in_folder}", status
    elif stage_id == "generation":
        has_pdf = bool(st.get("pdf_path"))
        has_script = bool(st.get("script_path"))
        if status == "completed":
            return "完了 PDF+原稿", status
        elif has_pdf and has_script:
            return "完了 PDF+原稿", "completed"
        elif has_pdf or has_script:
            return "一部完了", "in_progress"
    elif stage_id == "podcast_review":
        review_count = st.get("review_count", 0)
        last_pos = st.get("last_position_sec", 0.0)
        if status == "reviewed":
            return f"レビュー済 ({review_count}回)", "completed"
        elif status == "in_progress":
            mins, secs = int(last_pos) // 60, int(last_pos) % 60
            return f"レビュー中 {mins:02d}:{secs:02d}", "in_progress"
        else:
            return "未レビュー", "pending"
    elif stage_id == "upload":
        uploaded = st.get("uploaded_count", 0)
        if status == "completed":
            return f"完了 {uploaded}件", status
        elif status == "in_progress" and uploaded > 0:
            return f"進行中 {uploaded}件", status

    label = STATUS_LABELS.get(status, status)
    return label, status


# ─── メインGUIアプリ ───

class PipelineLauncherApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("AI調査パイプライン ランチャー")
        self.root.geometry("900x600")
        self.root.resizable(True, True)
        self.root.configure(bg="#263238")

        self.folder: Path | None = None
        self.pipeline_state: dict = {}
        self._stage_widgets: list[dict] = []
        self._processes: dict[str, subprocess.Popen] = {}
        self._poll_job = None
        # コンボ表示名（"〇 フォルダ名" など）→ 実際のフォルダ名 のマッピング
        self._display_to_folder: dict[str, str] = {}

        self._build_ui()
        self._populate_folder_list()
        self._start_polling()

    def _build_ui(self):
        root = self.root

        # ===== タイトル =====
        title_frame = Frame(root, bg="#1565C0", pady=10)
        title_frame.pack(fill=X)
        Label(title_frame, text="AI調査パイプライン ランチャー",
              font=("Yu Gothic UI", 16, "bold"), fg="white", bg="#1565C0").pack()

        # ===== フォルダ選択バー =====
        folder_bar = Frame(root, bg="#37474F", pady=6, padx=12)
        folder_bar.pack(fill=X)

        Label(folder_bar, text="フォルダ:", font=("Yu Gothic UI", 10),
              bg="#37474F", fg="#CFD8DC").pack(side=LEFT)

        # フォルダコンボボックス
        self.folder_var = StringVar()
        self.folder_combo = ttk.Combobox(folder_bar, textvariable=self.folder_var,
                                          width=35, font=("Yu Gothic UI", 10))
        self.folder_combo.pack(side=LEFT, padx=6)
        self.folder_combo.bind("<<ComboboxSelected>>", self._on_folder_combo_select)

        Button(folder_bar, text="フォルダ選択", font=("Yu Gothic UI", 10),
               bg="#1565C0", fg="white",
               command=self._select_folder).pack(side=LEFT, padx=4)

        Button(folder_bar, text="新規フォルダ作成", font=("Yu Gothic UI", 10),
               bg="#4CAF50", fg="white",
               command=self._create_new_folder).pack(side=LEFT, padx=4)

        self.folder_status = Label(folder_bar, text="",
                                    font=("Yu Gothic UI", 9), bg="#37474F", fg="#90A4AE")
        self.folder_status.pack(side=LEFT, padx=8)

        # ===== ステージ一覧 =====
        stages_frame = Frame(root, bg="#37474F", pady=6, padx=12)
        stages_frame.pack(fill=X, padx=8, pady=4)

        # ヘッダー
        header_row = Frame(stages_frame, bg="#455A64")
        header_row.pack(fill=X, pady=(0, 2))
        Label(header_row, text="ステージ", font=("Yu Gothic UI", 10, "bold"),
              bg="#455A64", fg="white", width=24, anchor=W).grid(row=0, column=0, padx=4, pady=2)
        Label(header_row, text="状態", font=("Yu Gothic UI", 10, "bold"),
              bg="#455A64", fg="white", width=22, anchor=W).grid(row=0, column=1, padx=4)
        Label(header_row, text="説明", font=("Yu Gothic UI", 10, "bold"),
              bg="#455A64", fg="white", width=24, anchor=W).grid(row=0, column=2, padx=4)
        Label(header_row, text="操作", font=("Yu Gothic UI", 10, "bold"),
              bg="#455A64", fg="white", width=8, anchor=W).grid(row=0, column=3, padx=4)

        self._stage_widgets = []
        for i, stage in enumerate(STAGES):
            bg = "#37474F" if i % 2 == 0 else "#3E515A"
            row_frame = Frame(stages_frame, bg=bg)
            row_frame.pack(fill=X)

            # ステージ名
            Label(row_frame, text=stage["label"], font=("Yu Gothic UI", 10, "bold"),
                  bg=bg, fg="#E0E0E0", width=24, anchor=W).grid(row=0, column=0, padx=4, pady=4)

            # 状態ラベル
            status_var = StringVar(value="待機中")
            status_label = Label(row_frame, textvariable=status_var,
                                  font=("Yu Gothic UI", 10), bg=bg, fg="#9E9E9E",
                                  width=22, anchor=W)
            status_label.grid(row=0, column=1, padx=4)

            # 説明
            Label(row_frame, text=stage["desc"], font=("Yu Gothic UI", 9),
                  bg=bg, fg="#B0BEC5", width=24, anchor=W).grid(row=0, column=2, padx=4)

            # 起動ボタン
            btn = Button(row_frame, text="起動 ▶",
                         font=("Yu Gothic UI", 10, "bold"),
                         bg="#1565C0", fg="white", width=8,
                         command=lambda s=stage: self._launch_stage(s))
            btn.grid(row=0, column=3, padx=8, pady=3)

            self._stage_widgets.append({
                "stage": stage,
                "status_var": status_var,
                "status_label": status_label,
                "launch_btn": btn,
                "bg": bg,
            })

        # ===== 区切り線 =====
        sep = Frame(root, bg="#546E7A", height=2)
        sep.pack(fill=X, padx=8)

        # ===== ログエリア =====
        log_frame = Frame(root, bg="#263238")
        log_frame.pack(fill=BOTH, expand=True, padx=8, pady=4)

        Label(log_frame, text="ログ", font=("Yu Gothic UI", 9, "bold"),
              bg="#263238", fg="#90A4AE").pack(anchor=W)

        text_frame = Frame(log_frame)
        text_frame.pack(fill=BOTH, expand=True)
        self.log_text = Text(text_frame, wrap=WORD, font=("Consolas", 9),
                              state=DISABLED, bg="#1e2a30", fg="#b0c4ce", height=8)
        log_scroll = Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_scroll.pack(side=RIGHT, fill=Y)

    def _log(self, msg: str, color: str = "#b0c4ce"):
        def _do():
            self.log_text.config(state=NORMAL)
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(END, f"[{ts}] {msg}\n")
            self.log_text.see(END)
            self.log_text.config(state=DISABLED)
        self.root.after(0, _do)

    # ─── フォルダ一覧 ───

    @staticmethod
    def _is_folder_complete(folder: Path) -> bool:
        """WeeklyReport PDF とポッドキャスト MP3 の両方が生成済みなら True"""
        try:
            state_file = folder / "_pipeline_state.json"
            if not state_file.exists():
                return False
            state = json.loads(state_file.read_text(encoding="utf-8"))
            gen = state.get("stages", {}).get("generation", {})
            pdf_path = gen.get("pdf_path")
            mp3_path = gen.get("mp3_path")
            return bool(pdf_path and Path(pdf_path).exists()
                        and mp3_path and Path(mp3_path).exists())
        except Exception:
            return False

    def _make_display_name(self, folder_name: str) -> str:
        """フォルダ名から表示名（完了時は '〇 ' プレフィックス付き）を返す"""
        folder = OUTPUT_DIR / folder_name
        return f"〇 {folder_name}" if self._is_folder_complete(folder) else folder_name

    def _populate_folder_list(self):
        if not OUTPUT_DIR.exists():
            return
        folder_names = sorted(
            [d.name for d in OUTPUT_DIR.iterdir()
             if d.is_dir() and not d.name.startswith("_")],
            reverse=True,
        )
        # 表示名を生成してマッピングを構築
        self._display_to_folder = {}
        display_names = []
        for name in folder_names:
            display = self._make_display_name(name)
            self._display_to_folder[display] = name
            display_names.append(display)

        self.folder_combo["values"] = display_names
        if display_names:
            self.folder_combo.set(display_names[0])
            self._load_folder(OUTPUT_DIR / folder_names[0])

    def _on_folder_combo_select(self, event=None):
        display = self.folder_var.get()
        if display:
            # 表示名（"〇 xxx" or "xxx"）から実フォルダ名を逆引きして読み込み
            actual = self._display_to_folder.get(display, display)
            self._load_folder(OUTPUT_DIR / actual)

    def _select_folder(self):
        folder = filedialog.askdirectory(
            initialdir=str(OUTPUT_DIR) if OUTPUT_DIR.exists() else str(BASE_DIR),
            title="調査アウトプットフォルダを選択",
        )
        if folder:
            self._load_folder(Path(folder))

    def _create_new_folder(self):
        from tkinter.simpledialog import askstring
        name = askstring("新規フォルダ作成", "フォルダ名を入力してください\n例: 全固体電池調査_20260417",
                         parent=self.root)
        if not name:
            return
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        new_folder = OUTPUT_DIR / name
        new_folder.mkdir(exist_ok=True)
        self._log(f"フォルダ作成: {name}")
        self._populate_folder_list()
        # 新規フォルダはまだ未完了なので 〇 なしの表示名
        self.folder_combo.set(name)
        self._load_folder(new_folder)

    def _load_folder(self, folder: Path):
        self.folder = folder
        self.pipeline_state = load_pipeline_state(folder)
        self.folder_status.config(text=folder.name)

        # このフォルダの表示名（〇付き or なし）を取得
        display = self._make_display_name(folder.name)
        # マッピングに登録されていなければ追加
        if display not in self._display_to_folder:
            self._display_to_folder[display] = folder.name
            vals = list(self.folder_combo["values"])
            vals.insert(0, display)
            self.folder_combo["values"] = vals
        self.folder_var.set(display)

        self._log(f"フォルダ読み込み: {folder.name}")
        self._update_stage_display()

    # ─── ステージ表示更新 ───

    def _update_stage_display(self):
        if not self.folder:
            return
        for widget_info in self._stage_widgets:
            stage = widget_info["stage"]
            status_var = widget_info["status_var"]
            status_label = widget_info["status_label"]

            display_text, status_key = get_stage_display(
                stage["id"], self.pipeline_state, self.folder)

            color = STATUS_COLORS.get(status_key, "#9E9E9E")
            status_var.set(display_text)
            status_label.config(fg=color)

    # ─── ポーリング ───

    def _start_polling(self):
        self._poll()

    def _poll(self):
        if self.folder and self.folder.exists():
            new_state = load_pipeline_state(self.folder)
            if new_state != self.pipeline_state:
                self.pipeline_state = new_state
                self._update_stage_display()
                # 完了状態が変わった可能性があるのでコンボの表示名を更新
                new_display = self._make_display_name(self.folder.name)
                current_display = self.folder_var.get()
                if new_display != current_display:
                    # マッピングを更新して表示名を切り替える
                    if current_display in self._display_to_folder:
                        del self._display_to_folder[current_display]
                    self._display_to_folder[new_display] = self.folder.name
                    vals = [new_display if v == current_display else v
                            for v in self.folder_combo["values"]]
                    self.folder_combo["values"] = vals
                    self.folder_var.set(new_display)

        # 終了したプロセスをクリーンアップ
        finished = [k for k, p in self._processes.items() if p.poll() is not None]
        for k in finished:
            del self._processes[k]
            self._log(f"ステージ終了: {k}")

        self._poll_job = self.root.after(POLL_INTERVAL_MS, self._poll)

    # ─── ステージ起動 ───

    def _launch_stage(self, stage: dict):
        script_path = BASE_DIR / stage["script"]
        if not script_path.exists():
            messagebox.showerror("エラー", f"スクリプトが見つかりません:\n{script_path}")
            return

        # 既に起動中か確認
        stage_id = stage["id"]
        if stage_id in self._processes and self._processes[stage_id].poll() is None:
            if not messagebox.askyesno("確認", f"{stage['label']} は既に起動中です。\n再度起動しますか？"):
                return

        args = [sys.executable, str(script_path)]
        if self.folder:
            args.append(str(self.folder))

        try:
            proc = subprocess.Popen(
                args,
                creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0,
            )
            self._processes[stage_id] = proc
            self._log(f"起動: {stage['label']} (PID:{proc.pid})")
        except Exception as e:
            messagebox.showerror("起動エラー", f"起動に失敗しました:\n{e}")
            self._log(f"起動エラー: {stage['label']} - {e}")

    def run(self):
        self.root.mainloop()
        # ポーリング停止
        if self._poll_job:
            self.root.after_cancel(self._poll_job)


if __name__ == "__main__":
    app = PipelineLauncherApp()
    app.run()
