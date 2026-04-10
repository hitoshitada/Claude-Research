"""
ポッドキャスト レビュー＆修正ツール
- 生成済みポッドキャストを再生し、修正点を入力
- 修正部分だけ再生成 → 前後10秒を再生して確認
- 修正内容を修正ログに蓄積（本体のポッドキャスト生成で参照）
"""

import os
import re
import subprocess
import threading
import time
import shutil
import hashlib
from pathlib import Path
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Entry, Listbox, Scrollbar,
    Text, messagebox, StringVar, END, DISABLED, NORMAL,
    BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, N, S,
    VERTICAL, HORIZONTAL, WORD, SINGLE,
)
from tkinter import ttk

# ─── 定数 ───
BASE_DIR = Path(r"C:\Users\hitos\OneDrive\AI関連\DeepResearchをつかった情報調査")
OUTPUT_DIR = BASE_DIR / "調査アウトプット"
CORRECTIONS_LOG = BASE_DIR / "ポッドキャスト修正ログ.txt"

VOICEPEAK_EXE = r"C:\Program Files\VOICEPEAK\voicepeak.exe"
NARRATOR_F = "Japanese Female 1"
NARRATOR_M = "Japanese Male 1"
SPEED = 100
MAX_CHARS = 140
MAX_RETRIES = 3
RETRY_DELAY = 2.0


# ─── 共通ユーティリティ（podcast_generator.py と同じロジック） ───

def parse_dialogue_script(text: str) -> list[dict]:
    """F:/M:形式の対話原稿をパース"""
    lines = []
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("F:") or line.startswith("F："):
            lines.append({"speaker": "F", "text": line[2:].strip()})
        elif line.startswith("M:") or line.startswith("M："):
            lines.append({"speaker": "M", "text": line[2:].strip()})
        else:
            if lines:
                lines[-1]["text"] += line
    return lines


def split_long_text(text: str, max_chars: int = MAX_CHARS) -> list[str]:
    """長いテキストをmax_chars以内に分割"""
    if len(text) <= max_chars:
        return [text]
    segments = []
    sentences = re.split(r'(?<=[。！？])', text)
    current = ""
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        if len(s) > max_chars:
            if current:
                segments.append(current)
                current = ""
            parts = re.split(r'(?<=[、，])', s)
            for p in parts:
                if len(current) + len(p) <= max_chars:
                    current += p
                else:
                    if current:
                        segments.append(current)
                    current = p
        elif len(current) + len(s) <= max_chars:
            current += s
        else:
            if current:
                segments.append(current)
            current = s
    if current:
        segments.append(current)
    return segments


def build_segments(dialogue: list[dict]) -> list[dict]:
    """対話リストをセグメント（speaker, text）のフラットリストに展開"""
    segments = []
    for line in dialogue:
        parts = split_long_text(line["text"])
        for part in parts:
            segments.append({"speaker": line["speaker"], "text": part})
    return segments


def generate_wav(text: str, narrator: str, output_path: Path) -> bool:
    """VoicePeakでWAV生成（リトライ付き、ハング対策あり）"""
    cmd = [
        VOICEPEAK_EXE, "--say", text,
        "--narrator", narrator,
        "--speed", str(SPEED),
        "--out", str(output_path),
    ]
    for attempt in range(MAX_RETRIES):
        try:
            result = subprocess.run(cmd, capture_output=True, timeout=60)
            if result.returncode == 0 and output_path.exists():
                return True
        except subprocess.TimeoutExpired:
            # VoicePeakがハングした場合は強制終了
            subprocess.run(["taskkill", "/im", "voicepeak.exe", "/f"],
                          capture_output=True)
            time.sleep(2)
        except Exception:
            pass
        if attempt < MAX_RETRIES - 1:
            time.sleep(RETRY_DELAY)
    return False


def get_wav_duration(wav_path: Path) -> float:
    """ffprobeでWAVの長さ（秒）を取得"""
    try:
        result = subprocess.run(
            ["ffprobe", "-v", "quiet", "-show_entries",
             "format=duration", "-of", "csv=p=0", str(wav_path)],
            capture_output=True, text=True, timeout=10,
        )
        return float(result.stdout.strip())
    except Exception:
        return 3.0  # デフォルト推定


def combine_wavs_to_mp3(wav_dir: Path, output_path: Path, segment_count: int) -> bool:
    """WAVファイルを番号順に結合しMP3に変換。
    日本語パスを回避するため一時ディレクトリで作業する。"""
    import tempfile

    tmp_dir = Path(tempfile.mkdtemp(prefix="podcast_combine_"))
    try:
        # WAVファイルを一時ディレクトリにコピー
        for i in range(segment_count):
            wav = wav_dir / f"seg_{i:03d}.wav"
            if wav.exists():
                shutil.copy2(wav, tmp_dir / f"seg_{i:03d}.wav")

        # filelistを相対パスで作成
        list_file = tmp_dir / "filelist.txt"
        with open(list_file, "w", encoding="utf-8") as f:
            for i in range(segment_count):
                if (tmp_dir / f"seg_{i:03d}.wav").exists():
                    f.write(f"file 'seg_{i:03d}.wav'\n")

        cmd = [
            "ffmpeg", "-y", "-f", "concat", "-safe", "0",
            "-i", "filelist.txt",
            "-codec:a", "libmp3lame", "-b:a", "192k", "-ar", "44100",
            "output.mp3",
        ]
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=300, cwd=str(tmp_dir)
        )
        tmp_out = tmp_dir / "output.mp3"
        if result.returncode == 0 and tmp_out.exists():
            shutil.copy2(tmp_out, output_path)
            return True
        return False
    except Exception:
        return False
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def get_audio_duration(file_path: Path) -> float:
    """ffprobeで音声ファイルの総再生時間（秒）を取得"""
    try:
        result = subprocess.run(
            ["ffprobe", "-v", "quiet", "-show_entries",
             "format=duration", "-of", "csv=p=0", str(file_path)],
            capture_output=True, text=True, timeout=10,
        )
        return float(result.stdout.strip())
    except Exception:
        return 0.0


def play_audio(file_path: Path, start_sec: float = 0, duration_sec: float = 0) -> subprocess.Popen:
    """ffplayで音声を再生（ノンブロッキング）"""
    cmd = ["ffplay", "-nodisp", "-autoexit"]
    if start_sec > 0:
        cmd += ["-ss", f"{start_sec:.1f}"]
    if duration_sec > 0:
        cmd += ["-t", f"{duration_sec:.1f}"]
    cmd.append(str(file_path))
    return subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def format_time(seconds: float) -> str:
    """秒数を mm:ss 形式に変換"""
    m = int(seconds) // 60
    s = int(seconds) % 60
    return f"{m:02d}:{s:02d}"


def append_correction_log(topic_name: str, old_text: str, new_text: str):
    """修正ログにエントリを追加"""
    # ファイルが存在しなければヘッダーを作成
    if not CORRECTIONS_LOG.exists():
        CORRECTIONS_LOG.write_text(
            "=== ポッドキャスト修正ログ ===\n"
            "このファイルの内容はポッドキャスト原稿生成時のプロンプトに含められます。\n"
            "同じ読み間違いや表現の問題を繰り返さないためのものです。\n\n",
            encoding="utf-8",
        )

    with open(CORRECTIONS_LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] {topic_name}\n")
        f.write(f"変更前: {old_text}\n")
        f.write(f"変更後: {new_text}\n")
        f.write("---\n")


# ─── メインGUIアプリ ───

class PodcastReviewerApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("ポッドキャスト レビュー＆修正ツール")
        self.root.geometry("900x750")
        self.root.resizable(True, True)

        # 状態変数
        self.selected_folder: Path | None = None
        self.mp3_path: Path | None = None
        self.script_path: Path | None = None
        self.script_text: str = ""
        self.dialogue: list[dict] = []
        self.segments: list[dict] = []        # {speaker, text}
        self.seg_durations: list[float] = []  # 各セグメントの秒数
        self.work_dir: Path | None = None
        self.playback_process: subprocess.Popen | None = None
        self.topic_name: str = ""
        self.is_initialized = False
        self.total_duration: float = 0.0     # MP3の総再生時間（秒）
        self.play_start_time: float = 0.0    # 再生開始した実時間
        self.play_offset: float = 0.0        # 再生開始位置（秒）
        self.play_duration: float = 0.0      # 再生する長さ（0=最後まで）
        self.is_playing: bool = False
        self._seek_dragging: bool = False     # シークバーをドラッグ中か
        self._update_job = None               # afterジョブID

        self._build_ui()

    def _build_ui(self):
        root = self.root
        root.configure(bg="#f5f5f5")

        # ===== タイトル =====
        title_frame = Frame(root, bg="#1a237e", pady=8)
        title_frame.pack(fill=X)
        Label(title_frame, text="🎙 ポッドキャスト レビュー＆修正ツール",
              font=("Meiryo UI", 14, "bold"), fg="white", bg="#1a237e").pack()

        # ===== フォルダ選択 =====
        folder_frame = ttk.LabelFrame(root, text="1. フォルダ選択", padding=8)
        folder_frame.pack(fill=X, padx=10, pady=5)

        self.folder_listbox = Listbox(folder_frame, height=5, font=("Meiryo UI", 10),
                                       selectmode=SINGLE)
        folder_scroll = Scrollbar(folder_frame, orient=VERTICAL,
                                   command=self.folder_listbox.yview)
        self.folder_listbox.configure(yscrollcommand=folder_scroll.set)
        self.folder_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        folder_scroll.pack(side=RIGHT, fill=Y)

        self._populate_folders()

        btn_frame1 = Frame(folder_frame)
        btn_frame1.pack(side=RIGHT, padx=5)
        Button(btn_frame1, text="選択して読み込み", font=("Meiryo UI", 10),
               bg="#1a237e", fg="white", command=self._on_folder_select).pack(pady=2)

        # ===== 再生コントロール =====
        play_frame = ttk.LabelFrame(root, text="2. 再生", padding=8)
        play_frame.pack(fill=X, padx=10, pady=5)

        self.file_label = Label(play_frame, text="（フォルダを選択してください）",
                                 font=("Meiryo UI", 10), anchor=W)
        self.file_label.pack(fill=X)

        btn_frame2 = Frame(play_frame)
        btn_frame2.pack(fill=X, pady=5)
        self.btn_play = Button(btn_frame2, text="▶ 再生", font=("Meiryo UI", 10),
                                bg="#4CAF50", fg="white", state=DISABLED,
                                command=self._play_full, width=8)
        self.btn_play.pack(side=LEFT, padx=3)
        self.btn_stop = Button(btn_frame2, text="■ 停止", font=("Meiryo UI", 10),
                                bg="#f44336", fg="white", state=DISABLED,
                                command=self._stop_playback, width=8)
        self.btn_stop.pack(side=LEFT, padx=3)

        # 現在時刻ラベル
        self.time_label = Label(btn_frame2, text="00:00", font=("Meiryo UI", 9),
                                 width=6, anchor=E)
        self.time_label.pack(side=LEFT, padx=(8, 2))

        # シークバー
        from tkinter import DoubleVar
        self.seek_var = DoubleVar(value=0)
        self.seek_bar = ttk.Scale(btn_frame2, from_=0, to=100,
                                   orient=HORIZONTAL, variable=self.seek_var,
                                   command=self._on_seek_move)
        self.seek_bar.pack(side=LEFT, fill=X, expand=True, padx=2)
        self.seek_bar.bind("<ButtonPress-1>", self._on_seek_press)
        self.seek_bar.bind("<ButtonRelease-1>", self._on_seek_release)

        # 総時間ラベル
        self.total_time_label = Label(btn_frame2, text="00:00",
                                       font=("Meiryo UI", 9), width=6, anchor=W)
        self.total_time_label.pack(side=LEFT, padx=(2, 0))

        # ===== 原稿表示 =====
        script_frame = ttk.LabelFrame(root, text="3. 原稿テキスト", padding=8)
        script_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        self.script_display = Text(script_frame, wrap=WORD, font=("Meiryo UI", 9),
                                    height=10, state=DISABLED, bg="#fafafa")
        script_scroll = Scrollbar(script_frame, orient=VERTICAL,
                                   command=self.script_display.yview)
        self.script_display.configure(yscrollcommand=script_scroll.set)
        self.script_display.pack(side=LEFT, fill=BOTH, expand=True)
        script_scroll.pack(side=RIGHT, fill=Y)

        # タグ設定（色分け）
        self.script_display.tag_configure("female", foreground="#c62828")
        self.script_display.tag_configure("male", foreground="#1565c0")

        # ===== 修正入力 =====
        edit_frame = ttk.LabelFrame(root, text="4. 修正入力", padding=8)
        edit_frame.pack(fill=X, padx=10, pady=5)

        row1 = Frame(edit_frame)
        row1.pack(fill=X, pady=2)
        Label(row1, text="変更前:", font=("Meiryo UI", 10), width=8, anchor=E).pack(side=LEFT)
        self.entry_old = Entry(row1, font=("Meiryo UI", 10))
        self.entry_old.pack(side=LEFT, fill=X, expand=True, padx=5)

        row2 = Frame(edit_frame)
        row2.pack(fill=X, pady=2)
        Label(row2, text="変更後:", font=("Meiryo UI", 10), width=8, anchor=E).pack(side=LEFT)
        self.entry_new = Entry(row2, font=("Meiryo UI", 10))
        self.entry_new.pack(side=LEFT, fill=X, expand=True, padx=5)

        btn_frame3 = Frame(edit_frame)
        btn_frame3.pack(fill=X, pady=5)
        self.btn_apply = Button(btn_frame3, text="修正を適用", font=("Meiryo UI", 10, "bold"),
                                 bg="#FF9800", fg="white", state=DISABLED,
                                 command=self._apply_correction)
        self.btn_apply.pack(side=LEFT, padx=3)
        self.btn_confirm = Button(btn_frame3, text="✔ 修正OK", font=("Meiryo UI", 10),
                                   bg="#4CAF50", fg="white", state=DISABLED,
                                   command=self._confirm_correction)
        self.btn_confirm.pack(side=LEFT, padx=3)
        self.btn_redo = Button(btn_frame3, text="✖ やり直し", font=("Meiryo UI", 10),
                                bg="#f44336", fg="white", state=DISABLED,
                                command=self._redo_correction)
        self.btn_redo.pack(side=LEFT, padx=3)
        self.btn_finish = Button(btn_frame3, text="修正完了・終了", font=("Meiryo UI", 10),
                                  bg="#607D8B", fg="white",
                                  command=self._finish)
        self.btn_finish.pack(side=RIGHT, padx=3)

        # ===== ステータス =====
        status_frame = ttk.LabelFrame(root, text="ステータス", padding=5)
        status_frame.pack(fill=X, padx=10, pady=(0, 10))

        self.status_var = StringVar(value="フォルダを選択してください")
        Label(status_frame, textvariable=self.status_var,
              font=("Meiryo UI", 9), anchor=W, fg="#333").pack(fill=X)

        # プログレスバー
        self.progress = ttk.Progressbar(status_frame, mode="determinate")
        self.progress.pack(fill=X, pady=3)

        # 修正前のバックアップ
        self._backup_old_text = ""
        self._backup_segments = []
        self._backup_dialogue = []
        self._correction_seg_indices = []

        # バックグラウンド初期化の制御
        self._init_cancel = threading.Event()
        self._wav_lock = threading.Lock()

    def _populate_folders(self):
        """出力フォルダ一覧を表示"""
        self.folder_listbox.delete(0, END)
        if not OUTPUT_DIR.exists():
            return
        folders = sorted(
            [d.name for d in OUTPUT_DIR.iterdir() if d.is_dir() and not d.name.startswith("_")],
            reverse=True,
        )
        for f in folders:
            self.folder_listbox.insert(END, f)

    def _on_folder_select(self):
        """フォルダ選択時の処理"""
        sel = self.folder_listbox.curselection()
        if not sel:
            messagebox.showwarning("選択なし", "フォルダを選択してください")
            return

        folder_name = self.folder_listbox.get(sel[0])
        self.selected_folder = OUTPUT_DIR / folder_name
        self.topic_name = folder_name.split("_")[0] if "_" in folder_name else folder_name

        # ポッドキャストファイルを探す
        mp3_files = list(self.selected_folder.glob("*ポッドキャスト.mp3"))
        script_files = list(self.selected_folder.glob("*ポッドキャスト原稿.txt"))

        if not mp3_files:
            messagebox.showerror("エラー", "ポッドキャストMP3ファイルが見つかりません")
            return
        if not script_files:
            messagebox.showerror("エラー", "ポッドキャスト原稿ファイルが見つかりません")
            return

        self.mp3_path = mp3_files[0]
        self.script_path = script_files[0]

        self.file_label.config(text=f"📁 {self.mp3_path.name}")
        self.status_var.set("原稿を読み込み中...")

        # MP3の総時間を取得してシークバーを初期化
        self.total_duration = get_audio_duration(self.mp3_path)
        if self.total_duration > 0:
            self.seek_bar.config(to=self.total_duration)
            self.total_time_label.config(text=format_time(self.total_duration))
        self.seek_var.set(0)
        self.time_label.config(text="00:00")

        # 原稿を読み込み・表示
        self.script_text = self.script_path.read_text(encoding="utf-8")
        self.dialogue = parse_dialogue_script(self.script_text)
        self.segments = build_segments(self.dialogue)

        self._display_script()

        # 再生ボタン・修正ボタンを即座に有効化
        self.btn_play.config(state=NORMAL)
        self.btn_stop.config(state=NORMAL)
        self.btn_apply.config(state=NORMAL)

        # WAV初期化をバックグラウンドで実行（修正機能用）
        threading.Thread(target=self._initialize_wavs, daemon=True).start()

    def _display_script(self):
        """原稿テキストを色分けして表示"""
        self.script_display.config(state=NORMAL)
        self.script_display.delete("1.0", END)
        for line in self.dialogue:
            prefix = "F:" if line["speaker"] == "F" else "M:"
            tag = "female" if line["speaker"] == "F" else "male"
            self.script_display.insert(END, f"{prefix}{line['text']}\n", tag)
        self.script_display.config(state=DISABLED)

    def _initialize_wavs(self):
        """全セグメントのWAVを生成してキャッシュ（キャンセル可能）"""
        try:
            with self._wav_lock:
                self.work_dir = self.selected_folder / "_review_work"
                self.work_dir.mkdir(exist_ok=True)

                # --- 原稿ハッシュでキャッシュの有効性を検証 ---
                script_hash = hashlib.md5(
                    self.script_text.encode("utf-8")
                ).hexdigest()
                hash_file = self.work_dir / "_script_hash.txt"
                cache_valid = False
                if hash_file.exists():
                    old_hash = hash_file.read_text(encoding="utf-8").strip()
                    cache_valid = (old_hash == script_hash)

                if not cache_valid:
                    self.root.after(0, lambda: self.status_var.set(
                        "原稿が更新されています。キャッシュをクリアして再生成します..."))
                    for old_wav in self.work_dir.glob("seg_*.wav"):
                        old_wav.unlink()
                    hash_file.write_text(script_hash, encoding="utf-8")

                total = len(self.segments)
                self.seg_durations = [0.0] * total

                self.root.after(0, lambda: self.progress.config(maximum=total, value=0))
                self.root.after(0, lambda: self.status_var.set(
                    f"音声セグメントを準備中... (0/{total})"))

                fail_count = 0
                for i, seg in enumerate(self.segments):
                    # キャンセルチェック
                    if self._init_cancel.is_set():
                        self.root.after(0, lambda: self.status_var.set(
                            "バックグラウンド初期化を中断しました"))
                        return

                    wav_path = self.work_dir / f"seg_{i:03d}.wav"

                    if wav_path.exists():
                        self.seg_durations[i] = get_wav_duration(wav_path)
                        self.root.after(0, lambda v=i+1: self.progress.config(value=v))
                        self.root.after(0, lambda v=i+1, t=total: self.status_var.set(
                            f"音声セグメント読み込み中... ({v}/{t}) ※キャッシュ使用"))
                        continue

                    narrator = NARRATOR_F if seg["speaker"] == "F" else NARRATOR_M
                    speaker = "女性" if seg["speaker"] == "F" else "男性"

                    self.root.after(0, lambda v=i+1, t=total, s=speaker, l=len(seg["text"]):
                        self.status_var.set(f"音声生成中... ({v}/{t}) {s} {l}字"))

                    if generate_wav(seg["text"], narrator, wav_path):
                        self.seg_durations[i] = get_wav_duration(wav_path)
                    else:
                        fail_count += 1

                    self.root.after(0, lambda v=i+1: self.progress.config(value=v))

                self.is_initialized = True

                if fail_count > 0:
                    self.root.after(0, lambda fc=fail_count, t=total: self.status_var.set(
                        f"修正準備完了 — {t}セグメント（{fc}件失敗あり）"))
                else:
                    self.root.after(0, lambda t=total: self.status_var.set(
                        f"修正準備完了 — {t}セグメント"))

        except Exception as e:
            self.root.after(0, lambda err=str(e): self.status_var.set(
                f"WAV初期化でエラー: {err[:80]}"))

        finally:
            # ボタンは必ず有効化（エラーやハングがあっても確実に実行）
            self.is_initialized = True
            self.root.after(0, lambda: self.btn_apply.config(state=NORMAL))
        self.root.after(0, lambda: self.progress.config(value=0))

    def _play_from(self, start_sec: float = 0, duration_sec: float = 0):
        """指定位置から再生を開始"""
        if not self.mp3_path or not self.mp3_path.exists():
            return
        self._stop_playback_internal()

        self.play_offset = start_sec
        self.play_duration = duration_sec
        self.play_start_time = time.time()
        self.is_playing = True

        self.playback_process = play_audio(self.mp3_path, start_sec, duration_sec)
        self.btn_play.config(text="⏸ 一時停止", command=self._pause)
        self.btn_stop.config(state=NORMAL)

        # シークバー更新タイマー開始
        self._start_seek_update()

    def _play_full(self):
        """MP3をシークバー位置から再生"""
        if not self.mp3_path or not self.mp3_path.exists():
            return

        # MP3の総時間を取得（初回のみ）
        if self.total_duration <= 0:
            self.total_duration = get_audio_duration(self.mp3_path)
            if self.total_duration > 0:
                self.seek_bar.config(to=self.total_duration)
                self.total_time_label.config(text=format_time(self.total_duration))

        # シークバーの現在位置から再生
        start = self.seek_var.get()
        self._play_from(start)
        self.status_var.set(f"再生中: {self.mp3_path.name} ({format_time(start)}〜)")

    def _pause(self):
        """一時停止 — 現在位置を保持"""
        if self.is_playing:
            elapsed = time.time() - self.play_start_time
            current_pos = self.play_offset + elapsed
            self._stop_playback_internal()
            self.seek_var.set(min(current_pos, self.total_duration))
            self.time_label.config(text=format_time(current_pos))
            self.btn_play.config(text="▶ 再生", command=self._play_full)
            self.status_var.set(f"一時停止: {format_time(current_pos)}")

    def _stop_playback(self):
        """停止してシークバーをリセット"""
        self._stop_playback_internal()
        self.seek_var.set(0)
        self.time_label.config(text="00:00")
        self.btn_play.config(text="▶ 再生", command=self._play_full)
        self.status_var.set("停止しました")

    def _stop_playback_internal(self):
        """再生プロセスの停止（UIリセットなし）"""
        self.is_playing = False
        if self._update_job:
            self.root.after_cancel(self._update_job)
            self._update_job = None
        if self.playback_process and self.playback_process.poll() is None:
            self.playback_process.terminate()
            try:
                self.playback_process.wait(timeout=3)
            except Exception:
                self.playback_process.kill()
        self.playback_process = None

    def _start_seek_update(self):
        """シークバーを定期更新"""
        if not self.is_playing:
            return

        # ffplayプロセスが終了したか確認
        if self.playback_process and self.playback_process.poll() is not None:
            # 再生終了
            end_pos = self.play_offset + (self.play_duration if self.play_duration > 0
                                           else self.total_duration - self.play_offset)
            self.seek_var.set(min(end_pos, self.total_duration))
            self.time_label.config(text=format_time(min(end_pos, self.total_duration)))
            self.is_playing = False
            self.btn_play.config(text="▶ 再生", command=self._play_full)
            self.status_var.set("再生完了")
            return

        # ドラッグ中はバーを更新しない
        if not self._seek_dragging:
            elapsed = time.time() - self.play_start_time
            current_pos = self.play_offset + elapsed
            if current_pos <= self.total_duration:
                self.seek_var.set(current_pos)
                self.time_label.config(text=format_time(current_pos))

        self._update_job = self.root.after(300, self._start_seek_update)

    def _on_seek_press(self, event):
        """シークバーのドラッグ開始"""
        self._seek_dragging = True

    def _on_seek_release(self, event):
        """シークバーのドラッグ終了 — その位置から再生"""
        self._seek_dragging = False
        new_pos = self.seek_var.get()
        self.time_label.config(text=format_time(new_pos))

        if self.is_playing:
            self._play_from(new_pos)
            self.status_var.set(f"再生中: {format_time(new_pos)}〜")

    def _on_seek_move(self, value):
        """シークバーの値が変わった時（ドラッグ中の表示更新）"""
        if self._seek_dragging:
            pos = float(value)
            self.time_label.config(text=format_time(pos))

    def _find_affected_segments(self, old_text: str) -> list[int]:
        """変更前テキストを含むセグメントのインデックスを返す"""
        indices = []
        for i, seg in enumerate(self.segments):
            if old_text in seg["text"]:
                indices.append(i)
        return indices

    def _get_segment_start_time(self, seg_index: int) -> float:
        """セグメントの開始時刻（秒）を計算"""
        return sum(self.seg_durations[:seg_index])

    def _apply_correction(self):
        """修正を適用"""
        old_text = self.entry_old.get().strip()
        new_text = self.entry_new.get().strip()

        if not old_text or not new_text:
            messagebox.showwarning("入力不足", "変更前と変更後の両方を入力してください")
            return

        if old_text == new_text:
            messagebox.showwarning("同一テキスト", "変更前と変更後が同じです")
            return

        # work_dirが未作成なら作成
        if self.work_dir is None:
            self.work_dir = self.selected_folder / "_review_work"
            self.work_dir.mkdir(exist_ok=True)

        # 原稿内にテキストが存在するか確認（dialogue行で検索）
        found_in_script = False
        for line in self.dialogue:
            if old_text in line["text"]:
                found_in_script = True
                break

        if not found_in_script:
            messagebox.showerror("テキスト未検出",
                f"「{old_text}」が原稿内に見つかりません。\n正確なテキストを入力してください。")
            return

        # バックアップ
        self._backup_old_text = old_text
        self._backup_segments = [dict(s) for s in self.segments]
        self._backup_dialogue = [dict(d) for d in self.dialogue]

        # ボタン状態変更
        self.btn_apply.config(state=DISABLED)
        self.btn_play.config(state=DISABLED)
        self.entry_old.config(state=DISABLED)
        self.entry_new.config(state=DISABLED)

        # バックグラウンドで修正実行
        threading.Thread(target=self._do_correction,
                         args=(old_text, new_text), daemon=True).start()

    def _do_correction(self, old_text: str, new_text: str):
        """修正処理（バックグラウンド）"""
        # バックグラウンド初期化を中断してロック取得
        self._init_cancel.set()
        self._wav_lock.acquire()
        self._init_cancel.clear()  # 次回の初期化用にリセット
        try:
            self.root.after(0, lambda: self.status_var.set("修正を適用中..."))

            # デバッグ: 修正前の状態をログ
            log_path = self.work_dir / "_debug_log.txt" if self.work_dir else None
            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"\n=== 修正開始: '{old_text}' → '{new_text}' ===\n")
                    for i, d in enumerate(self.dialogue):
                        if old_text in d["text"] or new_text in d["text"]:
                            lf.write(f"  修正前 dialogue[{i}]: {d['text'][:80]}\n")

            # 1. 対話原稿を更新
            for line in self.dialogue:
                if old_text in line["text"]:
                    line["text"] = line["text"].replace(old_text, new_text)

            # デバッグ: 修正後の状態をログ
            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    for i, d in enumerate(self.dialogue):
                        if new_text in d["text"]:
                            lf.write(f"  修正後 dialogue[{i}]: {d['text'][:80]}\n")
                    # 全dialogueのダンプ
                    lf.write("  --- 全dialogue ---\n")
                    for i, d in enumerate(self.dialogue):
                        lf.write(f"  [{i}] {d['speaker']}: {d['text'][:60]}\n")

            # 2. セグメントを再構築
            new_segments = build_segments(self.dialogue)

            # 3. 変更されたセグメントを特定
            changed_indices = []
            for i in range(min(len(new_segments), len(self.segments))):
                if i >= len(self.segments) or new_segments[i]["text"] != self.segments[i]["text"]:
                    changed_indices.append(i)
            # 新しく追加されたセグメント
            for i in range(len(self.segments), len(new_segments)):
                changed_indices.append(i)

            self.segments = new_segments
            self._correction_seg_indices = changed_indices
            total_changed = len(changed_indices)

            # seg_durationsをセグメント数に合わせる
            while len(self.seg_durations) < len(self.segments):
                self.seg_durations.append(3.0)
            if len(self.seg_durations) > len(self.segments):
                self.seg_durations = self.seg_durations[:len(self.segments)]

            # 4. WAVが無いセグメント + 変更セグメントを生成
            need_gen = set(changed_indices)
            for i, seg in enumerate(self.segments):
                wav_path = self.work_dir / f"seg_{i:03d}.wav"
                if not wav_path.exists():
                    need_gen.add(i)
            need_gen = sorted(need_gen)

            total_gen = len(need_gen)
            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"  変更セグメント: {changed_indices}\n")
                    lf.write(f"  生成必要: {need_gen} ({total_gen}個)\n")

            self.root.after(0, lambda t=total_gen: self.status_var.set(
                f"音声セグメントを生成中... (0/{t})"))

            for ci, seg_idx in enumerate(need_gen):
                seg = self.segments[seg_idx]
                narrator = NARRATOR_F if seg["speaker"] == "F" else NARRATOR_M
                wav_path = self.work_dir / f"seg_{seg_idx:03d}.wav"

                speaker = "女性" if seg["speaker"] == "F" else "男性"
                self.root.after(0, lambda v=ci+1, t=total_gen, s=speaker:
                    self.status_var.set(f"生成中... ({v}/{t}) {s}"))

                if generate_wav(seg["text"], narrator, wav_path):
                    self.seg_durations[seg_idx] = get_wav_duration(wav_path)

            # 5. MP3を再構築
            self.root.after(0, lambda: self.status_var.set("MP3を再構築中..."))
            mp3_ok = combine_wavs_to_mp3(self.work_dir, self.mp3_path, len(self.segments))

            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"  MP3再構築: {'OK' if mp3_ok else 'FAILED'}\n")

            # 6. 原稿ファイルも更新
            updated_script = "\n".join(
                f"{line['speaker']}:{line['text']}" for line in self.dialogue
            )
            self.script_path.write_text(updated_script, encoding="utf-8")
            self.script_text = updated_script

            # 6b. キャッシュハッシュも更新
            new_hash = hashlib.md5(updated_script.encode("utf-8")).hexdigest()
            hash_file = self.work_dir / "_script_hash.txt"
            hash_file.write_text(new_hash, encoding="utf-8")

            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"  原稿保存完了 hash={new_hash[:8]}\n")
                    lf.write(f"=== 修正完了 ===\n")

            # 7. 修正部分の前後10秒を再生
            if changed_indices:
                first_changed = changed_indices[0]
                start_time = self._get_segment_start_time(first_changed)
                play_start = max(0, start_time - 5)
                changed_duration = sum(
                    self.seg_durations[i] for i in changed_indices
                    if i < len(self.seg_durations)
                )
                play_duration = changed_duration + 10

                self.root.after(0, lambda: self._stop_playback())
                time.sleep(0.3)
                self.root.after(0, lambda ps=play_start, pd=play_duration:
                    self._play_correction_preview(ps, pd))

            # 8. UI更新
            self.root.after(0, self._display_script)
            self.root.after(0, lambda: self.btn_confirm.config(state=NORMAL))
            self.root.after(0, lambda: self.btn_redo.config(state=NORMAL))
            self.root.after(0, lambda tc=total_changed: self.status_var.set(
                f"修正適用完了 — {tc}セグメント再生成。確認してください。"))

        except Exception as e:
            import traceback
            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"  !!! 例外発生: {e}\n")
                    lf.write(traceback.format_exc())
            # エラー時はバックアップから復元してUIを元に戻す
            if self._backup_segments:
                self.segments = self._backup_segments
                self.dialogue = self._backup_dialogue
            self.root.after(0, lambda err=str(e): self.status_var.set(
                f"修正エラー: {err[:100]}"))
            self.root.after(0, self._reset_edit_state)
        finally:
            self._wav_lock.release()

    def _play_correction_preview(self, start_sec: float, duration_sec: float):
        """修正部分の前後を再生"""
        # MP3総時間を再取得（再構築後に変わる可能性）
        self.total_duration = get_audio_duration(self.mp3_path)
        if self.total_duration > 0:
            self.seek_bar.config(to=self.total_duration)
            self.total_time_label.config(text=format_time(self.total_duration))

        self._play_from(start_sec, duration_sec)
        self.status_var.set(
            f"修正部分を再生中（{format_time(start_sec)}〜{format_time(start_sec + duration_sec)}）")

    def _confirm_correction(self):
        """修正を確定"""
        old_text = self.entry_old.get().strip()
        new_text = self.entry_new.get().strip()

        # 修正ログに保存
        append_correction_log(self.topic_name, old_text, new_text)

        self._stop_playback()

        # 入力をクリア・状態リセット
        self.entry_old.config(state=NORMAL)
        self.entry_new.config(state=NORMAL)
        self.entry_old.delete(0, END)
        self.entry_new.delete(0, END)
        self.btn_apply.config(state=NORMAL)
        self.btn_play.config(state=NORMAL)
        self.btn_confirm.config(state=DISABLED)
        self.btn_redo.config(state=DISABLED)

        self.status_var.set(
            f"修正を確定しました（ログに保存済み）。他に修正があれば入力してください。")

    def _redo_correction(self):
        """修正を取り消し"""
        self._stop_playback()

        # バックアップから復元
        if self._backup_segments:
            self.segments = self._backup_segments
            self.dialogue = self._backup_dialogue

            # WAVを復元（変更されたセグメントを再生成）
            self.root.after(0, lambda: self.status_var.set("修正を取り消し中..."))
            threading.Thread(target=self._undo_correction, daemon=True).start()
        else:
            self._reset_edit_state()

    def _undo_correction(self):
        """取り消し処理（バックグラウンド）"""
        # 変更されたセグメントのWAVを元に戻す
        for seg_idx in self._correction_seg_indices:
            if seg_idx < len(self.segments):
                seg = self.segments[seg_idx]
                narrator = NARRATOR_F if seg["speaker"] == "F" else NARRATOR_M
                wav_path = self.work_dir / f"seg_{seg_idx:03d}.wav"
                if generate_wav(seg["text"], narrator, wav_path):
                    self.seg_durations[seg_idx] = get_wav_duration(wav_path)

        # MP3を再構築
        combine_wavs_to_mp3(self.work_dir, self.mp3_path, len(self.segments))

        # 原稿ファイルも復元
        updated_script = "\n".join(
            f"{line['speaker']}:{line['text']}" for line in self.dialogue
        )
        self.script_path.write_text(updated_script, encoding="utf-8")
        self.script_text = updated_script

        # キャッシュハッシュも復元
        restored_hash = hashlib.md5(updated_script.encode("utf-8")).hexdigest()
        hash_file = self.work_dir / "_script_hash.txt"
        hash_file.write_text(restored_hash, encoding="utf-8")

        self.root.after(0, self._display_script)
        self.root.after(0, self._reset_edit_state)
        self.root.after(0, lambda: self.status_var.set(
            "修正を取り消しました。再度入力してください。"))

    def _reset_edit_state(self):
        """編集状態をリセット"""
        self.entry_old.config(state=NORMAL)
        self.entry_new.config(state=NORMAL)
        self.btn_apply.config(state=NORMAL)
        self.btn_play.config(state=NORMAL)
        self.btn_confirm.config(state=DISABLED)
        self.btn_redo.config(state=DISABLED)

    def _finish(self):
        """修正完了・終了"""
        self._stop_playback()

        # 作業ディレクトリの削除確認
        if self.work_dir and self.work_dir.exists():
            if messagebox.askyesno("終了確認",
                    "作業用WAVファイルを削除しますか？\n"
                    "（「いいえ」を選ぶと次回の読み込みが高速になります）"):
                shutil.rmtree(self.work_dir, ignore_errors=True)

        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PodcastReviewerApp()
    app.run()
