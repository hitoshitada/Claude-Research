"""
ポッドキャスト レビュー＆修正ツール
- 生成済みポッドキャストを再生し、修正点を入力
- 修正部分だけ再生成 → 前後10秒を再生して確認
- 修正内容を修正ログに蓄積（本体のポッドキャスト生成で参照）
- 中断・再開機能、バージョン管理、レビュー完了管理
"""

import sys
import os
import re
import subprocess
import threading
import time
import shutil
import hashlib
import json
from pathlib import Path
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Entry, Listbox, Scrollbar,
    Text, messagebox, StringVar, END, DISABLED, NORMAL,
    BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, N, S,
    VERTICAL, HORIZONTAL, WORD, SINGLE,
)
from tkinter import ttk

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

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

def _smart_replace(text: str, old: str, new: str) -> str:
    """テキストを置換する。
    old が英字のみで構成されている場合は「単語境界」を考慮して置換する。
    → 前後が英字でない位置（日本語・空白・記号・文頭文末）のみ置換。
    → 例: "SEMI" → "セミ" に変換するとき "SEMICONDUCTOR" は変換しない。
    old に英字以外（日本語・数字・記号）が含まれる場合は通常の str.replace を使用。
    """
    if re.match(r'^[A-Za-z]+$', old):
        # 前後が英字でない位置だけ置換（日本語の前後、空白・記号の前後はOK）
        pattern = r'(?<![A-Za-z])' + re.escape(old) + r'(?![A-Za-z])'
        return re.sub(pattern, new, text)
    return text.replace(old, new)


def _smart_contains(text: str, word: str) -> bool:
    """テキスト中に word が含まれるか判定する。
    word が英字のみの場合は単語境界を考慮（_smart_replace と同じルール）。
    """
    if re.match(r'^[A-Za-z]+$', word):
        pattern = r'(?<![A-Za-z])' + re.escape(word) + r'(?![A-Za-z])'
        return bool(re.search(pattern, text))
    return word in text


def parse_dialogue_script(text: str) -> list[dict]:
    """F:/M:形式の対話原稿をパース。
    感情タグ付き形式 F[H]: / M[N]: にも対応。感情タグを保持する。
    """
    # F: / F[TAG]: にマッチ。感情タグをキャプチャグループで取得
    pattern = re.compile(r'^([FM])(?:\[([A-Z]+)\])?[：:]\s*(.*)')
    lines = []
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        m = pattern.match(line)
        if m:
            speaker = m.group(1)          # "F" or "M"
            emotion = m.group(2) or "N"   # "H", "E", "N" etc. デフォルトは "N"
            content = m.group(3).strip()
            lines.append({"speaker": speaker, "emotion": emotion, "text": content})
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
    """対話リストをセグメント（speaker, emotion, text）のフラットリストに展開"""
    segments = []
    for line in dialogue:
        parts = split_long_text(line["text"])
        emotion = line.get("emotion", "N")
        for part in parts:
            segments.append({"speaker": line["speaker"], "emotion": emotion, "text": part})
    return segments


def _load_voicepeak_yaml() -> dict:
    """config/voicepeak.yaml から感情プリセットを読み込む"""
    try:
        import yaml
        yaml_path = BASE_DIR / "config" / "voicepeak.yaml"
        with open(yaml_path, encoding="utf-8") as f:
            return yaml.safe_load(f)
    except Exception:
        return {}


def _get_emotion_params(speaker: str, emotion: str) -> dict:
    """感情タグに対応する speed / pitch / emotion パラメータを返す"""
    cfg = _load_voicepeak_yaml()
    presets = cfg.get("emotion_presets", {})
    speaker_presets = presets.get(speaker, {})
    params = speaker_presets.get(emotion) or speaker_presets.get("N") or {}
    return {
        "speed":   str(params.get("speed",   SPEED)),
        "pitch":   str(params.get("pitch",   0)),
        "emotion": params.get("emotion", ""),
    }


def generate_wav(text: str, narrator: str, output_path: Path,
                 speaker: str = "F", emotion: str = "N") -> bool:
    """VoicePeakでWAV生成（リトライ付き、ハング対策あり）"""
    ep = _get_emotion_params(speaker, emotion)
    cmd = [
        VOICEPEAK_EXE, "--say", text,
        "--narrator", narrator,
        "--speed", ep["speed"],
        "--pitch", ep["pitch"],
        "--out", str(output_path),
    ]
    if ep["emotion"]:
        cmd += ["--emotion", ep["emotion"]]
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

        # 直接修正ボタンバー
        direct_bar = Frame(script_frame)
        direct_bar.pack(fill=X, pady=(0, 4))
        self.btn_direct_edit = Button(
            direct_bar, text="直接修正", font=("Meiryo UI", 9),
            bg="#7B1FA2", fg="white", state=DISABLED,
            command=self._start_direct_edit)
        self.btn_direct_edit.pack(side=LEFT, padx=2)
        self.btn_test_edit = Button(
            direct_bar, text="修正のテスト", font=("Meiryo UI", 9),
            bg="#0288D1", fg="white", state=DISABLED,
            command=self._test_direct_edit)
        self.btn_test_edit.pack(side=LEFT, padx=2)
        self.btn_apply_edit = Button(
            direct_bar, text="修正の適用", font=("Meiryo UI", 9),
            bg="#388E3C", fg="white", state=DISABLED,
            command=self._apply_direct_edit)
        self.btn_apply_edit.pack(side=LEFT, padx=2)
        self.btn_cancel_edit = Button(
            direct_bar, text="修正の中止", font=("Meiryo UI", 9),
            bg="#D32F2F", fg="white", state=DISABLED,
            command=self._cancel_direct_edit)
        self.btn_cancel_edit.pack(side=LEFT, padx=2)

        # テキスト表示エリア
        text_area = Frame(script_frame)
        text_area.pack(fill=BOTH, expand=True)
        self.script_display = Text(text_area, wrap=WORD, font=("Meiryo UI", 9),
                                    height=10, state=DISABLED, bg="#fafafa")
        script_scroll = Scrollbar(text_area, orient=VERTICAL,
                                   command=self.script_display.yview)
        self.script_display.configure(yscrollcommand=script_scroll.set)
        self.script_display.pack(side=LEFT, fill=BOTH, expand=True)
        script_scroll.pack(side=RIGHT, fill=Y)

        # タグ設定（色分け）
        self.script_display.tag_configure("female", foreground="#c62828")
        self.script_display.tag_configure("male", foreground="#1565c0")
        # 再生位置ハイライトタグ（反転色）
        self.script_display.tag_configure("playing", background="#1565c0", foreground="white")
        self.script_display.tag_raise("playing")  # female/male タグより前面に

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

        # レビュー完了・バージョン管理ボタン
        review_frame = ttk.LabelFrame(root, text="レビュー管理", padding=5)
        review_frame.pack(fill=X, padx=10, pady=(0, 4))

        review_btn_row = Frame(review_frame)
        review_btn_row.pack(fill=X)

        self.btn_review_complete = Button(
            review_btn_row, text="レビュー完了 ✓", font=("Meiryo UI", 10, "bold"),
            bg="#2E7D32", fg="white", state=DISABLED,
            command=self._complete_review)
        self.btn_review_complete.pack(side=LEFT, padx=3)

        self.btn_save_version = Button(
            review_btn_row, text="バージョン保存", font=("Meiryo UI", 10),
            bg="#0277BD", fg="white", state=DISABLED,
            command=self._save_version)
        self.btn_save_version.pack(side=LEFT, padx=3)

        self.review_status_label = Label(
            review_frame, text="[未レビュー]",
            font=("Meiryo UI", 10, "bold"), fg="#9E9E9E")
        self.review_status_label.pack(side=LEFT, padx=8)

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
        self._correction_start_pos: float = 0.0  # 修正入力開始時のシーク位置

        # バックグラウンド初期化の制御
        self._init_cancel = threading.Event()
        self._wav_lock = threading.Lock()

        # 直接修正モード用
        self._direct_edit_mode: bool = False
        self._direct_edit_start_pos: float = 0.0
        self._direct_edit_backup_dialogue: list = []
        self._direct_edit_test_process: subprocess.Popen | None = None
        self._direct_edit_temp_dir: Path | None = None

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
        self.btn_direct_edit.config(state=NORMAL)

        # WAV初期化をバックグラウンドで実行（修正機能用）
        threading.Thread(target=self._initialize_wavs, daemon=True).start()

        # レビュー管理ボタンを有効化
        self.btn_review_complete.config(state=NORMAL)
        self.btn_save_version.config(state=NORMAL)

        # レビュー状態を適用（タイトル更新・シーク位置復元）
        self.root.after(200, self._apply_review_state_after_load)

    def _display_script(self):
        """原稿テキストを色分けして表示し、セグメント→文字位置マップを構築"""
        self.script_display.config(state=NORMAL)
        self.script_display.delete("1.0", END)

        # セグメント→(widgetライン番号, 文字オフセット) マッピングを構築
        self._seg_positions = []  # [(widget_line, char_col), ...]

        for line_idx, line in enumerate(self.dialogue):
            prefix = "F:" if line["speaker"] == "F" else "M:"
            tag = "female" if line["speaker"] == "F" else "male"
            self.script_display.insert(END, f"{prefix}{line['text']}\n", tag)

            # このdialogue行を分割した各セグメントの文字位置を記録
            parts = split_long_text(line["text"])
            char_offset = 0
            for part in parts:
                # Textウィジェットは1始まりの行番号
                self._seg_positions.append((line_idx + 1, len(prefix) + char_offset))
                char_offset += len(part)

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

                    if generate_wav(seg["text"], narrator, wav_path,
                                   speaker=seg["speaker"],
                                   emotion=seg.get("emotion", "N")):
                        self.seg_durations[i] = get_wav_duration(wav_path)
                    else:
                        fail_count += 1

                    self.root.after(0, lambda v=i+1: self.progress.config(value=v))

                self.is_initialized = True

                # キャッシュ無効時（スクリプト更新時）はMP3も再構築して総時間を更新
                if not cache_valid and self.mp3_path:
                    self.root.after(0, lambda: self.status_var.set(
                        "MP3を再構築中（スクリプト更新）..."))
                    mp3_ok = combine_wavs_to_mp3(
                        self.work_dir, self.mp3_path, len(self.segments))
                    if mp3_ok:
                        new_dur = get_audio_duration(self.mp3_path)
                        if new_dur > 0:
                            self.root.after(0, lambda d=new_dur: self._apply_new_duration(d))

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

    def _apply_new_duration(self, duration: float):
        """MP3総時間を更新してシークバーに反映（バックグラウンドスレッドから after(0,...) で呼ぶ）"""
        self.total_duration = duration
        try:
            self.seek_bar.config(to=duration)
            self.total_time_label.config(text=format_time(duration))
        except Exception:
            pass

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
        # ハイライトを消去
        self.script_display.tag_remove("playing", "1.0", END)

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
            # ハイライトを消去
            self.script_display.tag_remove("playing", "1.0", END)
            return

        # ドラッグ中はバーを更新しない
        if not self._seek_dragging:
            elapsed = time.time() - self.play_start_time
            current_pos = self.play_offset + elapsed
            if current_pos <= self.total_duration:
                self.seek_var.set(current_pos)
                self.time_label.config(text=format_time(current_pos))
            # 再生位置ハイライト + スクロール更新
            self._update_playback_highlight(current_pos)

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

    def _get_current_seg_index(self, current_pos: float) -> int:
        """現在の再生位置（秒）からセグメントインデックスを返す"""
        n = len(self.seg_durations)
        if n == 0 or sum(self.seg_durations) <= 0:
            return 0

        # 累積和で現在位置のセグメントを特定
        # ※ duration=0 のセグメント（WAV未生成）は累積に加わらず通過する
        cumulative = 0.0
        for i, dur in enumerate(self.seg_durations):
            cumulative += dur
            if current_pos < cumulative:
                return i

        # ループを全部通過した = current_pos が既知の合計を超えている
        # duration=0 のセグメントが多くて実際の位置が未知区間にある場合。
        # 末尾（N-1）に飛ばず、位置比率で「今いる行」を推定する。
        if self.total_duration > 0:
            ratio = min(1.0, current_pos / self.total_duration)
            return max(0, min(n - 1, int(ratio * n)))
        return max(0, n - 1)

    def _update_playback_highlight(self, current_pos: float):
        """再生位置に対応するテキストをハイライトし、中央にスクロール"""
        if not hasattr(self, '_seg_positions') or not self._seg_positions:
            return
        if not self.seg_durations:
            return

        seg_idx = self._get_current_seg_index(current_pos)
        if seg_idx >= len(self._seg_positions):
            return

        widget_line, base_col = self._seg_positions[seg_idx]

        # セグメント内のおおよその文字位置を推定
        seg_start = self._get_segment_start_time(seg_idx)
        seg_dur = self.seg_durations[seg_idx] if seg_idx < len(self.seg_durations) else 1.0
        seg_text = self.segments[seg_idx]["text"] if seg_idx < len(self.segments) else ""
        seg_len = len(seg_text)

        if seg_dur > 0 and seg_len > 0:
            # WAV duration 確定済み → 正確な経過時間で文字位置を計算
            elapsed_in_seg = max(0.0, current_pos - seg_start)
            char_in_seg = int((elapsed_in_seg / seg_dur) * seg_len)
            char_in_seg = min(char_in_seg, seg_len - 1)
        elif seg_len > 0 and self.total_duration > 0:
            # seg_dur=0（WAV未生成）: 全体比率からセグメント内の文字位置を推定
            # _get_current_seg_index の比率フォールバックと同じ計算軸を使う
            n = len(self.seg_durations)
            if n > 0:
                seg_frac_start = seg_idx / n
                seg_frac_end = (seg_idx + 1) / n
                seg_frac_range = seg_frac_end - seg_frac_start
                total_frac = current_pos / self.total_duration
                if seg_frac_range > 0:
                    within_seg = (total_frac - seg_frac_start) / seg_frac_range
                    within_seg = max(0.0, min(1.0, within_seg))
                    char_in_seg = min(int(within_seg * seg_len), seg_len - 1)
                else:
                    char_in_seg = 0
            else:
                char_in_seg = 0
        else:
            char_in_seg = 0

        hl_start_col = base_col + char_in_seg
        hl_end_col = hl_start_col + 10

        start_idx = f"{widget_line}.{hl_start_col}"
        end_idx = f"{widget_line}.{hl_end_col}"

        # playing タグを更新
        self.script_display.tag_remove("playing", "1.0", END)
        self.script_display.tag_add("playing", start_idx, end_idx)

        # 中央スクロール
        self._scroll_to_center(start_idx)

    def _scroll_to_center(self, index: str):
        """指定インデックスがテキストウィジェットの中央に来るようにスクロール"""
        # 行番号をインデックス文字列から取得（例: "5.12" → 5）
        try:
            resolved = self.script_display.index(index)
            line_num = int(resolved.split('.')[0])
        except Exception:
            return

        # テキスト全体の行数（ENDの行番号）
        try:
            total_lines = int(self.script_display.index(END).split('.')[0])
        except Exception:
            return
        if total_lines <= 1:
            return

        # この行のテキスト全体に対する位置割合（0.0〜1.0）
        # ※ 最終行はENDの1行前なので (line_num - 1) / (total_lines - 1) で正規化
        char_frac = (line_num - 1) / max(total_lines - 1, 1)

        # 現在のビュー幅
        yview_top, yview_bot = self.script_display.yview()
        view_range = yview_bot - yview_top
        if view_range <= 0:
            return

        # 対象行がビューの中央に来るよう top を設定
        new_top = char_frac - view_range / 2
        new_top = max(0.0, min(1.0 - view_range, new_top))

        self.script_display.yview_moveto(new_top)

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
            self.entry_old.focus_set()
            return

        if old_text == new_text:
            messagebox.showwarning("同一テキスト", "変更前と変更後が同じです")
            self.entry_old.focus_set()
            return

        # 表示テキストからコピーした際に含まれる "F:" "M:" "F[H]:" 等のプレフィックスを除去
        old_text = re.sub(r'^[FM](\[[A-Z]+\])?[：:]\s*', '', old_text).strip()
        if not old_text:
            messagebox.showwarning("入力不足", "変更前テキストが空です")
            self.entry_old.focus_set()
            return

        # work_dirが未作成なら作成
        if self.work_dir is None:
            self.work_dir = self.selected_folder / "_review_work"
            self.work_dir.mkdir(exist_ok=True)

        # 原稿内にテキストが存在するか確認（dialogue行で検索）
        # 英字のみの場合は単語境界を考慮（長い英単語の一部にマッチしない）
        found_in_script = False
        for line in self.dialogue:
            if _smart_contains(line["text"], old_text):
                found_in_script = True
                break

        if not found_in_script:
            messagebox.showinfo("テキスト未検出",
                f"「{old_text}」が原稿内に見つかりませんでした。\n"
                "コピーした範囲を確認して、再度入力してください。")
            # フォーカスを変更前フィールドに戻し、テキストを全選択して再入力しやすくする
            self.entry_old.focus_set()
            self.entry_old.selection_range(0, END)
            return

        # 修正開始時のシーク位置を記憶（修正OK後にここに戻る）
        self._correction_start_pos = self.seek_var.get()

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
            # 英字のみの old_text は単語境界を考慮して置換（長い英単語の一部は変えない）
            for line in self.dialogue:
                if _smart_contains(line["text"], old_text):
                    line["text"] = _smart_replace(line["text"], old_text, new_text)

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

            # 3. 変更されたセグメントを特定（テキスト・話者・感情タグの変化を検出）
            changed_indices = []
            for i in range(min(len(new_segments), len(self.segments))):
                if i >= len(self.segments):
                    changed_indices.append(i)
                elif (new_segments[i]["text"] != self.segments[i]["text"] or
                      new_segments[i].get("speaker") != self.segments[i].get("speaker") or
                      new_segments[i].get("emotion", "N") != self.segments[i].get("emotion", "N")):
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

                if generate_wav(seg["text"], narrator, wav_path,
                               speaker=seg["speaker"],
                               emotion=seg.get("emotion", "N")):
                    self.seg_durations[seg_idx] = get_wav_duration(wav_path)

            # 5. MP3を再構築
            self.root.after(0, lambda: self.status_var.set("MP3を再構築中..."))
            mp3_ok = combine_wavs_to_mp3(self.work_dir, self.mp3_path, len(self.segments))

            if log_path:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"  MP3再構築: {'OK' if mp3_ok else 'FAILED'}\n")

            # 6. 原稿ファイルも更新（感情タグを保持して書き戻す）
            updated_script = "\n".join(
                f"{line['speaker']}[{line.get('emotion', 'N')}]:{line['text']}"
                for line in self.dialogue
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
            # 初期化が途中でキャンセルされていた場合は再開（未初期化セグメントを埋める）
            self.root.after(500, self._restart_init_if_needed)

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

        # シークバーを修正開始前の位置に戻す（▶再生でその場所から続きを聴けるように）
        restore_pos = min(self._correction_start_pos, self.total_duration)
        self.seek_var.set(restore_pos)
        self.time_label.config(text=format_time(restore_pos))

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
            f"修正を確定しました（ログに保存済み）。▶再生で {format_time(restore_pos)} から続きを再生できます。")

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
                if generate_wav(seg["text"], narrator, wav_path,
                               speaker=seg["speaker"],
                               emotion=seg.get("emotion", "N")):
                    self.seg_durations[seg_idx] = get_wav_duration(wav_path)

        # MP3を再構築
        combine_wavs_to_mp3(self.work_dir, self.mp3_path, len(self.segments))

        # 原稿ファイルも復元（感情タグを保持して書き戻す）
        updated_script = "\n".join(
            f"{line['speaker']}[{line.get('emotion', 'N')}]:{line['text']}"
            for line in self.dialogue
        )
        self.script_path.write_text(updated_script, encoding="utf-8")
        self.script_text = updated_script

        # キャッシュハッシュも復元
        restored_hash = hashlib.md5(updated_script.encode("utf-8")).hexdigest()
        hash_file = self.work_dir / "_script_hash.txt"
        hash_file.write_text(restored_hash, encoding="utf-8")

        self.root.after(0, self._display_script)
        self.root.after(0, self._reset_edit_state)
        # シークバーを修正開始前の位置に戻す
        restore_pos = min(self._correction_start_pos, self.total_duration)
        self.root.after(0, lambda p=restore_pos: self.seek_var.set(p))
        self.root.after(0, lambda p=restore_pos: self.time_label.config(
            text=format_time(p)))
        self.root.after(0, lambda p=restore_pos: self.status_var.set(
            f"修正を取り消しました。▶再生で {format_time(p)} から再生できます。"))
        # 初期化が途中でキャンセルされていた場合は再開
        self.root.after(500, self._restart_init_if_needed)

    def _reset_edit_state(self):
        """編集状態をリセット"""
        self.entry_old.config(state=NORMAL)
        self.entry_new.config(state=NORMAL)
        self.btn_apply.config(state=NORMAL)
        self.btn_play.config(state=NORMAL)
        self.btn_confirm.config(state=DISABLED)
        self.btn_redo.config(state=DISABLED)

    def _restart_init_if_needed(self):
        """未初期化セグメント（duration=0）があれば _initialize_wavs を再起動する"""
        if not self.is_initialized:
            return  # まだ初回初期化が動いている
        if any(d == 0 for d in self.seg_durations):
            threading.Thread(target=self._initialize_wavs, daemon=True).start()

    # ─── 直接修正モード ───

    def _start_direct_edit(self):
        """直接修正モードに入る：再生停止→テキスト編集可能化"""
        # 現在の再生位置を保存
        if self.is_playing:
            elapsed = time.time() - self.play_start_time
            self._direct_edit_start_pos = self.play_offset + elapsed
        else:
            self._direct_edit_start_pos = self.seek_var.get()

        self._stop_playback_internal()
        self.seek_var.set(self._direct_edit_start_pos)
        self.time_label.config(text=format_time(self._direct_edit_start_pos))

        # dialogueをバックアップ
        self._direct_edit_backup_dialogue = [dict(d) for d in self.dialogue]

        # テキストウィジェットを感情タグ付きで書き換えて編集可能に
        self.script_display.config(state=NORMAL)
        self.script_display.delete("1.0", END)
        for line in self.dialogue:
            emotion = line.get("emotion", "N")
            prefix = f"{line['speaker']}[{emotion}]:"
            tag = "female" if line["speaker"] == "F" else "male"
            self.script_display.insert(END, f"{prefix}{line['text']}\n", tag)

        self._direct_edit_mode = True

        # カーソルを現在再生位置に対応する行へ移動
        seg_idx = self._get_current_seg_index(self._direct_edit_start_pos)
        if hasattr(self, '_seg_positions') and seg_idx < len(self._seg_positions):
            widget_line, _ = self._seg_positions[seg_idx]
            self.script_display.mark_set("insert", f"{widget_line}.0")
            self.script_display.see(f"{widget_line}.0")
        self.script_display.focus_set()

        # ボタン状態
        self.btn_direct_edit.config(state=DISABLED)
        self.btn_test_edit.config(state=NORMAL)
        self.btn_apply_edit.config(state=NORMAL)
        self.btn_cancel_edit.config(state=NORMAL)
        self.btn_play.config(state=DISABLED)
        self.btn_apply.config(state=DISABLED)
        self.status_var.set("直接修正モード — テキストを編集後「修正のテスト」または「修正の適用」を押してください")

    def _test_direct_edit(self):
        """修正箇所前後のテスト音声を生成・再生"""
        current_text = self.script_display.get("1.0", END)
        new_dialogue = parse_dialogue_script(current_text)
        if not new_dialogue:
            messagebox.showwarning("パースエラー", "テキストをパースできませんでした")
            return

        # 変更された行を特定（テキスト・話者・感情タグの変化を検出）
        changed = []
        n = max(len(self._direct_edit_backup_dialogue), len(new_dialogue))
        for i in range(n):
            old_l = self._direct_edit_backup_dialogue[i] if i < len(self._direct_edit_backup_dialogue) else None
            new_l = new_dialogue[i] if i < len(new_dialogue) else None
            if old_l is None or new_l is None:
                changed.append((i, old_l, new_l))
            elif (old_l["text"] != new_l["text"] or
                  old_l.get("speaker") != new_l.get("speaker") or
                  old_l.get("emotion", "N") != new_l.get("emotion", "N")):
                changed.append((i, old_l, new_l))

        if not changed:
            messagebox.showinfo("変更なし", "テキストの変更が検出されませんでした")
            return

        test_segs = self._build_test_segments(changed, new_dialogue)
        if not test_segs:
            messagebox.showwarning("テストなし", "テスト用セグメントを構築できませんでした")
            return

        self.btn_test_edit.config(state=DISABLED)
        self.status_var.set("テスト音声を生成中...")
        threading.Thread(target=self._generate_and_play_test,
                         args=(test_segs,), daemon=True).start()

    def _build_test_segments(self, changed: list, new_dialogue: list) -> list:
        """変更箇所の前後コンテキストを含むテスト用セグメントリストを構築"""
        changed_indices = [i for i, _, _ in changed]
        if not changed_indices:
            return []
        first_idx = changed_indices[0]
        last_idx = changed_indices[-1]

        result = []
        ctx_start = max(0, first_idx - 1)
        ctx_end = min(len(new_dialogue) - 1, last_idx + 1)

        for i in range(ctx_start, ctx_end + 1):
            if i >= len(new_dialogue):
                break
            line = new_dialogue[i]
            if i in changed_indices:
                old_l = next((o for idx, o, _ in changed if idx == i), None)
                snippet = self._extract_snippet_around_change(
                    old_text=old_l["text"] if old_l else "",
                    new_text=line["text"],
                    window=50,
                )
                if snippet:
                    result.append({"speaker": line["speaker"],
                                   "emotion": line.get("emotion", "N"),
                                   "text": snippet})
            else:
                # コンテキスト行：後半の短い文節だけ（前コンテキスト）か先頭（後コンテキスト）
                text = line["text"]
                if i < first_idx:
                    # 直前行：末尾30字以内
                    tail = text[-30:] if len(text) > 30 else text
                    # 文節境界で切る
                    for ch in ('。', '！', '？', '、'):
                        pos = tail.find(ch)
                        if pos >= 0:
                            tail = tail[pos + 1:]
                            break
                    if tail.strip():
                        result.append({"speaker": line["speaker"],
                                       "emotion": line.get("emotion", "N"),
                                       "text": tail.strip()})
                else:
                    # 直後行：先頭30字以内
                    head = text[:30] if len(text) > 30 else text
                    for ch in ('。', '！', '？'):
                        pos = head.find(ch)
                        if pos >= 0:
                            head = head[:pos + 1]
                            break
                    if head.strip():
                        result.append({"speaker": line["speaker"],
                                       "emotion": line.get("emotion", "N"),
                                       "text": head.strip()})
        return result

    def _extract_snippet_around_change(self, old_text: str, new_text: str,
                                        window: int = 50) -> str:
        """変更箇所の前後 ~window 字を文節境界でカットしたスニペットを返す"""
        if not new_text:
            return ""
        if len(new_text) <= window * 2:
            return new_text

        # 変更の開始位置を特定（先頭から最初に異なる文字）
        min_len = min(len(old_text), len(new_text))
        change_start = min_len
        for i in range(min_len):
            if old_text[i] != new_text[i]:
                change_start = i
                break

        center = min(change_start, len(new_text) - 1)
        raw_start = max(0, center - window)
        raw_end = min(len(new_text), center + window)

        # 開始点を文節境界（句読点の直後）に調整
        start = raw_start
        for i in range(center, raw_start - 1, -1):
            if new_text[i] in '。！？':
                start = i + 1
                break

        # 終了点を文節境界に調整
        end = raw_end
        for i in range(center, raw_end):
            if new_text[i] in '。！？':
                end = i + 1
                break

        return new_text[start:end].strip()

    def _generate_and_play_test(self, test_segs: list):
        """テスト音声をバックグラウンドで生成・再生"""
        import tempfile as _tempfile

        # 既存テスト再生を停止
        if self._direct_edit_test_process and self._direct_edit_test_process.poll() is None:
            self._direct_edit_test_process.terminate()

        # 古い一時ディレクトリをクリア
        if self._direct_edit_temp_dir and self._direct_edit_temp_dir.exists():
            shutil.rmtree(self._direct_edit_temp_dir, ignore_errors=True)

        tmp_dir = Path(_tempfile.mkdtemp(prefix="podcast_test_"))
        self._direct_edit_temp_dir = tmp_dir

        try:
            wav_names = []
            idx = 0
            for seg in test_segs:
                if not seg["text"].strip():
                    continue
                parts = split_long_text(seg["text"])
                narrator = NARRATOR_F if seg["speaker"] == "F" else NARRATOR_M
                for part in parts:
                    wav_path = tmp_dir / f"t{idx:03d}.wav"
                    if generate_wav(part, narrator, wav_path,
                                    speaker=seg["speaker"],
                                    emotion=seg.get("emotion", "N")):
                        wav_names.append(wav_path.name)
                    idx += 1

            if not wav_names:
                self.root.after(0, lambda: messagebox.showerror(
                    "生成失敗", "テスト音声の生成に失敗しました"))
                return

            if len(wav_names) == 1:
                out_path = tmp_dir / wav_names[0]
            else:
                list_file = tmp_dir / "fl.txt"
                list_file.write_text(
                    "\n".join(f"file '{n}'" for n in wav_names), encoding="utf-8")
                subprocess.run(
                    ["ffmpeg", "-y", "-f", "concat", "-safe", "0",
                     "-i", "fl.txt", "-codec:a", "libmp3lame", "-b:a", "192k",
                     "test_out.mp3"],
                    capture_output=True, cwd=str(tmp_dir), timeout=120)
                out_path = tmp_dir / "test_out.mp3"

            if not out_path.exists():
                self.root.after(0, lambda: messagebox.showerror(
                    "結合失敗", "テスト音声ファイルの作成に失敗しました"))
                return

            self._direct_edit_test_process = play_audio(out_path)
            self.root.after(0, lambda: self.status_var.set(
                "テスト再生中 ─ 「修正の適用」または「修正の中止」を選択してください"))

        except Exception as e:
            self.root.after(0, lambda err=str(e): self.status_var.set(
                f"テスト音声エラー: {err[:80]}"))
        finally:
            self.root.after(0, lambda: self.btn_test_edit.config(state=NORMAL))

    def _apply_direct_edit(self):
        """直接修正を適用"""
        # テスト再生を停止
        if self._direct_edit_test_process and self._direct_edit_test_process.poll() is None:
            self._direct_edit_test_process.terminate()

        current_text = self.script_display.get("1.0", END)
        new_dialogue = parse_dialogue_script(current_text)
        if not new_dialogue:
            messagebox.showwarning("パースエラー", "テキストをパースできませんでした")
            return

        # セグメント数が大幅に減った場合（3行以上削減）は確認ダイアログを出す
        old_count = len(self.segments)
        new_count = len(new_dialogue)
        if new_count < old_count - 2:
            diff = old_count - new_count
            answer = messagebox.askyesno(
                "行数が減少しています",
                f"編集前: {old_count}行  →  編集後: {new_count}行\n"
                f"（{diff}行が削除されています）\n\n"
                "誤って行を削除した可能性があります。\n"
                "このまま適用すると、削除された行の音声は失われます。\n\n"
                "本当に適用しますか？",
                icon="warning"
            )
            if not answer:
                self.btn_test_edit.config(state=NORMAL)
                self.btn_apply_edit.config(state=NORMAL)
                self.btn_cancel_edit.config(state=NORMAL)
                self.status_var.set("適用をキャンセルしました。テキストを確認してください。")
                return

        self.btn_test_edit.config(state=DISABLED)
        self.btn_apply_edit.config(state=DISABLED)
        self.btn_cancel_edit.config(state=DISABLED)

        threading.Thread(target=self._do_direct_apply,
                         args=(new_dialogue,), daemon=True).start()

    def _do_direct_apply(self, new_dialogue: list):
        """直接修正の適用処理（バックグラウンド）"""
        try:
            # work_dir が未作成なら作成
            if not self.work_dir:
                self.work_dir = self.selected_folder / "_review_work"
                self.work_dir.mkdir(exist_ok=True)

            old_segments = self.segments[:]
            self.dialogue = new_dialogue
            new_segments = build_segments(new_dialogue)

            # 変更されたセグメントを特定（テキスト・話者・感情タグの変化を検出）
            changed_indices = []
            for i in range(max(len(old_segments), len(new_segments))):
                if i >= len(new_segments):
                    pass
                elif i >= len(old_segments):
                    changed_indices.append(i)
                elif (old_segments[i]["text"] != new_segments[i]["text"] or
                      old_segments[i].get("speaker") != new_segments[i].get("speaker") or
                      old_segments[i].get("emotion", "N") != new_segments[i].get("emotion", "N")):
                    changed_indices.append(i)

            self.segments = new_segments

            # seg_durations をサイズ調整
            while len(self.seg_durations) < len(self.segments):
                self.seg_durations.append(3.0)
            if len(self.seg_durations) > len(self.segments):
                self.seg_durations = self.seg_durations[:len(self.segments)]

            total_gen = len(changed_indices)
            self.root.after(0, lambda t=total_gen: self.status_var.set(
                f"直接修正を適用中... 0/{t}"))

            for ci, seg_idx in enumerate(changed_indices):
                if seg_idx >= len(self.segments):
                    continue
                seg = self.segments[seg_idx]
                narrator = NARRATOR_F if seg["speaker"] == "F" else NARRATOR_M
                wav_path = self.work_dir / f"seg_{seg_idx:03d}.wav"
                self.root.after(0, lambda v=ci+1, t=total_gen: self.status_var.set(
                    f"直接修正を適用中... {v}/{t}"))
                if generate_wav(seg["text"], narrator, wav_path,
                                speaker=seg["speaker"],
                                emotion=seg.get("emotion", "N")):
                    self.seg_durations[seg_idx] = get_wav_duration(wav_path)

            # MP3再構築
            self.root.after(0, lambda: self.status_var.set("MP3を再構築中..."))
            combine_wavs_to_mp3(self.work_dir, self.mp3_path, len(self.segments))

            # 原稿ファイル更新
            updated = "\n".join(
                f"{d['speaker']}[{d.get('emotion','N')}]:{d['text']}"
                for d in self.dialogue
            ) + "\n"
            self.script_text = updated
            if self.script_path:
                self.script_path.write_text(updated, encoding="utf-8")
                new_hash = hashlib.md5(updated.encode("utf-8")).hexdigest()
                if self.work_dir:
                    (self.work_dir / "_script_hash.txt").write_text(new_hash, encoding="utf-8")

            self.root.after(0, self._finish_direct_edit_mode)
            self.root.after(0, lambda tc=total_gen: self.status_var.set(
                f"直接修正を適用しました（{tc}セグメント再生成）"))
            # MP3総時間を再取得してシークバーを更新し、修正開始位置から再生再開
            self.root.after(200, self._resume_after_direct_apply)
            # 初期化が途中でキャンセルされていた場合は再開
            self.root.after(700, self._restart_init_if_needed)

        except Exception as e:
            self.root.after(0, lambda err=str(e): self.status_var.set(
                f"直接修正エラー: {err[:80]}"))
            self.root.after(0, self._finish_direct_edit_mode)

    def _cancel_direct_edit(self):
        """直接修正を中止 — テキストを元に戻し中断位置から再生再開"""
        # テスト再生を停止
        if self._direct_edit_test_process and self._direct_edit_test_process.poll() is None:
            self._direct_edit_test_process.terminate()

        # dialogue をバックアップから復元
        self.dialogue = [dict(d) for d in self._direct_edit_backup_dialogue]
        self.segments = build_segments(self.dialogue)

        self._finish_direct_edit_mode()

        # 直接修正を開始した位置から再生再開
        self._play_from(self._direct_edit_start_pos)
        self.status_var.set(
            f"修正を中止しました。{format_time(self._direct_edit_start_pos)} から再生を再開します")

    def _finish_direct_edit_mode(self):
        """直接修正モードを終了してUIをリセット"""
        self._direct_edit_mode = False

        # 一時ディレクトリをクリーンアップ
        if self._direct_edit_temp_dir and self._direct_edit_temp_dir.exists():
            shutil.rmtree(self._direct_edit_temp_dir, ignore_errors=True)
        self._direct_edit_temp_dir = None

        # テキストウィジェットを通常表示（色分け・感情タグなし）に戻す
        self._display_script()

        # ボタン状態を戻す
        self.btn_direct_edit.config(state=NORMAL)
        self.btn_test_edit.config(state=DISABLED)
        self.btn_apply_edit.config(state=DISABLED)
        self.btn_cancel_edit.config(state=DISABLED)
        self.btn_play.config(state=NORMAL)
        self.btn_apply.config(state=NORMAL)

    def _resume_after_direct_apply(self):
        """直接修正適用後にMP3時間を再取得して修正開始位置から再生再開"""
        if not self.mp3_path or not self.mp3_path.exists():
            return
        # MP3の総時間を再取得（再構築後に変化している可能性）
        new_duration = get_audio_duration(self.mp3_path)
        if new_duration > 0:
            self.total_duration = new_duration
            self.seek_bar.config(to=new_duration)
            self.total_time_label.config(text=format_time(new_duration))
        # 修正を開始した位置から再生
        resume_pos = min(self._direct_edit_start_pos, self.total_duration)
        self._play_from(resume_pos)
        self.status_var.set(
            f"直接修正を適用しました。{format_time(resume_pos)} から再生を再開します")

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

    # ─────────────────────────────────────────────────────────────────
    # 中断・再開機能 / バージョン管理 / レビュー完了管理
    # ─────────────────────────────────────────────────────────────────

    def _get_review_state_path(self) -> Path | None:
        """レビュー状態ファイルのパスを返す"""
        if self.work_dir is None:
            if self.selected_folder:
                return self.selected_folder / "_review_work" / "_review_state.json"
        else:
            return self.work_dir / "_review_state.json"
        return None

    def _load_review_state(self) -> dict:
        """_review_state.json を読み込む。なければ初期値を返す"""
        path = self._get_review_state_path()
        if path and path.exists():
            try:
                return json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {
            "status": "unreviewed",
            "review_count": 0,
            "last_position_sec": 0.0,
            "last_reviewed": None,
            "versions": [],
        }

    def _save_review_state(self, review_state: dict):
        """_review_state.json に保存"""
        path = self._get_review_state_path()
        if path is None:
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(review_state, ensure_ascii=False, indent=2), encoding="utf-8")

        # _pipeline_state.json も更新
        if self.selected_folder:
            self._update_pipeline_state(review_state)

    def _update_pipeline_state(self, review_state: dict):
        """_pipeline_state.json の podcast_review セクションを更新"""
        state_file = self.selected_folder / "_pipeline_state.json"
        try:
            if state_file.exists():
                state = json.loads(state_file.read_text(encoding="utf-8"))
            else:
                state = {"folder": self.selected_folder.name,
                         "created": datetime.now().isoformat(timespec="seconds"),
                         "stages": {}}
            pr = state.setdefault("stages", {}).setdefault("podcast_review", {})
            pr["status"] = review_state.get("status", "unreviewed")
            pr["review_count"] = review_state.get("review_count", 0)
            pr["last_position_sec"] = review_state.get("last_position_sec", 0.0)
            if review_state.get("status") == "reviewed":
                pr["completed_at"] = datetime.now().isoformat(timespec="seconds")
            state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _apply_review_state_after_load(self):
        """フォルダ読み込み後にレビュー状態を適用し、タイトルとUIを更新"""
        review_state = self._load_review_state()
        status = review_state.get("status", "unreviewed")
        review_count = review_state.get("review_count", 0)
        last_pos = review_state.get("last_position_sec", 0.0)

        # ウィンドウタイトルとステータスバーを更新
        status_text = self._get_review_status_text(status, review_count, last_pos)
        self.root.title(f"ポッドキャスト レビュー＆修正ツール  {status_text}")
        self.status_var.set(status_text)

        # 初回起動ならオリジナルバックアップを作成
        if self.mp3_path and self.mp3_path.exists():
            original_path = self.mp3_path.parent / f"{self.mp3_path.stem}_original.mp3"
            if not original_path.exists():
                shutil.copy2(self.mp3_path, original_path)
                self.status_var.set(f"原版バックアップを作成しました: {original_path.name}")

        # last_position_sec があればシークバーを設定
        if last_pos > 0 and self.total_duration > 0:
            seek_pos = min(last_pos, self.total_duration)
            self.seek_var.set(seek_pos)
            self.time_label.config(text=format_time(seek_pos))

        # レビュー完了・バージョン保存・再レビューボタンを有効化
        self._update_review_buttons(status)

    def _get_review_status_text(self, status: str, review_count: int, last_pos: float) -> str:
        if status == "reviewed":
            return f"[レビュー済 ✓ ({review_count}回)]"
        elif status == "in_progress":
            mins, secs = int(last_pos) // 60, int(last_pos) % 60
            return f"[レビュー中 - 前回: {mins}分{secs:02d}秒]"
        else:
            return "[未レビュー]"

    def _update_review_buttons(self, status: str):
        """レビュー状態に応じてボタンの表示を更新"""
        if hasattr(self, "btn_review_complete"):
            if status == "reviewed":
                self.btn_review_complete.config(text="再レビュー開始", bg="#FF6F00",
                                                 command=self._start_re_review)
            else:
                self.btn_review_complete.config(text="レビュー完了 ✓", bg="#2E7D32",
                                                 command=self._complete_review)

    def _complete_review(self):
        """レビュー完了を記録"""
        review_state = self._load_review_state()
        review_state["status"] = "reviewed"
        review_state["review_count"] = review_state.get("review_count", 0) + 1
        review_state["last_reviewed"] = datetime.now().isoformat(timespec="seconds")
        current_pos = self.seek_var.get()
        review_state["last_position_sec"] = current_pos
        self._save_review_state(review_state)

        self._update_review_buttons("reviewed")
        self.root.title(f"ポッドキャスト レビュー＆修正ツール  "
                        f"[レビュー済 ✓ ({review_state['review_count']}回)]")
        self.status_var.set(f"レビュー完了！（{review_state['review_count']}回目）"
                            f" _review_state.json に保存しました")

    def _start_re_review(self):
        """再レビューを開始（ステータスを in_progress に戻す）"""
        review_state = self._load_review_state()
        review_state["status"] = "in_progress"
        self._save_review_state(review_state)
        self._update_review_buttons("in_progress")
        self.root.title(f"ポッドキャスト レビュー＆修正ツール  [再レビュー中]")
        self.status_var.set("再レビューを開始しました")

    def _save_version(self):
        """現在のMP3をバージョンとして保存"""
        if not self.mp3_path or not self.mp3_path.exists():
            messagebox.showwarning("エラー", "MP3ファイルが見つかりません")
            return

        review_state = self._load_review_state()
        versions = review_state.get("versions", [])
        n = len(versions) + 1
        version_path = self.mp3_path.parent / f"{self.mp3_path.stem}_v{n}.mp3"

        try:
            shutil.copy2(self.mp3_path, version_path)
            versions.append({
                "version": n,
                "path": str(version_path),
                "saved_at": datetime.now().isoformat(timespec="seconds"),
            })
            review_state["versions"] = versions
            self._save_review_state(review_state)
            self.status_var.set(f"バージョン v{n} を保存しました: {version_path.name}")
        except Exception as e:
            messagebox.showerror("エラー", f"バージョン保存に失敗しました:\n{e}")

    def _save_position_and_close(self):
        """現在位置を保存して閉じる"""
        if not self.selected_folder:
            self._finish()
            return

        if not messagebox.askyesno("終了確認", "現在位置を保存して閉じますか？"):
            return

        # 現在位置を保存
        review_state = self._load_review_state()
        current_pos = self.seek_var.get()
        review_state["last_position_sec"] = current_pos
        if review_state.get("status") == "unreviewed":
            review_state["status"] = "in_progress"
        self._save_review_state(review_state)

        self._stop_playback()
        # 作業ディレクトリの削除確認
        if self.work_dir and self.work_dir.exists():
            if messagebox.askyesno("WAV削除確認",
                    "作業用WAVファイルを削除しますか？\n"
                    "（「いいえ」を選ぶと次回の読み込みが高速になります）"):
                shutil.rmtree(self.work_dir, ignore_errors=True)

        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    import sys as _sys

    # フォルダ引数の処理
    initial_folder = None
    if len(_sys.argv) > 1:
        p = Path(_sys.argv[1])
        if p.is_dir():
            initial_folder = p

    app = PodcastReviewerApp()

    if initial_folder:
        # フォルダ一覧からマッチするものを選択
        OUTPUT_DIR_CHECK = Path(r"C:\Users\hitos\OneDrive\AI関連\DeepResearchをつかった情報調査\調査アウトプット")
        if initial_folder.parent == OUTPUT_DIR_CHECK:
            # リストボックスから選択
            folders = list(app.folder_listbox.get(0, END))
            if initial_folder.name in folders:
                idx = folders.index(initial_folder.name)
                app.folder_listbox.selection_set(idx)
                app.folder_listbox.see(idx)
                app.root.after(200, app._on_folder_select)

    # ウィンドウを閉じる際に現在位置を保存
    def _on_close():
        if app.mp3_path and app.selected_folder:
            app._save_position_and_close()
        else:
            app.root.destroy()

    app.root.protocol("WM_DELETE_WINDOW", _on_close)

    app.run()
