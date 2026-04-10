"""
ポッドキャスト動画生成アプリ
- フォルダ選択 → MP3+原稿から字幕付き動画を自動生成
- VoicePeakで正確なタイミング計測
- ASS字幕（カラオケフィル演出 + タイトルカード）
- ffmpegでMP4出力
"""

import os
import re
import subprocess
import shutil
import tempfile
import time
import threading
from pathlib import Path
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Listbox, Scrollbar, Text,
    messagebox, SINGLE, DISABLED, NORMAL, WORD, X, Y, BOTH,
    LEFT, RIGHT, END, W,
)
from tkinter import ttk

# =====================================================================
# 定数
# =====================================================================
BASE_DIR = Path(r"C:\Users\hitos\OneDrive\AI関連\DeepResearchをつかった情報調査")
OUTPUT_DIR = BASE_DIR / "調査アウトプット"
IMAGE_DIR = BASE_DIR / "ポッドキャスト動画作成用画像"

VOICEPEAK_EXE = r"C:\Program Files\VOICEPEAK\voicepeak.exe"
NARRATOR_F = "Japanese Female 1"
NARRATOR_M = "Japanese Male 1"
SPEED = 100
TITLE_DURATION = 4.0  # タイトルカード表示秒数


# =====================================================================
# ユーティリティ関数
# =====================================================================
def get_audio_duration(path: Path) -> float:
    """ffprobeで音声ファイルの長さ（秒）を取得"""
    try:
        r = subprocess.run(
            ["ffprobe", "-v", "quiet", "-show_entries", "format=duration",
             "-of", "csv=p=0", str(path)],
            capture_output=True, text=True, timeout=10,
        )
        return float(r.stdout.strip())
    except Exception:
        return 0.0


def ass_time(s: float) -> str:
    """ASSの時間フォーマット: H:MM:SS.CC"""
    h = int(s) // 3600
    m = (int(s) % 3600) // 60
    sec = int(s) % 60
    cs = int((s - int(s)) * 100)
    return f"{h}:{m:02d}:{sec:02d}.{cs:02d}"


def parse_dialogue(script_text: str) -> list[tuple[str, str]]:
    """F:/M:形式の対話原稿をパースする"""
    lines = []
    for line in script_text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("F:") or line.startswith("F："):
            lines.append(("F", line[2:].strip()))
        elif line.startswith("M:") or line.startswith("M："):
            lines.append(("M", line[2:].strip()))
    return lines


def find_background_image(folder_name: str) -> Path | None:
    """フォルダ名に合う背景画像を探す。なければ最初の画像を返す。"""
    if not IMAGE_DIR.exists():
        return None

    images = list(IMAGE_DIR.glob("*.png")) + list(IMAGE_DIR.glob("*.jpg"))
    if not images:
        return None

    # フォルダ名のトピック部分（例: "半導体PLP_20260329" → "半導体PLP"）
    topic = folder_name.split("_")[0] if "_" in folder_name else folder_name

    # トピック名を含む画像を探す
    for img in images:
        if topic in img.stem:
            return img

    # 見つからなければ最初の画像を使用
    return images[0]


def remove_gemini_logo(img_path: Path, tmp_path: Path):
    """画像の右下のジェミニロゴを除去してtmp_pathに保存"""
    try:
        from PIL import Image
        img = Image.open(img_path)
        w, h = img.size
        # 右下の領域を左隣のクリーンな部分で上書き
        clean = img.crop((w - 500, h - 150, w - 250, h))
        img.paste(clean, (w - 250, h - 150))
        img.save(str(tmp_path), "PNG")
    except ImportError:
        # PILがない場合はそのままコピー
        shutil.copy2(img_path, tmp_path)


# =====================================================================
# 動画生成エンジン
# =====================================================================
def generate_video(
    mp3_path: Path,
    script_path: Path,
    image_path: Path,
    output_path: Path,
    log_callback=None,
    progress_callback=None,
) -> bool:
    """ポッドキャスト動画を生成する

    Args:
        mp3_path: ポッドキャストMP3ファイル
        script_path: ポッドキャスト原稿TXTファイル
        image_path: 背景画像ファイル
        output_path: 出力MP4ファイルパス
        log_callback: ログメッセージ用コールバック(str)
        progress_callback: 進捗用コールバック(current, total)

    Returns:
        成功時True
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    def progress(cur, total):
        if progress_callback:
            progress_callback(cur, total)

    # --- 1. 原稿を読み込み・パース ---
    log("原稿を読み込み中...")
    script_text = script_path.read_text(encoding="utf-8")
    lines = parse_dialogue(script_text)
    if not lines:
        log("エラー: 台詞が見つかりません")
        return False
    log(f"台詞数: {len(lines)}")

    # --- 2. MP3の長さを取得 ---
    total_duration = get_audio_duration(mp3_path)
    if total_duration <= 0:
        log("エラー: MP3の長さを取得できません")
        return False
    log(f"MP3長さ: {total_duration:.1f}秒")

    # --- 3. VoicePeakで各台詞の正確な尺を計測 ---
    log("VoicePeakで各台詞の正確な尺を計測中...")
    timing_dir = Path(tempfile.mkdtemp(prefix="podcast_timing_"))

    line_durations = []
    total_lines = len(lines)

    for i, (speaker, text) in enumerate(lines):
        narrator = NARRATOR_F if speaker == "F" else NARRATOR_M
        wav_path = timing_dir / f"line_{i:03d}.wav"

        cmd = [
            VOICEPEAK_EXE, "--say", text,
            "--narrator", narrator, "--speed", str(SPEED),
            "--out", str(wav_path),
        ]
        for attempt in range(3):
            try:
                result = subprocess.run(cmd, capture_output=True, timeout=120)
                if result.returncode == 0 and wav_path.exists():
                    break
            except Exception:
                pass
            if attempt < 2:
                time.sleep(1)

        if wav_path.exists():
            dur = get_audio_duration(wav_path)
        else:
            dur = len(text) / 7.0
            log(f"  WARNING: {i+1}行目のWAV生成失敗、推定値使用")

        line_durations.append(dur)
        s_label = "F" if speaker == "F" else "M"
        log(f"  [{i+1:2d}/{total_lines}] {s_label}: {dur:.2f}秒 ({len(text)}字)")
        progress(i + 1, total_lines)

    shutil.rmtree(timing_dir, ignore_errors=True)

    # --- 4. タイミング補正 ---
    wav_total = sum(line_durations)
    log(f"WAV合計: {wav_total:.1f}秒 / MP3: {total_duration:.1f}秒")
    scale = total_duration / wav_total if wav_total > 0 else 1.0

    current = 0.0
    subtitles = []
    for i, ((speaker, text), dur) in enumerate(zip(lines, line_durations)):
        scaled_dur = dur * scale
        start = current
        end = current + scaled_dur
        subtitles.append((i + 1, start, end, speaker, text))
        current = end

    # --- 5. トピック名と日付を抽出 ---
    topic_display = output_path.parent.name.split("_")[0]
    date_match = re.search(r"(\d{4})(\d{2})(\d{2})", output_path.parent.name)
    if date_match:
        y = date_match.group(1)
        m = str(int(date_match.group(2)))
        d = str(int(date_match.group(3)))
        date_display = f"{y}年{m}月{d}日"
    else:
        date_display = ""

    # --- 6. ASS字幕ファイルを生成 ---
    log("ASS字幕を生成中...")
    tmp_dir = Path(tempfile.mkdtemp(prefix="podcast_video_"))
    ass_file = tmp_dir / "subtitles.ass"

    with open(ass_file, "w", encoding="utf-8-sig") as f:
        f.write("""[Script Info]
Title: Podcast Subtitles
ScriptType: v4.00+
PlayResX: 1920
PlayResY: 1080
WrapStyle: 0
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Female,Meiryo UI,48,&H00FFFFFF,&H00707070,&H00000000,&H96000000,-1,0,0,0,100,100,0,0,1,3,2,2,60,60,60,1
Style: Male,Meiryo UI,48,&H00FFCC66,&H00805533,&H00000000,&H96000000,-1,0,0,0,100,100,0,0,1,3,2,2,60,60,60,1
Style: Title,Meiryo UI,72,&H00FFFFFF,&H00FFFFFF,&H00000000,&HC8000000,-1,0,0,0,100,100,0,0,1,4,3,5,60,60,200,1
Style: TitleDate,Meiryo UI,52,&H0080DDFF,&H0080DDFF,&H00000000,&HC8000000,-1,0,0,0,100,100,0,0,1,3,2,5,60,60,120,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
""")
        # タイトルカード
        f.write(
            f"Dialogue: 0,{ass_time(0)},{ass_time(TITLE_DURATION)},Title,,0,0,0,,"
            f"{{\\fad(500,800)}}{topic_display}ウィークリーレポート\\N\n"
        )
        f.write(
            f"Dialogue: 0,{ass_time(0.3)},{ass_time(TITLE_DURATION)},TitleDate,,0,0,0,,"
            f"{{\\fad(600,800)}}{date_display}\\N\n"
        )

        # 字幕イベント（タイトル分オフセット）
        for idx, start, end, speaker, text in subtitles:
            style = "Female" if speaker == "F" else "Male"
            adj_start = start + TITLE_DURATION
            adj_end = end + TITLE_DURATION
            duration_cs = int((end - start) * 100)
            line = f"{{\\kf{duration_cs}}}{text}"
            f.write(
                f"Dialogue: 0,{ass_time(adj_start)},{ass_time(adj_end)},"
                f"{style},,0,0,0,,{line}\n"
            )

    # --- 7. 一時ファイルを準備 ---
    tmp_img = tmp_dir / "image.png"
    tmp_mp3 = tmp_dir / "audio.mp3"
    tmp_out = tmp_dir / "output.mp4"

    # 背景画像をコピー（ジェミニロゴ除去）
    remove_gemini_logo(image_path, tmp_img)
    shutil.copy2(mp3_path, tmp_mp3)

    # --- 8. ffmpegで動画生成 ---
    log("ffmpegで動画を生成中...")
    delay_ms = int(TITLE_DURATION * 1000)

    cmd = [
        "ffmpeg", "-y",
        "-loop", "1",
        "-i", "image.png",
        "-i", "audio.mp3",
        "-vf", "scale=1920:1080,ass=subtitles.ass",
        "-af", f"adelay={delay_ms}|{delay_ms},apad",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "23",
        "-c:a", "aac",
        "-b:a", "192k",
        "-t", f"{total_duration + TITLE_DURATION + 1:.1f}",
        "-pix_fmt", "yuv420p",
        "output.mp4",
    ]

    result = subprocess.run(
        cmd, capture_output=True, timeout=600, cwd=str(tmp_dir)
    )

    if result.returncode == 0 and tmp_out.exists():
        shutil.copy2(tmp_out, output_path)
        size_mb = output_path.stat().st_size / (1024 * 1024)
        log(f"動画生成完了: {output_path.name} ({size_mb:.1f} MB)")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return True
    else:
        stderr = result.stderr.decode("utf-8", errors="replace")
        log(f"ffmpegエラー: {stderr[-500:]}")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return False


# =====================================================================
# GUI アプリケーション
# =====================================================================
class PodcastVideoMakerApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("ポッドキャスト動画生成ツール")
        self.root.geometry("800x650")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

        self.is_processing = False
        self._build_ui()

    def _build_ui(self):
        root = self.root

        # ===== タイトルバー =====
        title_frame = Frame(root, bg="#1a237e", pady=8)
        title_frame.pack(fill=X)
        Label(
            title_frame,
            text="🎬 ポッドキャスト動画生成ツール",
            font=("Meiryo UI", 14, "bold"),
            fg="white", bg="#1a237e",
        ).pack()

        # ===== フォルダ選択 =====
        folder_frame = ttk.LabelFrame(root, text="1. フォルダ選択", padding=8)
        folder_frame.pack(fill=X, padx=10, pady=5)

        list_frame = Frame(folder_frame)
        list_frame.pack(side=LEFT, fill=BOTH, expand=True)

        self.folder_listbox = Listbox(
            list_frame, height=6, font=("Meiryo UI", 10), selectmode=SINGLE,
        )
        scroll = Scrollbar(list_frame, orient="vertical",
                           command=self.folder_listbox.yview)
        self.folder_listbox.configure(yscrollcommand=scroll.set)
        self.folder_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)

        self._populate_folders()

        btn_frame = Frame(folder_frame)
        btn_frame.pack(side=RIGHT, padx=8)

        self.btn_generate = Button(
            btn_frame, text="動画を生成",
            font=("Meiryo UI", 11, "bold"),
            bg="#4CAF50", fg="white", width=14,
            command=self._on_generate,
        )
        self.btn_generate.pack(pady=3)

        self.btn_generate_all = Button(
            btn_frame, text="全フォルダを一括生成",
            font=("Meiryo UI", 9),
            bg="#1565c0", fg="white", width=14,
            command=self._on_generate_all,
        )
        self.btn_generate_all.pack(pady=3)

        # ===== 情報表示 =====
        info_frame = ttk.LabelFrame(root, text="2. ファイル情報", padding=8)
        info_frame.pack(fill=X, padx=10, pady=5)

        self.info_label = Label(
            info_frame, text="フォルダを選択してください",
            font=("Meiryo UI", 10), anchor=W, justify=LEFT,
        )
        self.info_label.pack(fill=X)

        # ===== プログレスバー =====
        progress_frame = ttk.LabelFrame(root, text="3. 進捗", padding=8)
        progress_frame.pack(fill=X, padx=10, pady=5)

        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill=X, pady=(0, 5))

        self.status_var = Label(
            progress_frame, text="待機中",
            font=("Meiryo UI", 9), anchor=W,
        )
        self.status_var.pack(fill=X)

        # ===== ログ =====
        log_frame = ttk.LabelFrame(root, text="4. ログ", padding=8)
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        self.log_text = Text(
            log_frame, wrap=WORD, font=("Consolas", 9),
            height=12, state=DISABLED, bg="#1e1e1e", fg="#d4d4d4",
        )
        log_scroll = Scrollbar(log_frame, orient="vertical",
                               command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_scroll.pack(side=RIGHT, fill=Y)

        # フォルダ選択イベント
        self.folder_listbox.bind("<<ListboxSelect>>", self._on_folder_select)

    def _populate_folders(self):
        """出力フォルダ一覧を取得"""
        if not OUTPUT_DIR.exists():
            return
        folders = sorted(
            (d.name for d in OUTPUT_DIR.iterdir()
             if d.is_dir() and not d.name.startswith("_")),
            reverse=True,
        )
        for name in folders:
            self.folder_listbox.insert(END, name)

    def _on_folder_select(self, event=None):
        """フォルダ選択時に情報を表示"""
        sel = self.folder_listbox.curselection()
        if not sel:
            return
        folder_name = self.folder_listbox.get(sel[0])
        folder = OUTPUT_DIR / folder_name

        mp3 = self._find_podcast_mp3(folder)
        script = self._find_podcast_script(folder)
        image = find_background_image(folder_name)

        info_parts = [f"フォルダ: {folder_name}"]
        info_parts.append(f"MP3: {'✓ ' + mp3.name if mp3 else '✗ 見つかりません'}")
        info_parts.append(f"原稿: {'✓ ' + script.name if script else '✗ 見つかりません'}")
        info_parts.append(f"背景画像: {'✓ ' + image.name if image else '✗ 見つかりません'}")

        # 既存MP4チェック
        mp4 = self._find_existing_mp4(folder)
        if mp4:
            size = mp4.stat().st_size / (1024 * 1024)
            info_parts.append(f"既存MP4: {mp4.name} ({size:.1f} MB) — 再生成で上書きされます")

        self.info_label.config(text="\n".join(info_parts))

    def _find_podcast_mp3(self, folder: Path) -> Path | None:
        files = list(folder.glob("*ポッドキャスト.mp3"))
        return files[0] if files else None

    def _find_podcast_script(self, folder: Path) -> Path | None:
        files = list(folder.glob("*ポッドキャスト原稿.txt"))
        return files[0] if files else None

    def _find_existing_mp4(self, folder: Path) -> Path | None:
        files = list(folder.glob("*ポッドキャスト.mp4"))
        return files[0] if files else None

    def _log(self, message: str):
        """ログに追記（メインスレッドから呼ぶ）"""
        self.log_text.config(state=NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(END, f"[{timestamp}] {message}\n")
        self.log_text.see(END)
        self.log_text.config(state=DISABLED)

    def _log_safe(self, message: str):
        """別スレッドからログ追記"""
        self.root.after(0, lambda: self._log(message))

    def _set_status(self, text: str):
        self.root.after(0, lambda: self.status_var.config(text=text))

    def _set_progress(self, current: int, total: int):
        self.root.after(0, lambda: self.progress.config(
            maximum=total, value=current))

    def _set_buttons(self, enabled: bool):
        state = NORMAL if enabled else DISABLED
        self.root.after(0, lambda: self.btn_generate.config(state=state))
        self.root.after(0, lambda: self.btn_generate_all.config(state=state))

    def _on_generate(self):
        """選択フォルダの動画を生成"""
        sel = self.folder_listbox.curselection()
        if not sel:
            messagebox.showwarning("選択なし", "フォルダを選択してください")
            return
        folder_name = self.folder_listbox.get(sel[0])
        self._start_generation([folder_name])

    def _on_generate_all(self):
        """全フォルダの動画を一括生成"""
        count = self.folder_listbox.size()
        if count == 0:
            return
        all_folders = [self.folder_listbox.get(i) for i in range(count)]

        # MP3と原稿があるフォルダのみ
        valid = []
        for name in all_folders:
            folder = OUTPUT_DIR / name
            if self._find_podcast_mp3(folder) and self._find_podcast_script(folder):
                valid.append(name)

        if not valid:
            messagebox.showinfo("対象なし", "動画生成可能なフォルダがありません")
            return

        proceed = messagebox.askyesno(
            "一括生成確認",
            f"{len(valid)}フォルダの動画を生成します。\n"
            f"（MP3+原稿があるフォルダのみ）\n\n続行しますか？",
        )
        if proceed:
            self._start_generation(valid)

    def _start_generation(self, folder_names: list[str]):
        """バックグラウンドで動画生成を実行"""
        if self.is_processing:
            return
        self.is_processing = True
        self._set_buttons(False)

        def task():
            total_folders = len(folder_names)
            success_count = 0

            for fi, folder_name in enumerate(folder_names):
                self._log_safe(f"")
                self._log_safe(f"{'='*50}")
                self._log_safe(
                    f"[{fi+1}/{total_folders}] {folder_name}"
                )
                self._log_safe(f"{'='*50}")
                self._set_status(
                    f"処理中: {folder_name} ({fi+1}/{total_folders})"
                )

                folder = OUTPUT_DIR / folder_name
                mp3 = self._find_podcast_mp3(folder)
                script = self._find_podcast_script(folder)
                image = find_background_image(folder_name)

                if not mp3:
                    self._log_safe("  スキップ: MP3が見つかりません")
                    continue
                if not script:
                    self._log_safe("  スキップ: 原稿が見つかりません")
                    continue
                if not image:
                    self._log_safe("  スキップ: 背景画像が見つかりません")
                    continue

                # 出力ファイル名: {トピック名}{日付}ポッドキャスト.mp4
                mp4_name = mp3.stem + ".mp4"  # .mp3 → .mp4
                output_path = folder / mp4_name

                try:
                    ok = generate_video(
                        mp3_path=mp3,
                        script_path=script,
                        image_path=image,
                        output_path=output_path,
                        log_callback=self._log_safe,
                        progress_callback=self._set_progress,
                    )
                    if ok:
                        success_count += 1
                except Exception as e:
                    self._log_safe(f"エラー: {e}")

            self._log_safe(f"")
            self._log_safe(f"{'='*50}")
            self._log_safe(
                f"完了: {success_count}/{total_folders}フォルダの動画を生成しました"
            )
            self._set_status(
                f"完了 — {success_count}/{total_folders}フォルダ"
            )
            self._set_progress(0, 100)
            self._set_buttons(True)
            self.is_processing = False

            self.root.after(0, lambda: messagebox.showinfo(
                "完了",
                f"{success_count}/{total_folders}フォルダの動画生成が完了しました",
            ))

        threading.Thread(target=task, daemon=True).start()

    def run(self):
        self.root.mainloop()


# =====================================================================
# エントリーポイント
# =====================================================================
def main():
    app = PodcastVideoMakerApp()
    app.run()


if __name__ == "__main__":
    main()
