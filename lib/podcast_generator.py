"""
対話形式ポッドキャスト音声生成モジュール
- 記事一覧から男女対話形式のポッドキャスト原稿を生成（Gemini API）
- VoicePeakで男女ナレーターに振り分けて音声化（140字制限対応）
- リトライ機能付き
- ffmpegでWAVを結合しMP3を出力
"""

import os
import re
import subprocess
import tempfile
import time
import shutil
import yaml
from pathlib import Path
from typing import Optional, Callable

BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__))).parent
CORRECTIONS_LOG = BASE_DIR / "ポッドキャスト修正ログ.txt"

# --- 外部設定ファイルから読み込み ---
_CONFIG_DIR = BASE_DIR / "config"
_PROMPTS_DIR = BASE_DIR / "prompts"


def _load_yaml(filename: str) -> dict:
    """config/フォルダからYAMLファイルを読み込む"""
    path = _CONFIG_DIR / filename
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _load_prompt(filename: str) -> str:
    """prompts/フォルダからプロンプトテキストを読み込む"""
    path = _PROMPTS_DIR / filename
    return path.read_text(encoding="utf-8").strip()


# VoicePeak設定をYAMLから読み込み
_vp_config = _load_yaml("voicepeak.yaml")
VOICEPEAK_EXE = _vp_config["exe_path"]
SPEED = _vp_config["speed"]
MAX_CHARS = _vp_config["max_chars"]
MAX_RETRIES = _vp_config["max_retries"]
RETRY_DELAY = _vp_config["retry_delay"]

# キャラクター設定をYAMLから読み込み
_char_config = _load_yaml("characters.yaml")
NARRATOR_F = _char_config["female"]["narrator"]
NARRATOR_M = _char_config["male"]["narrator"]

# 感情プリセットを読み込み
EMOTION_PRESETS = _vp_config.get("emotion_presets", {})


def is_voicepeak_available() -> bool:
    """VoicePeakが利用可能か確認"""
    return os.path.exists(VOICEPEAK_EXE)


def is_ffmpeg_available() -> bool:
    """ffmpegが利用可能か確認"""
    try:
        result = subprocess.run(
            ["ffmpeg", "-version"],
            capture_output=True, text=True, timeout=10
        )
        return result.returncode == 0
    except Exception:
        return False


def _load_corrections_log() -> str:
    """修正ログファイルから過去の修正履歴を読み込む"""
    if not CORRECTIONS_LOG.exists():
        return ""

    try:
        content = CORRECTIONS_LOG.read_text(encoding="utf-8")
    except Exception:
        return ""

    # 修正エントリを抽出（変更前/変更後のペア）
    corrections = []
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("変更前:"):
            old = line.replace("変更前:", "").strip()
            if i + 1 < len(lines) and lines[i + 1].strip().startswith("変更後:"):
                new = lines[i + 1].strip().replace("変更後:", "").strip()
                # 変更前テキストが英字のみの場合は「単独語として使われているときのみ」注記を追加
                # （例: SEMI → セミ のとき、SEMICONDUCTORの中のSEMIは変えない）
                import re as _re
                if _re.match(r'^[A-Za-z]+$', old):
                    corrections.append(
                        f"- 「{old}」→「{new}」"
                        f"（※英単語として単独で使われている場合のみ変更。"
                        f"「{old}」が長い英単語の一部である場合は変更しないこと）"
                    )
                else:
                    corrections.append(f"- 「{old}」→「{new}」")
                i += 2
                continue
        i += 1

    if not corrections:
        return ""

    # 重複を除去
    unique = list(dict.fromkeys(corrections))
    return "\n## 過去の修正履歴（同じ間違いを繰り返さないこと）\n" + "\n".join(unique) + "\n"


def build_dialogue_prompt(topic_name: str) -> str:
    """対話形式ポッドキャスト原稿生成用プロンプト（外部ファイルから組み立て）"""
    corrections = _load_corrections_log()

    system_prompt = _load_prompt("system.md").replace("{topic_name}", topic_name)
    structure_prompt = _load_prompt("structure.md")
    style_prompt = _load_prompt("style.md")

    return f"""{system_prompt}

{structure_prompt}

{style_prompt}
{corrections}
## 記事一覧
"""


def generate_dialogue_script(
    client,
    articles: list,
    topic_name: str,
    model: str = "gemini-2.5-flash",
) -> str:
    """Gemini APIを使って対話形式ポッドキャスト原稿を生成する"""
    articles_text_parts = []
    for idx, a in enumerate(articles):
        articles_text_parts.append(
            f"### #{idx + 1:02d} {a.title_ja}\n"
            f"- 出典: {a.source_name} ({a.country})\n"
            f"- 日付: {a.publish_date}\n"
            f"- 概要: {a.summary_ja}\n"
        )
    articles_text = "\n".join(articles_text_parts)

    prompt = build_dialogue_prompt(topic_name) + articles_text

    response = client.models.generate_content(
        model=model,
        contents=prompt,
    )

    return response.text.strip()


def parse_dialogue_script(text: str) -> list[dict]:
    """F[感情]:/M[感情]:形式の対話原稿をパースし、speaker, text, emotion のリストを返す

    対応フォーマット:
      F[H]: テキスト  → speaker="F", emotion="H", text="テキスト"
      F: テキスト      → speaker="F", emotion="N", text="テキスト"（感情タグなし→N）
      M[E]: テキスト   → speaker="M", emotion="E", text="テキスト"
    """
    # 感情タグ付きパターン: F[H]: or M[E]: （全角コロンも対応）
    emotion_pattern = re.compile(r'^([FM])\[([A-Z])\]\s*[：:](.*)$')
    # 感情タグなしパターン: F: or M:
    simple_pattern = re.compile(r'^([FM])\s*[：:](.*)$')

    lines = []
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue

        m = emotion_pattern.match(line)
        if m:
            speaker = m.group(1)
            emotion = m.group(2)
            txt = m.group(3).strip()
            lines.append({"speaker": speaker, "emotion": emotion, "text": txt})
            continue

        m2 = simple_pattern.match(line)
        if m2:
            speaker = m2.group(1)
            txt = m2.group(2).strip()
            lines.append({"speaker": speaker, "emotion": "N", "text": txt})
            continue

        # 前の話者の続き
        if lines:
            lines[-1]["text"] += line

    return lines


def split_long_text(text: str, max_chars: int = MAX_CHARS) -> list[str]:
    """長いテキストを max_chars 以内に分割"""
    if len(text) <= max_chars:
        return [text]

    segments = []
    # 句点で分割を試みる
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
            # 読点で分割
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


def _get_emotion_params(speaker: str, emotion: str) -> dict:
    """感情タグからVoicePeakのパラメータを取得する

    Returns:
        {"emotion": "happy=70,fun=50", "speed": 105, "pitch": 30}
    """
    defaults = {"emotion": "", "speed": SPEED, "pitch": 0}

    speaker_presets = EMOTION_PRESETS.get(speaker, {})
    if not speaker_presets:
        return defaults

    preset = speaker_presets.get(emotion, speaker_presets.get("N", {}))
    return {
        "emotion": preset.get("emotion", ""),
        "speed": preset.get("speed", SPEED),
        "pitch": preset.get("pitch", 0),
    }


def generate_wav_with_retry(
    text: str,
    narrator: str,
    output_path: Path,
    max_retries: int = MAX_RETRIES,
    speaker: str = "F",
    emotion: str = "N",
) -> bool:
    """VoicePeakで1セグメントをWAV化（リトライ付き・感情パラメータ対応）"""
    params = _get_emotion_params(speaker, emotion)

    cmd = [
        VOICEPEAK_EXE,
        "--say", text,
        "--narrator", narrator,
        "--speed", str(params["speed"]),
        "--pitch", str(params["pitch"]),
        "--out", str(output_path),
    ]

    # 感情パラメータがあれば追加
    if params["emotion"]:
        cmd.extend(["--emotion", params["emotion"]])

    for attempt in range(max_retries):
        try:
            result = subprocess.run(
                cmd, capture_output=True,
                timeout=120,
            )
            if result.returncode == 0 and output_path.exists():
                return True
        except subprocess.TimeoutExpired:
            pass
        except Exception:
            pass

        # リトライ前に待機
        if attempt < max_retries - 1:
            time.sleep(RETRY_DELAY)

    return False


SENTENCE_PAUSE = _vp_config.get("sentence_pause", 0.3)  # 同一話者の文間の無音（秒）


def _split_sentences(text: str) -> list[str]:
    """テキストを文単位で分割する。

    句読点（。！？）がある場合はそこで分割。
    句読点がない場合は、文末表現（です/ます/した 等）+
    スペースのパターンで分割する。各文は140字制限も適用。
    """
    # まず句読点があるかチェック
    if re.search(r'[。！？]', text):
        raw_sentences = re.split(r'(?<=[。！？])', text)
    else:
        # 句読点なし → 文末パターン + スペース で区切る
        # パターン: 「です」「ます」「した」「ません」等 + スペース
        # re.split でキャプチャグループを使い、区切り部分も残す
        parts = re.split(
            r'((?:です|ます|した|ません|でしょう|ました|ています|ております)'
            r'(?:ね|よ|が|けど)?)\s+',
            text,
        )
        # parts は [text, match, text, match, text, ...] の形になるので
        # text + match を結合して文にする
        raw_sentences = []
        i = 0
        while i < len(parts):
            if i + 1 < len(parts) and parts[i + 1]:
                # テキスト部分 + 文末パターン
                raw_sentences.append(parts[i] + parts[i + 1])
                i += 2
            else:
                raw_sentences.append(parts[i])
                i += 1

    result = []
    for s in raw_sentences:
        s = s.strip()
        if not s:
            continue
        # さらに140字制限で分割
        result.extend(split_long_text(s))
    return result if result else [text]


def _generate_silence_wav(work_dir: Path, duration: float = SENTENCE_PAUSE) -> Path:
    """ffmpegで無音WAVファイルを生成する。
    日本語パスを回避するためASCII一時ディレクトリで生成してからコピーする。"""
    silence_path = work_dir / "_silence.wav"
    if silence_path.exists():
        return silence_path

    # ASCII一時ディレクトリで生成
    tmp_dir = Path(tempfile.mkdtemp(prefix="podcast_silence_"))
    tmp_wav = tmp_dir / "_silence.wav"
    try:
        cmd = [
            "ffmpeg", "-y",
            "-f", "lavfi",
            "-i", f"anullsrc=r=44100:cl=mono",
            "-t", f"{duration:.2f}",
            "-acodec", "pcm_s16le",
            str(tmp_wav),
        ]
        subprocess.run(cmd, capture_output=True, timeout=10)
        if tmp_wav.exists():
            shutil.copy2(tmp_wav, silence_path)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return silence_path


def generate_audio_segments(
    dialogue: list[dict],
    work_dir: Path,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> list[Path]:
    """対話原稿の各台詞をVoicePeakでWAV化（男女振り分け）

    同一話者のセリフ内で文（。！？）が切り替わる箇所に
    0.3秒の無音を挿入し、自然な間を作る。
    """
    # --- 1. 全台詞を文単位でセグメント化 ---
    # 各エントリ: (speaker, emotion, text, is_sentence_boundary)
    # is_sentence_boundary=True → この前に同一話者の文間の無音を挿入
    all_segments = []
    for line in dialogue:
        emotion = line.get("emotion", "N")
        sentences = _split_sentences(line["text"])
        for si, sentence in enumerate(sentences):
            is_boundary = (si > 0)  # 同じセリフの2文目以降
            all_segments.append((line["speaker"], emotion, sentence, is_boundary))

    # --- 2. 無音WAVを準備 ---
    silence_wav = _generate_silence_wav(work_dir)
    if not silence_wav.exists():
        if progress_callback:
            progress_callback("WARNING: 無音WAV生成に失敗。文間の間なしで続行します")

    # --- 3. 各セグメントをWAV化（感情パラメータ付き） ---
    wav_files = []
    total = len(all_segments)
    seg_idx = 0  # WAVファイル番号

    for i, (speaker, emotion, text, is_boundary) in enumerate(all_segments):
        # 同一話者の文間に無音を挿入
        if is_boundary and silence_wav.exists():
            wav_files.append(silence_wav)

        narrator = NARRATOR_F if speaker == "F" else NARRATOR_M
        wav_path = work_dir / f"seg_{seg_idx:03d}.wav"
        speaker_label = "女性" if speaker == "F" else "男性"
        emotion_label = {"N": "", "H": "🙂", "E": "🤩", "S": "🤔", "Q": "❓", "T": "😏"}.get(emotion, "")

        if progress_callback:
            progress_callback(
                f"音声生成: {i + 1}/{total} ({speaker_label}{emotion_label}, {len(text)}字)")

        if generate_wav_with_retry(text, narrator, wav_path,
                                   speaker=speaker, emotion=emotion):
            wav_files.append(wav_path)
        else:
            if progress_callback:
                progress_callback(
                    f"  WARNING: セグメント{i + 1}({speaker_label})が"
                    f"{MAX_RETRIES}回リトライ後も失敗")

        seg_idx += 1

    return wav_files


def combine_to_mp3(
    wav_files: list[Path],
    output_path: Path,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> bool:
    """WAVファイルをffmpegで結合しMP3に変換。
    日本語パスを回避するためASCII一時ディレクトリで作業する。"""
    if not wav_files:
        return False

    if progress_callback:
        progress_callback(f"{len(wav_files)}セグメントをMP3に結合中...")

    # ASCII一時ディレクトリにWAVをコピーして作業
    tmp_dir = Path(tempfile.mkdtemp(prefix="podcast_combine_"))
    try:
        # WAVファイルを連番でコピー（重複除去: silenceは1回だけコピー）
        copied = {}  # 元パス → tmp内のファイル名
        list_entries = []
        for i, wav in enumerate(wav_files):
            src = str(wav)
            if src in copied:
                # 既にコピー済み（silence等）
                list_entries.append(copied[src])
            else:
                tmp_name = f"w_{i:04d}.wav"
                shutil.copy2(wav, tmp_dir / tmp_name)
                copied[src] = tmp_name
                list_entries.append(tmp_name)

        # filelist を相対パスで作成
        list_file = tmp_dir / "filelist.txt"
        with open(list_file, "w", encoding="utf-8") as f:
            for name in list_entries:
                f.write(f"file '{name}'\n")

        cmd = [
            "ffmpeg", "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", "filelist.txt",
            "-codec:a", "libmp3lame",
            "-b:a", "192k",
            "-ar", "44100",
            "output.mp3",
        ]
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=300,
            cwd=str(tmp_dir),
        )
        tmp_out = tmp_dir / "output.mp3"
        if result.returncode == 0 and tmp_out.exists():
            shutil.copy2(tmp_out, output_path)
            size_mb = output_path.stat().st_size / (1024 * 1024)
            if progress_callback:
                progress_callback(f"MP3生成完了: {size_mb:.1f} MB")
            return True
        else:
            if progress_callback:
                progress_callback(f"ffmpegエラー: {result.stderr[:200]}")
            return False
    except subprocess.TimeoutExpired:
        if progress_callback:
            progress_callback("ffmpegがタイムアウトしました")
        return False
    except Exception as e:
        if progress_callback:
            progress_callback(f"MP3結合エラー: {str(e)[:200]}")
        return False
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def generate_podcast(
    client,
    articles: list,
    topic_name: str,
    output_folder: Path,
    date_str: str,
    model: str = "gemini-2.5-flash",
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Optional[Path]:
    """対話形式ポッドキャスト生成のメインフロー

    1. Geminiで男女対話原稿を生成
    2. F:/M:をパースし、140字以内に分割
    3. VoicePeakで男女ナレーターに振り分けて音声化（リトライ付き）
    4. ffmpegでMP3結合

    Returns:
        生成されたMP3ファイルのPathまたはNone
    """
    if not is_voicepeak_available():
        if progress_callback:
            progress_callback("VoicePeakが見つかりません。ポッドキャスト生成をスキップします。")
        return None

    if not is_ffmpeg_available():
        if progress_callback:
            progress_callback("ffmpegが見つかりません。ポッドキャスト生成をスキップします。")
        return None

    # 出力ファイルパス: 調査項目名+日付+ポッドキャスト.mp3
    output_mp3 = output_folder / f"{topic_name}{date_str}ポッドキャスト.mp3"

    # 一時作業ディレクトリ
    work_dir = output_folder / "_podcast_work"
    work_dir.mkdir(exist_ok=True)

    try:
        # Step 1: 対話形式ポッドキャスト原稿生成
        if progress_callback:
            progress_callback("対話形式ポッドキャスト原稿を生成中...")

        script = generate_dialogue_script(client, articles, topic_name, model)

        # 原稿をファイルに保存
        script_path = output_folder / f"{topic_name}{date_str}ポッドキャスト原稿.txt"
        script_path.write_text(script, encoding="utf-8")

        if progress_callback:
            progress_callback(f"原稿生成完了: {len(script)}文字")

        # Step 2: 対話原稿をパース
        dialogue = parse_dialogue_script(script)
        if not dialogue:
            if progress_callback:
                progress_callback("対話原稿のパースに失敗しました")
            return None

        f_count = sum(1 for d in dialogue if d["speaker"] == "F")
        m_count = sum(1 for d in dialogue if d["speaker"] == "M")
        if progress_callback:
            progress_callback(f"台詞数: {len(dialogue)} (女性{f_count}、男性{m_count})")

        # Step 3: VoicePeakで男女音声を生成（リトライ付き）
        wav_files = generate_audio_segments(dialogue, work_dir, progress_callback)

        if not wav_files:
            if progress_callback:
                progress_callback("音声ファイルが生成されませんでした")
            return None

        total_segs = sum(
            len(split_long_text(d["text"])) for d in dialogue
        )
        if progress_callback:
            progress_callback(f"音声生成完了: {len(wav_files)}/{total_segs}セグメント")

        # Step 4: MP3に結合
        success = combine_to_mp3(wav_files, output_mp3, progress_callback)

        if success:
            if progress_callback:
                progress_callback(f"ポッドキャスト完成: {output_mp3.name}")
            return output_mp3
        else:
            if progress_callback:
                progress_callback("MP3結合に失敗しました")
            return None

    finally:
        # 一時作業ディレクトリを削除
        try:
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass
