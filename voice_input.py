"""
個人用音声入力ツール (MVP v2)
  - faster-whisper + Qwen 2.5 7B (Ollama) でローカル完結
  - Right Alt シングル: 録音開始/停止トグル
  - Right Alt ダブルタップ (400ms以内): 画面下常駐オーバーレイの表示/非表示
  - 常駐オーバーレイ: 録音状態 + リアルタイム音声波形
  - ESC で終了
  - すべての設定は config.json
"""
from __future__ import annotations

import datetime as dt
import json
import math
import re
import os
import queue
import sys
import threading
import time
import urllib.request
from collections import deque
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# pythonw.exe 経由(コンソールなし)で起動された場合はログファイルへリダイレクト
if sys.executable.lower().endswith("pythonw.exe"):
    _log_path = Path(__file__).parent / "voice_input.log"
    _log_file = open(_log_path, "a", encoding="utf-8", buffering=1)
    sys.stdout = _log_file
    sys.stderr = _log_file
    print(f"\n===== {time.strftime('%Y-%m-%d %H:%M:%S')} 起動 =====")

import socket
import ctypes
import ctypes.wintypes as wt
import numpy as np
import sounddevice as sd
import pyperclip
import tkinter as tk
from pynput import keyboard as pkb
from pynput.keyboard import Controller as KbController, Key

# Pillow + Cairo (overlay 描画用)
from PIL import Image, ImageDraw, ImageFilter, ImageFont
import cairo

try:
    import pystray
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

# ===== HighDPI 対応 =====
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _detect_scale() -> float:
    try:
        return ctypes.windll.user32.GetDpiForSystem() / 96.0
    except Exception:
        return 1.0


SCALE = _detect_scale()


def s(v: float) -> int:
    return int(round(v * SCALE))


# ===== カラー定数 =====
PURPLE_DEEP = (88, 30, 180)
PURPLE_MAIN = (140, 80, 220)
PURPLE_LIGHT = (180, 130, 250)
PURPLE_PINK = (210, 130, 240)

RED_DEEP = (180, 20, 40)
RED_MAIN = (235, 55, 70)
RED_LIGHT = (255, 110, 120)
RED_PINK = (255, 160, 160)


# ===== Win32 構造体 =====
class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wt.DWORD),
        ("biWidth", wt.LONG),
        ("biHeight", wt.LONG),
        ("biPlanes", wt.WORD),
        ("biBitCount", wt.WORD),
        ("biCompression", wt.DWORD),
        ("biSizeImage", wt.DWORD),
        ("biXPelsPerMeter", wt.LONG),
        ("biYPelsPerMeter", wt.LONG),
        ("biClrUsed", wt.DWORD),
        ("biClrImportant", wt.DWORD),
    ]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", _BITMAPINFOHEADER),
        ("bmiColors", wt.DWORD * 3),
    ]


class _BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_ubyte),
        ("BlendFlags", ctypes.c_ubyte),
        ("SourceConstantAlpha", ctypes.c_ubyte),
        ("AlphaFormat", ctypes.c_ubyte),
    ]


_user32 = ctypes.WinDLL("user32", use_last_error=True)
_gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
_user32.UpdateLayeredWindow.argtypes = [
    wt.HWND, wt.HDC, ctypes.POINTER(wt.POINT),
    ctypes.POINTER(wt.SIZE), wt.HDC, ctypes.POINTER(wt.POINT),
    wt.COLORREF, ctypes.POINTER(_BLENDFUNCTION), wt.DWORD,
]
_user32.UpdateLayeredWindow.restype = wt.BOOL
_gdi32.CreateDIBSection.argtypes = [
    wt.HDC, ctypes.POINTER(_BITMAPINFO), wt.UINT,
    ctypes.POINTER(ctypes.c_void_p), wt.HANDLE, wt.DWORD,
]
_gdi32.CreateDIBSection.restype = wt.HBITMAP
_gdi32.CreateCompatibleDC.argtypes = [wt.HDC]
_gdi32.CreateCompatibleDC.restype = wt.HDC
_gdi32.DeleteDC.argtypes = [wt.HDC]
_gdi32.DeleteDC.restype = wt.BOOL
_gdi32.SelectObject.argtypes = [wt.HDC, wt.HGDIOBJ]
_gdi32.SelectObject.restype = wt.HGDIOBJ
_gdi32.DeleteObject.argtypes = [wt.HGDIOBJ]
_gdi32.DeleteObject.restype = wt.BOOL
_user32.GetDC.argtypes = [wt.HWND]
_user32.GetDC.restype = wt.HDC
_user32.ReleaseDC.argtypes = [wt.HWND, wt.HDC]
_user32.ReleaseDC.restype = ctypes.c_int
_user32.GetCursorPos.argtypes = [ctypes.POINTER(wt.POINT)]
_user32.GetCursorPos.restype = wt.BOOL
_user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
_user32.GetAsyncKeyState.restype = ctypes.c_short


def _get_cursor_pos() -> tuple[int, int]:
    pt = wt.POINT()
    _user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def _is_ctrl_down() -> bool:
    """Windows API で現時点の Ctrl 状態を直接取得 (stale state 回避)"""
    VK_CONTROL = 0x11
    return bool(_user32.GetAsyncKeyState(VK_CONTROL) & 0x8000)


def _cairo_surface_to_pil(surface: cairo.ImageSurface) -> Image.Image:
    w = surface.get_width()
    h = surface.get_height()
    data = bytes(surface.get_data())
    arr = np.frombuffer(data, dtype=np.uint8).reshape(h, w, 4)
    rgba = arr[:, :, [2, 1, 0, 3]].copy()
    alpha = rgba[:, :, 3:4].astype(np.float32)
    with np.errstate(divide="ignore", invalid="ignore"):
        rgba[:, :, :3] = np.where(
            alpha > 0,
            np.clip(rgba[:, :, :3].astype(np.float32) * 255.0 / alpha, 0, 255),
            0,
        ).astype(np.uint8)
    return Image.fromarray(rgba, "RGBA")


def _cairo_rounded_rect(ctx, x, y, w, h, r):
    ctx.new_sub_path()
    ctx.arc(x + w - r, y + r, r, -math.pi / 2, 0)
    ctx.arc(x + w - r, y + h - r, r, 0, math.pi / 2)
    ctx.arc(x + r, y + h - r, r, math.pi / 2, math.pi)
    ctx.arc(x + r, y + r, r, math.pi, 3 * math.pi / 2)
    ctx.close_path()


def enable_window_blur(hwnd: int) -> bool:
    """Windows 10/11 のアクリル/ぼかし効果を window に適用 (best effort)"""
    import ctypes
    from ctypes import wintypes

    class ACCENTPOLICY(ctypes.Structure):
        _fields_ = [
            ("AccentState", ctypes.c_uint),
            ("AccentFlags", ctypes.c_uint),
            ("GradientColor", ctypes.c_uint),
            ("AnimationId", ctypes.c_uint),
        ]

    class WINCOMPATTRDATA(ctypes.Structure):
        _fields_ = [
            ("Attribute", ctypes.c_int),
            ("Data", ctypes.POINTER(ctypes.c_int)),
            ("SizeOfData", ctypes.c_size_t),
        ]

    ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    WCA_ACCENT_POLICY = 19

    try:
        user32 = ctypes.WinDLL("user32")
        SetWindowCompositionAttribute = user32.SetWindowCompositionAttribute
        accent = ACCENTPOLICY()
        accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2
        # ABGR: alpha=0x99 (60%), B=FF G=FF R=FF (white)
        accent.GradientColor = 0x99FFFFFF
        accent.AnimationId = 0

        data = WINCOMPATTRDATA()
        data.Attribute = WCA_ACCENT_POLICY
        data.SizeOfData = ctypes.sizeof(accent)
        data.Data = ctypes.cast(
            ctypes.pointer(accent), ctypes.POINTER(ctypes.c_int)
        )
        SetWindowCompositionAttribute(hwnd, ctypes.pointer(data))
        return True
    except Exception:
        return False

try:
    import win32gui
    import win32con
    import win32process
    import psutil
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# faster_whisper は import が重い (ctranslate2/tokenizers ロードで数秒)。
# オーバーレイを最速で表示するため Transcriber.__init__ 内で遅延 import する。
# from faster_whisper import WhisperModel  # 遅延化


CONFIG_PATH = Path(__file__).parent / "config.json"

# 単一インスタンスロック (グローバルで保持しないと GC で閉じてしまう)
_INSTANCE_LOCK_SOCKET: socket.socket | None = None


USER_DICT_PATH = Path(__file__).parent / "user_dictionary.json"
OVERLAY_STATE_PATH = Path(__file__).parent / "overlay_state.json"

# 辞書フィルタの上限（Whisper は initial_prompt が 224 トークン固定上限なので少なめ）
WHISPER_DICT_TOP_N = 35
LLM_DICT_TOP_N = 60

CATEGORY_DEFAULT_PRIORITY = {"person": 10, "project": 8, "tool": 5, "concept": 3}
CATEGORY_BONUS = {"person": 3, "project": 3, "tool": 1}
CATEGORY_LABEL = {
    "person": "人名",
    "project": "プロジェクト",
    "tool": "ツール",
    "concept": "概念",
    "other": "その他",
}


def load_user_dictionary() -> list:
    """user_dictionary.json を読み込み、term リストを返す。欠けフィールドはデフォルト埋め。"""
    if not USER_DICT_PATH.exists():
        return []
    try:
        with open(USER_DICT_PATH, encoding="utf-8") as f:
            data = json.load(f)
        terms = data.get("terms", [])
        # 後方互換: 欠けているフィールドをデフォルトで埋める
        for t in terms:
            cat = t.get("category", "other")
            t.setdefault("priority", CATEGORY_DEFAULT_PRIORITY.get(cat, 2))
            t.setdefault("hit_count", 0)
            t.setdefault("last_used", None)
        return terms
    except Exception as e:
        print(f"[Dictionary] ロード失敗: {e}", flush=True)
        return []


def _recency_boost(last_used) -> float:
    """last_used (ISO 日付文字列) に基づく最近利用ブースト。"""
    if not last_used:
        return 0.0
    try:
        last = dt.date.fromisoformat(last_used)
    except Exception:
        return 0.0
    days = (dt.date.today() - last).days
    if days < 0:
        return 0.0
    if days <= 7:
        return 5.0
    if days <= 30:
        return 2.0
    return 0.0


def _score_term(entry: dict) -> float:
    cat = entry.get("category", "other")
    base = float(entry.get("priority", CATEGORY_DEFAULT_PRIORITY.get(cat, 2)))
    hit = float(entry.get("hit_count", 0)) * 0.5
    rec = _recency_boost(entry.get("last_used"))
    bonus = float(CATEGORY_BONUS.get(cat, 0))
    return base + hit + rec + bonus


def score_and_filter(terms: list, top_n: int) -> list:
    """スコア順にソートして上位 top_n 件を返す。"""
    scored = sorted(terms, key=_score_term, reverse=True)
    return scored[:top_n]


def build_grouped_llm_hint(terms: list) -> str:
    """category 別にグルーピングした LLM 用ヒントブロックを生成する。"""
    if not terms:
        return ""
    groups: dict[str, list[str]] = {}
    order: list[str] = []
    for t in terms:
        cat = t.get("category", "other")
        if cat not in groups:
            groups[cat] = []
            order.append(cat)
        groups[cat].append(t["term"])
    # 表示順: person → project → tool → concept → other、以降は出現順
    preferred = ["person", "project", "tool", "concept", "other"]
    ordered_cats = [c for c in preferred if c in groups] + [
        c for c in order if c not in preferred
    ]
    lines = ["\n\n【このユーザーがよく使う固有名詞】"]
    for cat in ordered_cats:
        label = CATEGORY_LABEL.get(cat, cat)
        lines.append(f"{label}: {'、'.join(groups[cat])}")
    lines.append("音が近い誤認識を発見したら、上記のいずれかに修正してください。")
    return "\n".join(lines)


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        config = json.load(f)

    # ユーザー辞書を Whisper / LLM プロンプトに注入
    terms = load_user_dictionary()
    if terms:
        # Whisper: top-35 (initial_prompt 224 トークン制限を考慮)
        whisper_terms = score_and_filter(terms, WHISPER_DICT_TOP_N)
        current_prompt = config.get("whisper", {}).get("initial_prompt", "")
        extra = ", ".join(t["term"] for t in whisper_terms)
        config["whisper"]["initial_prompt"] = f"{current_prompt}, {extra}".strip(", ")

        # LLM: top-60 を category 別グルーピングで注入
        llm_terms = score_and_filter(terms, LLM_DICT_TOP_N)
        hint_block = build_grouped_llm_hint(llm_terms)
        for key in ("default", "code", "casual"):
            if key in config.get("prompts", {}):
                config["prompts"][key] += hint_block

        print(
            f"[Dictionary] {len(terms)} 件ロード "
            f"(Whisper top-{len(whisper_terms)}, LLM top-{len(llm_terms)})",
            flush=True,
        )

    return config


# ========== 辞書 hit tracker ==========

class DictionaryTracker:
    """
    辞書 hit_count / last_used の更新をメモリ内で行い、
    明示的な flush() タイミングでだけアトミックに JSON 書き戻す。

    - 実行中は load_config() の結果とは独立して生の terms を保持
    - _stop_and_process() で最終整形後の formatted テキストを受け取り、
      含まれる term の hit_count を +1、last_used を今日にする
    - realtime loop からは呼ばない (重複カウント防止)
    """

    def __init__(self, path: Path):
        self.path = path
        self.lock = threading.Lock()
        self.terms: list = []
        self.dirty = False
        self._load()

    def _load(self):
        if not self.path.exists():
            return
        try:
            with open(self.path, encoding="utf-8") as f:
                data = json.load(f)
            self.terms = data.get("terms", [])
            for t in self.terms:
                cat = t.get("category", "other")
                t.setdefault("priority", CATEGORY_DEFAULT_PRIORITY.get(cat, 2))
                t.setdefault("hit_count", 0)
                t.setdefault("last_used", None)
        except Exception as e:
            print(f"[DictTracker] ロード失敗: {e}", flush=True)

    def track_hits(self, text: str):
        """text に含まれる term の hit_count/last_used を更新 (メモリ内のみ)。"""
        if not text or not self.terms:
            return
        today = dt.date.today().isoformat()
        with self.lock:
            for t in self.terms:
                term = t.get("term")
                if term and term in text:
                    t["hit_count"] = int(t.get("hit_count", 0)) + 1
                    t["last_used"] = today
                    self.dirty = True

    def flush(self):
        """dirty なら JSON にアトミック書き戻し。"""
        with self.lock:
            if not self.dirty:
                return
            try:
                with open(self.path, encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {"version": 2, "description": "", "terms": []}
            # terms を上書き (Claude が追加した新規 entry は _load 時点では含まれないので、
            # 最新ファイルを再読込して term 名でマージする)
            existing_by_term = {t.get("term"): t for t in data.get("terms", [])}
            for t in self.terms:
                name = t.get("term")
                if not name:
                    continue
                if name in existing_by_term:
                    # hit_count / last_used だけを上書き (他フィールドは最新ファイル優先)
                    existing_by_term[name]["hit_count"] = t.get("hit_count", 0)
                    existing_by_term[name]["last_used"] = t.get("last_used")
                else:
                    # 既にファイルから消えていた場合は無視
                    pass
            data["terms"] = list(existing_by_term.values())
            data["last_updated"] = dt.date.today().isoformat()
            tmp = self.path.with_suffix(".json.tmp")
            try:
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    f.write("\n")
                os.replace(tmp, self.path)
                self.dirty = False
                print("[DictTracker] 辞書 flush 完了", flush=True)
            except Exception as e:
                print(f"[DictTracker] flush 失敗: {e}", flush=True)


def acquire_single_instance(port: int = 51721) -> bool:
    """ローカルポートをバインドして単一インスタンスを保証する"""
    global _INSTANCE_LOCK_SOCKET
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.listen(1)
        _INSTANCE_LOCK_SOCKET = s
        return True
    except OSError:
        s.close()
        return False


# ========== 録音 (callback で波形データも push) ==========

class Recorder:
    def __init__(self, sample_rate: int, amp_queue: queue.Queue):
        self.sample_rate = sample_rate
        self.recording = False
        self.frames: list[np.ndarray] = []
        self._stream: sd.InputStream | None = None
        self.amp_queue = amp_queue
        # 録音中でなくてもモニタリング（波形用）可能にする
        self.monitoring = False

    def _callback(self, indata, frames, time_info, status):
        # 波形用: 録音中 or モニタリング中のどちらでも振幅を push
        if self.monitoring or self.recording:
            chunk = indata[:, 0]
            rms = float(np.sqrt(np.mean(chunk ** 2)))
            try:
                self.amp_queue.put_nowait(rms)
            except queue.Full:
                pass
        if self.recording:
            self.frames.append(indata.copy())

    def _ensure_stream(self):
        if self._stream is None:
            self._stream = sd.InputStream(
                samplerate=self.sample_rate,
                channels=1,
                dtype="float32",
                callback=self._callback,
                blocksize=int(self.sample_rate * 0.05),  # 50ms チャンク
            )
            self._stream.start()

    def start_monitoring(self):
        self.monitoring = True
        self._ensure_stream()

    def stop_monitoring(self):
        self.monitoring = False
        if not self.recording and self._stream:
            self._stream.stop()
            self._stream.close()
            self._stream = None

    def start_recording(self):
        self.frames = []
        self.recording = True
        self._ensure_stream()

    def stop_recording(self) -> np.ndarray:
        self.recording = False
        if self._stream and not self.monitoring:
            self._stream.stop()
            self._stream.close()
            self._stream = None
        if not self.frames:
            return np.zeros(0, dtype=np.float32)
        return np.concatenate(self.frames).flatten()

    def get_current_audio(self) -> np.ndarray:
        """録音継続中でも現在までの音声を返す (リアルタイム用)"""
        if not self.frames:
            return np.zeros(0, dtype=np.float32)
        frames_copy = list(self.frames)  # スナップショット
        if not frames_copy:
            return np.zeros(0, dtype=np.float32)
        return np.concatenate(frames_copy).flatten()


# ========== Whisper ==========

class Transcriber:
    def __init__(self, config: dict):
        w = config["whisper"]
        print(f"[Whisper] ロード中: {w['model']}...", flush=True)
        t0 = time.perf_counter()
        # 遅延 import (モジュールロード時間短縮のため)
        from faster_whisper import WhisperModel
        self.model = WhisperModel(
            w["model"],
            device=w["device"],
            compute_type=w["compute_type"],
        )
        self.config = w
        print(f"[Whisper] 完了 ({time.perf_counter()-t0:.1f}秒)", flush=True)

    # Whisper が無音/末尾でハルシネーションする既知フレーズ
    HALLUCINATION_PHRASES = [
        # --- 日本語: YouTube 動画エンディング系 ---
        "ご視聴ありがとうございました",
        "ご視聴ありがとうございます",
        "ご清聴ありがとうございました",
        "ご清聴ありがとうございます",
        "最後までご視聴いただきありがとうございました",
        "最後までご視聴ありがとうございました",
        "見てくれてありがとう",
        "チャンネル登録お願いします",
        "チャンネル登録よろしくお願いします",
        "高評価チャンネル登録よろしくお願いします",
        "次の動画でお会いしましょう",
        "おやすみなさい",
        "お疲れ様でした",
        "ありがとうございました",
        # --- 中国語: 字幕クレジット / エンディング ---
        "由 Amara.org 社群提供的字幕",
        "由Amara.org社群提供的字幕",
        "中文字幕志愿者",
        "谢谢观看",
        "谢谢观看 下集再见",
        "请订阅我的频道",
        "字幕由Amara.org社区提供",
    ]

    # 日本語以外の文字が混入していないか判定するための正規表現
    # CJK Unified Ideographs は日中共通なので許可、中国語専用の簡体字範囲を検出
    _RE_CHINESE_ONLY = re.compile(
        r"[\u2E80-\u2EFF"   # CJK Radicals Supplement (中国語寄り)
        r"\u3100-\u312F"    # Bopomofo (注音, 中国語専用)
        r"\u31A0-\u31BF"    # Bopomofo Extended
        r"]"
    )
    # 文全体が非日本語 (中国語/韓国語等) で構成されているかの簡易判定:
    # ひらがな・カタカナが1文字もなければ日本語テキストではない可能性が高い
    _RE_JAPANESE_KANA = re.compile(r"[\u3040-\u309F\u30A0-\u30FF]")

    # no_speech_prob がこの閾値以上のセグメントはハルシネーションの疑い
    NO_SPEECH_PROB_THRESHOLD = 0.5

    def transcribe(self, audio: np.ndarray) -> tuple[str, float]:
        t0 = time.perf_counter()
        segments, info = self.model.transcribe(
            audio,
            language=self.config["language"],
            beam_size=self.config["beam_size"],
            initial_prompt=self.config.get("initial_prompt"),
            vad_filter=self.config.get("vad_filter", True),
        )
        # セグメント単位で no_speech_prob フィルタ
        filtered_parts = []
        for s in segments:
            if s.no_speech_prob >= self.NO_SPEECH_PROB_THRESHOLD:
                print(
                    f"[Whisper] 無音セグメント除外 (prob={s.no_speech_prob:.2f}): {s.text.strip()[:30]}",
                    flush=True,
                )
                continue
            filtered_parts.append(s.text)
        text = "".join(filtered_parts).strip()
        text = self._strip_hallucinations(text)
        elapsed = (time.perf_counter() - t0) * 1000
        return text, elapsed

    def _strip_hallucinations(self, text: str) -> str:
        """末尾に付加されたハルシネーションフレーズを除去し、非日本語混入を検出する。"""
        if not text:
            return text

        # 1) フレーズリストによる除去 (末尾一致 & 完全一致)
        changed = True
        while changed:
            changed = False
            for phrase in self.HALLUCINATION_PHRASES:
                if text == phrase:
                    return ""
                if text.endswith(phrase):
                    text = text[: -len(phrase)].rstrip("。、,.!！ ")
                    changed = True

        # 2) 中国語専用文字の混入チェック
        if self._RE_CHINESE_ONLY.search(text):
            print(f"[Whisper] 中国語文字混入を検出、除外: {text[:40]}", flush=True)
            return ""

        # 3) かな文字が一切ない長文は非日本語ハルシネーションの疑い
        #    (漢字だけの短い単語は正当なケースがあるので 8文字以上に限定)
        if len(text) >= 8 and not self._RE_JAPANESE_KANA.search(text):
            print(f"[Whisper] 非日本語テキスト検出、除外: {text[:40]}", flush=True)
            return ""

        # 4) 同一フレーズの繰り返し検出 (3回以上)
        if len(text) >= 6:
            chunk = text[:len(text) // 3]
            if chunk and text == chunk * (len(text) // len(chunk)) and len(text) // len(chunk) >= 3:
                print(f"[Whisper] 繰り返しハルシネーション検出、除外: {chunk[:20]}x{len(text)//len(chunk)}", flush=True)
                return ""

        return text


# ========== Ollama ==========

class Formatter:
    # 日本語では使わない中国語助詞・簡体字。これが出たら中国語混入と判定して raw にフォールバック
    _CHINESE_MARKERS = frozenset("你吗呢请这让谢说给们个问时发没东车进简对过")

    def __init__(self, config: dict):
        self.llm_config = config["llm"]
        self.prompts = config["prompts"]

    @classmethod
    def _contains_chinese(cls, text: str) -> str | None:
        for c in text:
            if c in cls._CHINESE_MARKERS:
                return c
        return None

    def format(self, text: str, prompt_key: str = "default") -> tuple[str, float]:
        system = self.prompts.get(prompt_key, self.prompts["default"])
        payload = {
            "model": self.llm_config["model"],
            "system": system,
            "prompt": text,
            "stream": False,
            "keep_alive": self.llm_config.get("keep_alive", "30m"),
            "options": {
                "temperature": self.llm_config.get("temperature", 0.3),
                "num_predict": self.llm_config.get("num_predict", 512),
            },
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            self.llm_config["ollama_url"],
            data=data,
            headers={"Content-Type": "application/json"},
        )
        t0 = time.perf_counter()
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                body = json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            print(f"[Formatter] エラー: {e}", flush=True)
            return text, 0.0
        elapsed = (time.perf_counter() - t0) * 1000
        formatted = body.get("response", text).strip()

        marker = self._contains_chinese(formatted)
        if marker:
            print(f"[Formatter] 中国語混入検出 ('{marker}'), rawにフォールバック: {formatted}", flush=True)
            return text, elapsed

        return formatted, elapsed

    def warmup(self):
        try:
            self.format("こんにちは", "default")
        except Exception:
            pass


# ========== アクティブウィンドウ検出 ==========

def get_active_window_process() -> str:
    if not HAS_WIN32:
        return ""
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return psutil.Process(pid).name()
    except Exception:
        return ""


def pick_prompt_key(config: dict) -> str:
    app_name = get_active_window_process()
    routing = config.get("app_routing", {})
    for exe, key in routing.items():
        if exe.lower() == app_name.lower():
            return key
    return "default"


# ========== 出力 ==========

_kb_controller = KbController()


def output_text(text: str, config: dict):
    output_cfg = config.get("output", {})
    if output_cfg.get("copy_to_clipboard", True):
        pyperclip.copy(text)
    if output_cfg.get("auto_paste", True):
        time.sleep(0.05)
        # pynput Controller で Ctrl+V を送信
        _kb_controller.press(Key.ctrl)
        _kb_controller.press("v")
        _kb_controller.release("v")
        _kb_controller.release(Key.ctrl)


# ========== オーバーレイ (Cairo + PIL + UpdateLayeredWindow) ==========


class Overlay:
    """pill + glow + マイクボタン + 波形 を UpdateLayeredWindow で描画"""

    def __init__(
        self,
        root: tk.Tk,
        config: dict,
        amp_queue: queue.Queue,
        on_button_click=None,
        on_close_click=None,
    ):
        self.root = root
        self.full_config = config
        self.config = config.get("overlay", {})
        self.amp_queue = amp_queue
        self.on_button_click = on_button_click
        self.on_close_click = on_close_click

        # アニメーション用の pill 高さ (初期値: compact) — _compute_layout 前に初期化が必要
        self._compact_h_bootstrap = s(self.config.get("height", 72))
        self.current_pill_h = float(self._compact_h_bootstrap)

        # 永続状態を先に読む (is_minimized / 保存位置)
        state = self._load_state()
        self.is_minimized = bool(state.get("minimized", False))

        # レイアウト計算 (is_minimized に応じてすべての寸法を確定)
        self._compute_layout()

        # ポップアップボタン (−/+/×) のフェード/ホバー状態
        # 各ボタンは独立して near 判定 → フェードする
        self.close_btn_visible = 0.0
        self.min_btn_visible = 0.0
        self.expand_btn_visible = 0.0
        self.close_btn_hovered = False
        self.min_btn_hovered = False
        self.expand_btn_hovered = False

        # Mini ↔ Normal 遷移 (geometric morph)
        # current_pill_w を ease-out で target_pill_w に収束させ、
        # pill_x0/x1/mic_cx などを毎フレーム再計算する。
        self.current_pill_w = float(self.pill_w)
        self.target_pill_w = float(self.pill_w)
        self.transitioning = False
        self._transition_mic_cx_local = int(self.mic_cx)

        # ドラッグ状態
        self._drag_active = False     # マウスボタン押下中
        self._drag_moved = False      # 実際にドラッグ移動が発生したか
        self._press_target = None     # "mic" / "close" / "minimize" / "pill" / None
        self._drag_start_mx = 0
        self._drag_start_my = 0
        self._drag_start_wx = 0
        self._drag_start_wy = 0
        self._drag_threshold = s(5)
        self._saved_state = state

        # Waveform
        self.waveform_bars = 32
        self.amps = deque([0.0] * self.waveform_bars, maxlen=self.waveform_bars)
        self._amp_smoothed = 0.0  # EMA smoothing

        # State
        self.is_recording = False
        self.status_text = "waiting"
        # リアルタイム preview
        self.preview_text = ""       # raw Whisper
        self.preview_at = 0.0
        self.formatted_text = ""     # LLM 整形後
        self.formatted_at = 0.0
        self.visible = False
        self.phase = 0.0
        self.glow_intensity = 1.0

        # Intro
        self.intro_progress = 1.0
        self.intro_duration = 0.45
        self.intro_y_offset = s(40)
        self.intro_start = None

        # Create Toplevel
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)

        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        default_x = (sw - self.window_w) // 2
        default_y = sh - self.window_h - s(60)
        saved_x = self._saved_state.get("x")
        saved_y = self._saved_state.get("y")
        if isinstance(saved_x, int) and isinstance(saved_y, int):
            self.win_x = self._clamp_x(saved_x, sw)
            self.win_y = self._clamp_y(saved_y, sh)
        else:
            self.win_x = default_x
            self.win_y = default_y
        self.final_win_y = self.win_y
        self.win.geometry(
            f"{self.window_w}x{self.window_h}+{self.win_x}+{self.win_y}"
        )

        # Withdraw initially
        self.win.withdraw()

        # Set up as layered window (UpdateLayeredWindow 用)
        self.win.update_idletasks()
        self._make_layered()

        # Mouse handlers — press/motion/release でドラッグとクリックを判別
        self.win.bind("<ButtonPress-1>", self._on_press)
        self.win.bind("<B1-Motion>", self._on_motion)
        self.win.bind("<ButtonRelease-1>", self._on_release)

        # Font (日本語対応: PIL ImageFont で truetype ロード)
        self.font_pil = None
        for font_name in [
            "YuGothM.ttc",      # Yu Gothic Medium
            "YuGothR.ttc",      # Yu Gothic Regular
            "meiryo.ttc",       # Meiryo
            "msgothic.ttc",     # MS Gothic
            "segoeui.ttf",      # fallback
        ]:
            try:
                self.font_pil = ImageFont.truetype(font_name, s(13))
                print(f"[Overlay] font loaded: {font_name}", flush=True)
                break
            except Exception:
                continue
        if self.font_pil is None:
            self.font_pil = ImageFont.load_default()

        # Animation loop
        self._animate()

    # ---------- Win32 ----------

    def _hwnd(self) -> int:
        return int(self.win.wm_frame(), 16)

    def _make_layered(self):
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x00080000
        WS_EX_TOOLWINDOW = 0x00000080
        WS_EX_NOACTIVATE = 0x08000000
        try:
            hwnd = self._hwnd()
            style = _user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            _user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE,
                style | WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            )
        except Exception as e:
            print(f"[Overlay] layered set 失敗: {e}", flush=True)

    def _push_image(self, pil_rgba: Image.Image):
        hwnd = self._hwnd()
        arr = np.asarray(pil_rgba, dtype=np.uint8)
        alpha = arr[:, :, 3:4].astype(np.float32) / 255.0
        rgb = (arr[:, :, :3].astype(np.float32) * alpha).astype(np.uint8)
        a = arr[:, :, 3:4]
        b = rgb[:, :, 2:3]
        g = rgb[:, :, 1:2]
        r_ch = rgb[:, :, 0:1]
        bgra = np.concatenate([b, g, r_ch, a], axis=2)
        data = bgra.tobytes()

        bmi = _BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = self.window_w
        bmi.bmiHeader.biHeight = -self.window_h
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0

        ppvBits = ctypes.c_void_p()
        hBitmap = _gdi32.CreateDIBSection(
            0, ctypes.byref(bmi), 0, ctypes.byref(ppvBits), 0, 0,
        )
        if not hBitmap:
            return

        ctypes.memmove(ppvBits, data, len(data))

        hdcScreen = _user32.GetDC(0)
        hdcMem = _gdi32.CreateCompatibleDC(hdcScreen)
        hbmOld = _gdi32.SelectObject(hdcMem, hBitmap)

        size = wt.SIZE(self.window_w, self.window_h)
        ptSrc = wt.POINT(0, 0)
        ptDst = wt.POINT(self.win_x, self.win_y)

        blend = _BLENDFUNCTION()
        blend.BlendOp = 0
        blend.BlendFlags = 0
        blend.SourceConstantAlpha = 255
        blend.AlphaFormat = 1

        ULW_ALPHA = 0x00000002
        _user32.UpdateLayeredWindow(
            hwnd, hdcScreen,
            ctypes.byref(ptDst), ctypes.byref(size),
            hdcMem, ctypes.byref(ptSrc),
            0, ctypes.byref(blend), ULW_ALPHA,
        )

        _gdi32.SelectObject(hdcMem, hbmOld)
        _gdi32.DeleteObject(hBitmap)
        _gdi32.DeleteDC(hdcMem)
        _user32.ReleaseDC(0, hdcScreen)

    # ---------- パブリック API ----------

    def show(self):
        # Reset intro animation
        self.intro_start = time.time()
        self.intro_progress = 0.0
        self.win_y = self.final_win_y + self.intro_y_offset

        try:
            hwnd = self._hwnd()
            # SW_SHOWNOACTIVATE = 4
            _user32.ShowWindow(hwnd, 4)
        except Exception:
            self.win.deiconify()
        self.visible = True

    def hide(self):
        try:
            hwnd = self._hwnd()
            _user32.ShowWindow(hwnd, 0)  # SW_HIDE
        except Exception:
            self.win.withdraw()
        self.visible = False

    def toggle(self):
        if self.visible:
            self.hide()
        else:
            self.show()

    def set_recording(self, recording: bool, status: str = None):
        self.is_recording = recording
        if status is not None:
            self.status_text = status
        elif recording:
            self.status_text = "recording"
        else:
            self.status_text = "waiting"

    def set_status(self, text: str):
        self.status_text = text

    def set_preview_text(self, text: str):
        self.preview_text = text
        self.preview_at = time.time()

    def set_formatted_text(self, text: str):
        self.formatted_text = text
        self.formatted_at = time.time()

    # ---------- レイアウト計算 ----------

    def _compute_layout(self):
        """is_minimized に応じてすべてのレイアウト寸法を再計算。"""
        cfg = self.config
        pill_w_l = cfg.get("width", 480)
        compact_h_l = cfg.get("height", 72)
        expanded_extra_l = 80

        self._compact_h = s(compact_h_l)
        self._expanded_h = s(compact_h_l + expanded_extra_l)
        self._normal_pill_w = s(pill_w_l)  # mini↔normal 遷移の lerp 基準
        self.pill_radius = self._compact_h // 2  # radius は compact 基準で固定

        if self.is_minimized:
            # mini: mic 1 個分の丸ピル (幅 = compact_h で真円)
            self.pill_w = self._compact_h
            self.glow_pad_x = s(50)
            self.glow_pad_top = s(50)
            self.glow_pad_bottom = s(50)
            # ミニ中は高さアニメを無効化
            self.current_pill_h = float(self._compact_h)
        else:
            self.pill_w = s(pill_w_l)
            self.glow_pad_x = s(85)
            self.glow_pad_top = s(85)
            self.glow_pad_bottom = s(85)

        self.hover_area_extra = s(40)

        # Window サイズ
        self.window_w = self.pill_w + 2 * self.glow_pad_x
        if self.is_minimized:
            self.window_h = (
                self._compact_h + self.glow_pad_top + self.glow_pad_bottom
            )
        else:
            self.window_h = (
                self._expanded_h + self.glow_pad_top + self.glow_pad_bottom
            )

        # Pill 座標
        self.pill_x0 = self.glow_pad_x
        self.pill_x1 = self.pill_x0 + self.pill_w
        if self.is_minimized:
            self.pill_y1 = self.glow_pad_top + self._compact_h
            self.pill_y0 = self.glow_pad_top
        else:
            self.pill_y1 = self.glow_pad_top + self._expanded_h
            self.pill_y0 = self.pill_y1 - int(self.current_pill_h)

        self.original_cy = self.pill_y1 - self._compact_h // 2
        self.cy_mid = self.original_cy

        # マイクボタン
        self.mic_r = s(22)
        if self.is_minimized:
            self.mic_cx = self.pill_x0 + self.pill_w // 2
        else:
            self.mic_cx = self.pill_x1 - s(50)
        self.mic_cy = self.original_cy

        # ポップアップボタンサイズ (− / + / × 共通)
        self.close_btn_w = s(26)
        self.close_btn_h = s(26)

        # × (close): 従来どおり pill 下部中央にポップアップ
        pill_cx_mid = (self.pill_x0 + self.pill_x1) // 2
        self.close_btn_cx = pill_cx_mid
        self.close_btn_cy = self.pill_y1 + s(18)

        # − (minimize) / + (expand): pill 右上の外側にポップアップ
        # 位置は compact pill の上端基準 (expand アニメの影響を受けない)
        top_anchor = self.pill_y1 - self._compact_h
        top_y = top_anchor - s(18)
        right_margin = s(14)
        if self.is_minimized:
            # mini 中: + (expand) のみ右上に配置、− (minimize) は画面外
            self.expand_btn_cx = (
                self.pill_x1 - right_margin - self.close_btn_w // 2
            )
            self.expand_btn_cy = top_y
            self.min_btn_cx = -99999
            self.min_btn_cy = -99999
        else:
            # 通常時: − (minimize) のみ右上に配置、+ (expand) は画面外
            self.min_btn_cx = (
                self.pill_x1 - right_margin - self.close_btn_w // 2
            )
            self.min_btn_cy = top_y
            self.expand_btn_cx = -99999
            self.expand_btn_cy = -99999

    # ---------- 永続状態 ----------

    def _load_state(self) -> dict:
        try:
            with open(OVERLAY_STATE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"[Overlay] state load 失敗: {e}", flush=True)
        return {}

    def _save_state(self):
        try:
            data = {
                "x": int(self.win_x),
                "y": int(self.win_y),
                "minimized": bool(self.is_minimized),
            }
            with open(OVERLAY_STATE_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[Overlay] state save 失敗: {e}", flush=True)

    # ---------- 画面クランプ ----------

    def _clamp_x(self, x: int, screen_w: int) -> int:
        # 最低 60px は画面内に残す
        margin = s(60)
        min_x = margin - self.window_w
        max_x = screen_w - margin
        return max(min_x, min(max_x, x))

    def _clamp_y(self, y: int, screen_h: int) -> int:
        margin = s(60)
        min_y = 0
        max_y = screen_h - margin
        return max(min_y, min(max_y, y))

    # ---------- マウスハンドラ ----------

    def _hit_test(self, lx: int, ly: int) -> str | None:
        """window-local 座標からクリック対象を判定。"""
        half_w = self.close_btn_w // 2
        half_h = self.close_btn_h // 2
        # 各ボタンは独立に visible 判定 (off-screen 中は visible が 0 のままなのでスキップ)
        # expand (+) — mini 中のみ右上
        if self.expand_btn_visible > 0.5 and (
            self.expand_btn_cx - half_w <= lx <= self.expand_btn_cx + half_w
            and self.expand_btn_cy - half_h <= ly <= self.expand_btn_cy + half_h
        ):
            return "expand"
        # minimize (−) — 通常時のみ右上
        if self.min_btn_visible > 0.5 and (
            self.min_btn_cx - half_w <= lx <= self.min_btn_cx + half_w
            and self.min_btn_cy - half_h <= ly <= self.min_btn_cy + half_h
        ):
            return "minimize"
        # close (×) — 下部中央
        if self.close_btn_visible > 0.5 and (
            self.close_btn_cx - half_w <= lx <= self.close_btn_cx + half_w
            and self.close_btn_cy - half_h <= ly <= self.close_btn_cy + half_h
        ):
            return "close"

        # マイクボタン (円)
        dx = lx - self.mic_cx
        dy = ly - self.mic_cy
        if dx * dx + dy * dy <= self.mic_r * self.mic_r:
            return "mic"

        # Pill 本体 (ドラッグ可能領域)
        if (
            self.pill_x0 <= lx <= self.pill_x1
            and self.pill_y0 <= ly <= self.pill_y1
        ):
            return "pill"

        return None

    def _on_press(self, event):
        self._drag_active = True
        self._drag_moved = False
        self._drag_start_mx = event.x_root
        self._drag_start_my = event.y_root
        self._drag_start_wx = self.win_x
        self._drag_start_wy = self.win_y
        self._press_target = self._hit_test(event.x, event.y)

    def _on_motion(self, event):
        if not self._drag_active:
            return
        dx = event.x_root - self._drag_start_mx
        dy = event.y_root - self._drag_start_my
        if not self._drag_moved:
            if abs(dx) < self._drag_threshold and abs(dy) < self._drag_threshold:
                return
            # ドラッグ開始できるのは pill / mic をつかんだときだけ
            # (ボタン −/× は単なるクリック扱い)
            if self._press_target not in ("pill", "mic"):
                return
            self._drag_moved = True

        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        new_x = self._clamp_x(self._drag_start_wx + dx, sw)
        new_y = self._clamp_y(self._drag_start_wy + dy, sh)
        if new_x == self.win_x and new_y == self.win_y:
            return
        self.win_x = new_x
        self.win_y = new_y
        self.final_win_y = new_y  # intro アニメが走っていない前提で追従
        self.win.geometry(f"+{new_x}+{new_y}")

    def _on_release(self, event):
        was_drag = self._drag_active and self._drag_moved
        target = self._press_target
        self._drag_active = False
        self._drag_moved = False
        self._press_target = None

        if was_drag:
            self._save_state()
            return

        # クリック扱い
        if target == "mic":
            if self.on_button_click:
                self.on_button_click()
        elif target == "close":
            if self.on_close_click:
                self.on_close_click()
        elif target == "minimize":
            self._set_minimized(True)
        elif target == "expand":
            self._set_minimized(False)

    def _set_minimized(self, flag: bool):
        """ミニ ↔ 通常 の切替トリガ。geometric morph で繋ぐ。

        戦略:
        - normal → mini: window は normal サイズのまま、pill_w を compact_h へ ease。
          完了時に window を mini サイズへリサイズ (mic の screen 位置を anchor)。
        - mini → normal: 先に window を normal サイズへリサイズ (mic anchor)。
          pill_w は compact_h スタートのまま normal_pill_w へ ease。
        """
        if self.is_minimized == flag:
            return
        if self.transitioning:
            return  # 遷移中の割り込みは無視
        prev_mic_cx_screen = self.win_x + self.mic_cx
        prev_mic_cy_screen = self.win_y + self.mic_cy
        self.is_minimized = flag

        if flag:
            # normal → mini: window はそのまま。pill_w を縮めるアニメへ。
            self._transition_mic_cx_local = int(self.mic_cx)
            self.current_pill_w = float(self.pill_w)
            self.target_pill_w = float(self._compact_h)
        else:
            # mini → normal: 先に window を normal へリサイズ (mic anchor)
            self._compute_layout()
            new_x = prev_mic_cx_screen - self.mic_cx
            new_y = prev_mic_cy_screen - self.mic_cy
            sw = self.win.winfo_screenwidth()
            sh = self.win.winfo_screenheight()
            self.win_x = self._clamp_x(new_x, sw)
            self.win_y = self._clamp_y(new_y, sh)
            self.final_win_y = self.win_y
            self.win.geometry(
                f"{self.window_w}x{self.window_h}+{self.win_x}+{self.win_y}"
            )
            # pill_w は compact_h スタート → normal_pill_w ターゲット
            self._transition_mic_cx_local = int(self.mic_cx)
            self.current_pill_w = float(self._compact_h)
            self.target_pill_w = float(self.pill_w)

        self.transitioning = True
        # ポップアップは全部隠す (遷移中はホバー判定もスキップ)
        self.close_btn_visible = 0.0
        self.min_btn_visible = 0.0
        self.expand_btn_visible = 0.0
        self._save_state()

        # 遷移初期フレームの位置を即座に反映して再描画する。
        # 特に mini→normal では window リサイズ直後に古い mini bitmap が
        # upper-left に残ってしまうため、同期的に新しい状態を push する必要がある。
        self._apply_transition_positions()
        if self.visible:
            self._render()

    def _apply_transition_positions(self):
        """transitioning 中に current_pill_w から pill/mic/ボタン位置を再計算する。

        _animate の transition block と _set_minimized 初回フレーム強制レンダ
        の両方から呼ばれるヘルパ。
        """
        denom = float(self._normal_pill_w - self._compact_h)
        if denom > 0:
            t = (self.current_pill_w - self._compact_h) / denom
        else:
            t = 1.0
        t = max(0.0, min(1.0, t))

        mic_offset = (
            self._compact_h / 2.0
            + (s(50) - self._compact_h / 2.0) * t
        )
        self.pill_x1 = int(self._transition_mic_cx_local + mic_offset)
        self.pill_x0 = self.pill_x1 - int(self.current_pill_w)
        self.mic_cx = self._transition_mic_cx_local
        self.pill_w = int(self.current_pill_w)

        pill_cx_mid = (self.pill_x0 + self.pill_x1) // 2
        self.close_btn_cx = pill_cx_mid
        right_margin = s(14)
        btn_cx = self.pill_x1 - right_margin - self.close_btn_w // 2
        if self.is_minimized:
            self.expand_btn_cx = btn_cx
        else:
            self.min_btn_cx = btn_cx

    # ---------- アニメーションループ ----------

    def _animate(self):
        self.phase += 0.05

        # ===== Pill 高さアニメーション =====
        # 完全 mini 中のみスナップ。遷移中は window が normal サイズなので
        # 通常の ease ロジックを使い、target は compact (または録音中は expanded)。
        if self.is_minimized and not self.transitioning:
            # 完全 mini — compact_h に固定
            self.current_pill_h = float(self._compact_h)
            self.pill_y0 = self.glow_pad_top
        else:
            if self.is_minimized:
                # 遷移中 (normal → mini): pill_h も compact へ ease
                target_h = self._compact_h
            else:
                target_h = (
                    self._expanded_h if self.is_recording else self._compact_h
                )
            diff = target_h - self.current_pill_h
            if abs(diff) < 0.5:
                self.current_pill_h = float(target_h)
            else:
                self.current_pill_h += diff * 0.18  # ease-out smoothing
            self.pill_y0 = self.pill_y1 - int(self.current_pill_h)

            # − (minimize) ボタンは pill 上端に追従 (録音中に pill が伸びても被らない)
            if not self.is_minimized:
                self.min_btn_cy = self.pill_y0 - s(18)

        # ===== Mini ↔ Normal geometric morph =====
        # current_pill_w を ease-out で target へ収束させ、pill/mic/ボタン位置を再計算。
        if self.transitioning:
            diff_w = self.target_pill_w - self.current_pill_w
            if abs(diff_w) < 0.4:
                self.current_pill_w = self.target_pill_w
                morph_finished = True
            else:
                self.current_pill_w += diff_w * 0.22  # ease-out
                morph_finished = False

            self._apply_transition_positions()

            # アニメ完了処理
            if morph_finished:
                if self.is_minimized:
                    # mini 用に window をリサイズ (mic screen pos 保持)
                    prev_mic_cx_screen = self.win_x + self.mic_cx
                    prev_mic_cy_screen = self.win_y + self.mic_cy
                    self._compute_layout()
                    new_x = prev_mic_cx_screen - self.mic_cx
                    new_y = prev_mic_cy_screen - self.mic_cy
                    sw = self.win.winfo_screenwidth()
                    sh = self.win.winfo_screenheight()
                    self.win_x = self._clamp_x(new_x, sw)
                    self.win_y = self._clamp_y(new_y, sh)
                    self.final_win_y = self.win_y
                    self.win.geometry(
                        f"{self.window_w}x{self.window_h}"
                        f"+{self.win_x}+{self.win_y}"
                    )
                self.current_pill_w = float(self.pill_w)
                self.target_pill_w = self.current_pill_w
                self.transitioning = False

        # ===== マウスホバー判定 (−/+/× ポップアップボタン) =====
        # 各ボタンは独立して判定する: 近接 (rect + padding) でそのボタンのみフェードイン、
        # 正確な rect に入ったら hovered フラグを立てて色を変える。
        # transition 中はポップアップを全部隠してホバー判定もスキップする。
        if self.visible and not self.transitioning:
            mx, my = _get_cursor_pos()
            lx = mx - self.win_x
            ly = my - self.win_y

            half_w = self.close_btn_w // 2
            half_h = self.close_btn_h // 2
            prox = s(18)  # 近づいたら表示されるまでの余裕

            def _in_rect(bcx, bcy, pad=0):
                return (
                    bcx - half_w - pad <= lx <= bcx + half_w + pad
                    and bcy - half_h - pad <= ly <= bcy + half_h + pad
                )

            # 正確な rect (click / hover color)
            in_close_btn = _in_rect(self.close_btn_cx, self.close_btn_cy)
            in_min_btn = _in_rect(self.min_btn_cx, self.min_btn_cy)
            in_expand_btn = _in_rect(self.expand_btn_cx, self.expand_btn_cy)
            self.close_btn_hovered = in_close_btn
            self.min_btn_hovered = in_min_btn
            self.expand_btn_hovered = in_expand_btn

            # 近接 zone (fade-in trigger)
            near_close = _in_rect(self.close_btn_cx, self.close_btn_cy, prox)
            near_min = _in_rect(self.min_btn_cx, self.min_btn_cy, prox)
            near_expand = _in_rect(self.expand_btn_cx, self.expand_btn_cy, prox)

            def _ease(cur, tgt):
                d = tgt - cur
                if abs(d) < 0.02:
                    return tgt
                return cur + d * 0.25

            self.close_btn_visible = _ease(
                self.close_btn_visible, 1.0 if near_close else 0.0
            )
            self.min_btn_visible = _ease(
                self.min_btn_visible, 1.0 if near_min else 0.0
            )
            self.expand_btn_visible = _ease(
                self.expand_btn_visible, 1.0 if near_expand else 0.0
            )
        else:
            self.close_btn_visible = 0.0
            self.min_btn_visible = 0.0
            self.expand_btn_visible = 0.0
            self.close_btn_hovered = False
            self.min_btn_hovered = False
            self.expand_btn_hovered = False

        # 実際の音声振幅を queue から読む
        latest_amp = 0.0
        try:
            while True:
                latest_amp = max(latest_amp, self.amp_queue.get_nowait())
        except queue.Empty:
            pass

        # EMA スムージング
        target = latest_amp if self.is_recording else 0.0
        self._amp_smoothed = self._amp_smoothed * 0.55 + target * 0.45
        self.amps.append(self._amp_smoothed)

        # Intro
        if self.intro_progress < 1.0:
            if self.intro_start is None:
                self.intro_start = time.time()
            elapsed = time.time() - self.intro_start
            t = min(1.0, elapsed / self.intro_duration)
            eased = 1.0 - (1.0 - t) ** 5
            self.intro_progress = eased
            self.win_y = int(
                self.final_win_y + (1.0 - eased) * self.intro_y_offset
            )

        # 可視の時だけ描画
        if self.visible:
            self._render()

        self.root.after(16, self._animate)

    # ---------- レンダリング ----------

    def _render(self):
        img = Image.new("RGBA", (self.window_w, self.window_h), (0, 0, 0, 0))

        glow = self._render_glow()
        img = Image.alpha_composite(img, glow)

        cairo_img = self._draw_pill_and_contents_cairo()
        img = Image.alpha_composite(img, cairo_img)

        # PIL で日本語対応テキスト描画 (上段: status + preview)
        self._draw_text_pil(img)

        # Close button (pill 下のホバー popup)
        self._draw_close_button_pil(img)

        if self.intro_progress < 1.0:
            arr = np.asarray(img, dtype=np.uint8).copy()
            arr[:, :, 3] = (
                arr[:, :, 3].astype(np.float32) * self.intro_progress
            ).astype(np.uint8)
            img = Image.fromarray(arr, "RGBA")

        self._push_image(img)

    def _draw_text_pil(self, img: Image.Image):
        """PIL で status text と preview text を描画 (日本語対応)"""
        # ミニ中はテキスト一切描画しない (マイクだけ表示)
        if self.is_minimized:
            return
        d = ImageDraw.Draw(img)
        font = self.font_pil

        # ===== Status text (既存の dot の右) =====
        if self.is_recording:
            status_color = (180, 40, 60, 255)
        else:
            status_color = (110, 115, 130, 255)
        status_x = self.pill_x0 + s(24 + 12)
        # ベースライン調整: text ascent を計算して cy 中央
        bbox_t = d.textbbox((0, 0), self.status_text, font=font)
        th = bbox_t[3] - bbox_t[1]
        status_y = self.original_cy - th // 2 - bbox_t[1]
        d.text(
            (status_x, status_y),
            self.status_text,
            font=font,
            fill=status_color,
        )

        # ===== Preview text (pill の上部空間、最大3行 wrap) =====
        # pill が十分に拡張されていなければ preview 非表示
        if self.current_pill_h < self._compact_h + s(20):
            return

        # 表示するテキストを決定 (formatted が新しければそれ、無ければ preview)
        now = time.time()
        if self.formatted_at > self.preview_at and self.formatted_text:
            display_text = self.formatted_text
            glow_age = now - self.formatted_at
            # 短め 0.5 秒で急峻に減衰 (ease-out quart)
            t = max(0.0, min(1.0, glow_age / 0.5))
            glow_val = (1.0 - t) ** 4
        else:
            display_text = self.preview_text
            glow_val = 0.0

        if not display_text:
            return

        # Preview 表示エリア (pill 上部の追加領域)
        pv_x0 = self.pill_x0 + s(24)
        pv_x1 = self.pill_x1 - s(24)
        area_w = pv_x1 - pv_x0

        # preview 用の上部領域 (pill_y0 + 余白 ～ original_cy - compact_h/2 - 余白)
        pv_top = self.pill_y0 + s(10)
        pv_bottom = self.original_cy - self._compact_h // 2 - s(4)
        pv_area_h = pv_bottom - pv_top
        if pv_area_h < s(10):
            return

        # テキストを文字単位で折り返し (日本語対応)
        line_height = s(19)
        max_lines = 3

        # bottom-up で wrap: 末尾から逆算して、表示できる最後の行群を決定
        # 単純な forward wrap を使い、最後の max_lines 行を取る
        def wrap_forward(text: str) -> list:
            lines = []
            current = ""
            for ch in text:
                test = current + ch
                bbox = d.textbbox((0, 0), test, font=font)
                w = bbox[2] - bbox[0]
                if w > area_w and current:
                    lines.append(current)
                    current = ch
                else:
                    current = test
            if current:
                lines.append(current)
            return lines

        all_lines = wrap_forward(display_text)
        display_lines = all_lines[-max_lines:]
        n_lines = len(display_lines)

        # 下から順に配置 (最新の行が一番下)
        total_h = n_lines * line_height
        start_y = pv_bottom - total_h

        # 各行の描画位置を事前計算
        line_positions = []
        for i, line in enumerate(display_lines):
            ly = start_y + i * line_height
            line_positions.append((pv_x0, ly, line))

        # ===== テキスト自体に紫グロー効果 =====
        # 別 layer に紫テキストを描画 → ガウシアンブラー → halo として合成
        if glow_val > 0.02:
            glow_layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
            glow_draw = ImageDraw.Draw(glow_layer)
            halo_alpha = int(255 * glow_val)
            for (lx, ly, line) in line_positions:
                glow_draw.text(
                    (lx, ly),
                    line,
                    font=font,
                    fill=(180, 120, 255, halo_alpha),
                )
            blur_r = s(3) + int(s(3) * glow_val)
            glow_layer = glow_layer.filter(
                ImageFilter.GaussianBlur(radius=blur_r)
            )
            img.alpha_composite(glow_layer)
            d = ImageDraw.Draw(img)

        # ===== シャープテキスト (上に重ねる) =====
        # 色: glow=1 で鮮やか紫、glow=0 でダークグレー
        base = (70, 75, 90)
        purple = (165, 100, 245)
        tr = int(base[0] + (purple[0] - base[0]) * glow_val)
        tg = int(base[1] + (purple[1] - base[1]) * glow_val)
        tb = int(base[2] + (purple[2] - base[2]) * glow_val)
        for (lx, ly, line) in line_positions:
            d.text((lx, ly), line, font=font, fill=(tr, tg, tb, 250))

    def _draw_close_button_pil(self, img: Image.Image):
        """pill 周りのホバー popup 式 ボタン群 (−/+/×)

        PIL の ImageDraw は ellipse/line に AA が効かないため、
        3x スーパーサンプリングして LANCZOS で縮小してから合成する。
        各ボタンは独立した *_visible を alpha として持つ。
        """
        btn_w = self.close_btn_w
        btn_h = self.close_btn_h
        SS = 3  # supersample factor
        ssw = btn_w * SS
        ssh = btn_h * SS

        def draw_btn(cx, cy, symbol, hovered, alpha):
            if alpha < 0.02:
                return
            # 色: ホバー時は symbol 別にアクセント色
            if hovered:
                if symbol == "x":
                    bg = (235, 55, 70, int(245 * alpha))
                    border = (200, 30, 45, int(220 * alpha))
                else:  # minimize / expand
                    bg = (60, 130, 220, int(245 * alpha))
                    border = (40, 100, 200, int(220 * alpha))
                fg = (255, 255, 255, int(245 * alpha))
            else:
                bg = (255, 255, 255, int(235 * alpha))
                border = (200, 205, 215, int(180 * alpha))
                fg = (100, 105, 120, int(235 * alpha))

            # 3x 解像度で描画してからダウンスケール
            layer = Image.new("RGBA", (ssw, ssh), (0, 0, 0, 0))
            ld = ImageDraw.Draw(layer)

            # 背景円 (outline の太さも SS 倍しないと縮小後に消える)
            ld.ellipse(
                (0, 0, ssw - 1, ssh - 1),
                fill=bg, outline=border, width=SS,
            )

            pad = (btn_w // 4) * SS
            lw = max(SS, s(2) * SS)
            cx_loc = ssw // 2
            cy_loc = ssh // 2
            if symbol == "x":
                ld.line(
                    [(pad, pad), (ssw - 1 - pad, ssh - 1 - pad)],
                    fill=fg, width=lw,
                )
                ld.line(
                    [(ssw - 1 - pad, pad), (pad, ssh - 1 - pad)],
                    fill=fg, width=lw,
                )
            elif symbol == "minus":
                ld.line(
                    [(pad, cy_loc), (ssw - 1 - pad, cy_loc)],
                    fill=fg, width=lw,
                )
            elif symbol == "plus":
                ld.line(
                    [(pad, cy_loc), (ssw - 1 - pad, cy_loc)],
                    fill=fg, width=lw,
                )
                ld.line(
                    [(cx_loc, pad), (cx_loc, ssh - 1 - pad)],
                    fill=fg, width=lw,
                )

            small = layer.resize((btn_w, btn_h), Image.LANCZOS)
            img.alpha_composite(
                small, dest=(cx - btn_w // 2, cy - btn_h // 2)
            )

        # × (close) — 下部中央
        draw_btn(
            self.close_btn_cx, self.close_btn_cy,
            "x", self.close_btn_hovered, self.close_btn_visible,
        )
        # + (expand) — mini 中のみ (通常時は座標が画面外 + visible が 0)
        if self.is_minimized:
            draw_btn(
                self.expand_btn_cx, self.expand_btn_cy,
                "plus", self.expand_btn_hovered, self.expand_btn_visible,
            )
        # − (minimize) — 通常時のみ (mini 中は座標が画面外 + visible が 0)
        else:
            draw_btn(
                self.min_btn_cx, self.min_btn_cy,
                "minus", self.min_btn_hovered, self.min_btn_visible,
            )

    def _render_glow(self):
        glow = Image.new("RGBA", (self.window_w, self.window_h), (0, 0, 0, 0))
        gd = ImageDraw.Draw(glow)

        breathing = 0.7 + 0.3 * math.sin(self.phase * 1.5)

        if self.transitioning:
            # 遷移中: current_pill_w から t を計算して glow パラメータを lerp
            denom = float(self._normal_pill_w - self._compact_h)
            if denom > 0:
                t = (self.current_pill_w - self._compact_h) / denom
            else:
                t = 0.0
            t = max(0.0, min(1.0, t))
            base_alpha = int(255 * self.glow_intensity * breathing)
            blob_rx = int(s(28) + (s(62) - s(28)) * t)
            blob_ry = int(s(22) + (s(32) - s(22)) * t)
            blur_radius = int(s(14) + (s(22) - s(14)) * t)
            halo_pad = int(s(8) + (s(14) - s(8)) * t)
            wander_speed = 0.55
        elif self.is_minimized:
            # mini: 小さな丸ピルに合わせて glow を縮小。
            # 通常モードの blob/blur サイズをそのまま使うと、小さな window の端で
            # GaussianBlur がクリップされて四角い枠が見えてしまうため。
            base_alpha = int(255 * self.glow_intensity * breathing)
            blob_rx = s(28)
            blob_ry = s(22)
            blur_radius = s(14)
            wander_speed = 0.9 if self.is_recording else 0.55
            halo_pad = s(8)
        elif self.is_recording:
            base_alpha = int(255 * self.glow_intensity * breathing)
            blob_rx = s(70)
            blob_ry = s(38)
            blur_radius = s(26)
            wander_speed = 0.9
            halo_pad = s(14)
        else:
            base_alpha = int(255 * self.glow_intensity * breathing)
            blob_rx = s(62)
            blob_ry = s(32)
            blur_radius = s(22)
            wander_speed = 0.55
            halo_pad = s(14)
        base_alpha = max(0, min(255, base_alpha))

        if self.is_recording:
            colors = [
                RED_DEEP, RED_MAIN, RED_LIGHT, RED_PINK,
                RED_MAIN, RED_LIGHT, RED_DEEP, RED_PINK,
            ]
            halo_color = RED_LIGHT
        else:
            colors = [
                PURPLE_DEEP, PURPLE_MAIN, PURPLE_LIGHT, PURPLE_PINK,
                PURPLE_MAIN, PURPLE_LIGHT, PURPLE_DEEP, PURPLE_PINK,
            ]
            halo_color = PURPLE_LIGHT

        GOLDEN = 1.6180339887
        pill_w_range = self.pill_w - s(60)

        for i, color in enumerate(colors):
            phase_i = self.phase * wander_speed + i * GOLDEN * math.pi
            x_wander = (
                0.45 * math.sin(phase_i * 0.7)
                + 0.30 * math.sin(phase_i * 0.43 + 1.2)
                + 0.15 * math.sin(phase_i * 0.31 + 2.7)
            )
            x_norm = 0.5 + x_wander / 2.0
            x_norm = max(0.0, min(1.0, x_norm))
            cx = self.pill_x0 + s(30) + x_norm * pill_w_range

            y_wander = (
                0.6 * math.sin(phase_i * 0.63 + 0.8)
                + 0.4 * math.sin(phase_i * 0.37 + 1.9)
            )
            if i < 4:
                cy = self.pill_y1 - s(2) + y_wander * s(8)
            else:
                cy = self.pill_y0 + s(2) + y_wander * s(8)

            size_pulse = 1.0 + 0.10 * math.sin(phase_i * 0.51 + i)
            gd.ellipse(
                (cx - blob_rx * size_pulse, cy - blob_ry * size_pulse,
                 cx + blob_rx * size_pulse, cy + blob_ry * size_pulse),
                fill=color + (base_alpha,),
            )

        halo_alpha = int(base_alpha * 0.55)
        gd.rounded_rectangle(
            (self.pill_x0 - halo_pad, self.pill_y0 - halo_pad,
             self.pill_x1 + halo_pad, self.pill_y1 + halo_pad),
            radius=self.pill_radius + halo_pad,
            fill=halo_color + (halo_alpha,),
        )

        glow = glow.filter(ImageFilter.GaussianBlur(radius=blur_radius))
        return glow

    def _draw_pill_and_contents_cairo(self):
        surface = cairo.ImageSurface(
            cairo.FORMAT_ARGB32, self.window_w, self.window_h
        )
        ctx = cairo.Context(surface)
        ctx.set_antialias(cairo.ANTIALIAS_BEST)

        # 白 pill 本体
        _cairo_rounded_rect(
            ctx,
            self.pill_x0, self.pill_y0,
            self.pill_x1 - self.pill_x0, self.pill_y1 - self.pill_y0,
            self.pill_radius,
        )
        ctx.set_source_rgba(1.0, 1.0, 1.0, 248 / 255)
        ctx.fill()

        cy = self.original_cy

        # ミニ中は dot / status / 波形は描かない (マイクだけ)
        if self.is_minimized:
            self._draw_mic_button_cairo(ctx)
            return _cairo_surface_to_pil(surface)

        # ===== ステータスドット (左) =====
        if self.is_recording:
            dot_rgb = RED_MAIN
        else:
            dot_rgb = PURPLE_MAIN

        dot_cx = self.pill_x0 + s(24)
        dot_cy = cy
        dot_r = s(4)
        ctx.arc(dot_cx, dot_cy, dot_r, 0, 2 * math.pi)
        ctx.set_source_rgba(
            dot_rgb[0] / 255, dot_rgb[1] / 255, dot_rgb[2] / 255, 1.0,
        )
        ctx.fill()

        # Status / preview テキストは PIL で別途描画

        # ===== 波形 (round cap 縦線) =====
        # status 文字の幅は固定推定 (PIL が描くので Cairo 側では空ける)
        status_w = s(60) if self.is_recording else s(50)
        wave_x0 = self.pill_x0 + s(24 + 12) + status_w + s(12)
        wave_x1 = self.mic_cx - self.mic_r - s(20)
        wave_cy = cy
        wave_max_h = s(20)

        n = len(self.amps)
        if n > 0:
            total_w = wave_x1 - wave_x0
            unit = total_w / (n * 2 - 1)
            bar_w = unit
            gap = unit

            if self.is_recording:
                r_, g_, b_ = 200, 30, 55
                a_ = 1.0
            else:
                r_, g_, b_ = 140, 145, 165
                a_ = 0.85
            ctx.set_source_rgba(r_ / 255, g_ / 255, b_ / 255, a_)
            ctx.set_line_cap(cairo.LINE_CAP_ROUND)
            ctx.set_line_width(bar_w)

            min_dot = bar_w / 2
            for i, amp in enumerate(self.amps):
                scaled = min(1.0, math.sqrt(max(0, amp) * 2.5))
                bh = max(min_dot, scaled * wave_max_h)
                x = wave_x0 + bar_w / 2 + i * (bar_w + gap)
                ctx.move_to(x, wave_cy - bh)
                ctx.line_to(x, wave_cy + bh)
                ctx.stroke()

        # ===== マイクボタン =====
        self._draw_mic_button_cairo(ctx)

        # × 閉じボタンは PIL で pill 下部にホバー時のみ描画

        return _cairo_surface_to_pil(surface)

    def _draw_mic_button_cairo(self, ctx):
        """マイクボタン (円背景 + ハロー + アイコン) を描画。"""
        if self.is_recording:
            mic_rgb = RED_MAIN
            for hr, alpha in [
                (self.mic_r + s(7), 40),
                (self.mic_r + s(5), 60),
                (self.mic_r + s(3), 80),
            ]:
                ctx.arc(self.mic_cx, self.mic_cy, hr, 0, 2 * math.pi)
                ctx.set_source_rgba(
                    mic_rgb[0] / 255, mic_rgb[1] / 255, mic_rgb[2] / 255,
                    alpha / 255,
                )
                ctx.fill()

            gradient = cairo.RadialGradient(
                self.mic_cx - self.mic_r * 0.3,
                self.mic_cy - self.mic_r * 0.3, 0,
                self.mic_cx, self.mic_cy, self.mic_r,
            )
            lighter = (
                min(255, mic_rgb[0] + 40),
                min(255, mic_rgb[1] + 40),
                min(255, mic_rgb[2] + 40),
            )
            gradient.add_color_stop_rgba(
                0, lighter[0] / 255, lighter[1] / 255, lighter[2] / 255, 1.0,
            )
            gradient.add_color_stop_rgba(
                1, mic_rgb[0] / 255, mic_rgb[1] / 255, mic_rgb[2] / 255, 1.0,
            )
            ctx.arc(self.mic_cx, self.mic_cy, self.mic_r, 0, 2 * math.pi)
            ctx.set_source(gradient)
            ctx.fill()
        else:
            # フラットデザイン
            mic_rgb = PURPLE_MAIN
            ctx.arc(self.mic_cx, self.mic_cy, self.mic_r, 0, 2 * math.pi)
            ctx.set_source_rgba(
                mic_rgb[0] / 255, mic_rgb[1] / 255, mic_rgb[2] / 255, 1.0,
            )
            ctx.fill()

        # マイクアイコン (白カプセル + スタンド + ベース)
        ctx.set_source_rgba(1.0, 1.0, 1.0, 1.0)
        cap_w = s(6)
        cap_h = s(11)
        _cairo_rounded_rect(
            ctx,
            self.mic_cx - cap_w, self.mic_cy - cap_h,
            2 * cap_w, 2 * cap_h - s(3),
            cap_w,
        )
        ctx.fill()
        ctx.set_line_width(s(2))
        ctx.set_line_cap(cairo.LINE_CAP_ROUND)
        ctx.move_to(self.mic_cx, self.mic_cy + s(7))
        ctx.line_to(self.mic_cx, self.mic_cy + s(13))
        ctx.stroke()
        ctx.move_to(self.mic_cx - s(6), self.mic_cy + s(13))
        ctx.line_to(self.mic_cx + s(6), self.mic_cy + s(13))
        ctx.stroke()


# ========== メインアプリ ==========

class VoiceInputApp:
    def __init__(self, config: dict):
        self.config = config
        self.amp_queue: queue.Queue = queue.Queue(maxsize=200)
        self.recorder = Recorder(16000, self.amp_queue)
        # Transcriber は遅延ロード (起動を早くするため)
        self.transcriber: Transcriber | None = None
        self.formatter = Formatter(config)
        self.dict_tracker = DictionaryTracker(USER_DICT_PATH)
        self.is_ready = False
        self.is_recording = False
        self.is_active = True  # ON で起動 (Right Alt が有効)
        self.lock = threading.Lock()
        self.max_duration = config.get("max_duration_sec", 60)
        self._recording_start = 0.0

        # リアルタイム入力関連
        self.realtime_config = config.get("realtime", {})
        self._realtime_stop_event = threading.Event()
        self._realtime_thread: threading.Thread | None = None
        self._llm_task_running = False  # LLM 重複実行防止
        self._last_raw_text = ""  # 前回の raw whisper 結果 (変更検知用)
        self._last_formatted_text = ""  # 前回の LLM 結果 (グロー発火用)

        # ダブルタップ検出
        self.double_tap_ms = config.get("double_tap_ms", 400)
        self._last_alt_press_ms = 0.0
        self._pending_single_tap_timer: threading.Timer | None = None
        self._right_alt_is_down = False  # 未使用 (time-based dedup に移行)
        self._ctrl_is_down = False  # 未使用 (GetAsyncKeyState で直接確認)
        self._last_alt_raw_ms = 0.0  # auto-repeat 抑止用 raw timestamp

        # GUI
        self.root: tk.Tk | None = None
        self.overlay: Overlay | None = None

    def warmup(self):
        """バックグラウンドで Whisper + Ollama を並列ロード。"""
        # オーバーレイ表示は「起動中」のみ (set_status は run() 側で設定済み)

        def load_whisper():
            try:
                print("[Loader] Whisper ロード開始...", flush=True)
                t0 = time.perf_counter()
                self.transcriber = Transcriber(self.config)
                print(f"[Loader] Whisper ロード完了 ({time.perf_counter()-t0:.1f}s)", flush=True)
            except Exception as e:
                print(f"[Loader] Whisper エラー: {e}", flush=True)
                import traceback
                traceback.print_exc()

        def load_ollama():
            try:
                print("[Loader] Ollama ウォームアップ中...", flush=True)
                t0 = time.perf_counter()
                self.formatter.warmup()
                print(f"[Loader] Ollama 完了 ({time.perf_counter()-t0:.1f}s)", flush=True)
            except Exception as e:
                print(f"[Loader] Ollama エラー: {e}", flush=True)

        t_all = time.perf_counter()
        th_w = threading.Thread(target=load_whisper, daemon=True)
        th_o = threading.Thread(target=load_ollama, daemon=True)
        th_w.start()
        th_o.start()
        th_w.join()
        th_o.join()

        if self.transcriber is not None:
            self.is_ready = True
            print(f"[Loader] 準備完了 ✓ (合計 {time.perf_counter()-t_all:.1f}s)", flush=True)
            self._schedule_ui(self._on_ready)
        else:
            print("[Loader] Whisper ロード失敗のため未準備", flush=True)

    def _on_ready(self):
        """ロード完了時に UI 更新"""
        if self.overlay:
            self.overlay.set_status("waiting")

    def _hide_overlay(self):
        """× ボタン押下: バーを隠すがアプリは終了しない"""
        if self.overlay and self.overlay.visible:
            self.overlay.hide()
            if not self.is_recording:
                self.recorder.stop_monitoring()

    # -------- ホットキー処理 (pynput) --------

    def _on_pynput_press(self, key):
        # Right Alt 検出
        is_right_alt = False
        if key == Key.alt_r:
            is_right_alt = True
        elif hasattr(Key, "alt_gr") and key == Key.alt_gr:
            is_right_alt = True
        else:
            try:
                if hasattr(key, "vk") and key.vk == 165:
                    is_right_alt = True
            except Exception:
                pass

        if not is_right_alt:
            return

        # 時間ベース dedup (auto-repeat 抑止)
        now_ms = time.time() * 1000
        if now_ms - self._last_alt_raw_ms < 50:
            return
        self._last_alt_raw_ms = now_ms

        if not self.is_active:
            return

        # Right Alt + Ctrl 同時押し → 終了 (GetAsyncKeyState で直接確認)
        if _is_ctrl_down():
            print("[ホットキー] Right Alt + Ctrl → 終了", flush=True)
            self._schedule_ui(self._quit)
            return

        try:
            self._handle_alt_press()
        except Exception as e:
            print(f"[ERROR] _handle_alt_press: {e}", flush=True)
            import traceback
            traceback.print_exc()

    def _on_pynput_release(self, key):
        # release は今は使ってない (time-based dedup に移行したため)
        pass

    def _handle_alt_press(self):
        now_ms = time.time() * 1000
        delta = now_ms - self._last_alt_press_ms
        self._last_alt_press_ms = now_ms

        if delta < self.double_tap_ms:
            # ダブルタップ: 保留中のシングルタップ処理をキャンセル
            if self._pending_single_tap_timer:
                self._pending_single_tap_timer.cancel()
                self._pending_single_tap_timer = None
            self._last_alt_press_ms = 0  # 次の判定をリセット
            self._schedule_ui(self._on_double_tap)
        else:
            # シングルタップ候補: ダブルタップ猶予後に実行
            if self._pending_single_tap_timer:
                self._pending_single_tap_timer.cancel()
            t = threading.Timer(
                (self.double_tap_ms + 10) / 1000,
                lambda: self._schedule_ui(self._on_single_tap),
            )
            self._pending_single_tap_timer = t
            t.daemon = True
            t.start()

    def _schedule_ui(self, fn):
        """キーボードスレッドから tkinter メインスレッドへディスパッチ"""
        if self.root:
            self.root.after(0, fn)
        else:
            fn()

    def _on_single_tap(self):
        """Right Alt 1回: バー表示中のみ録音開始/停止"""
        if not self.overlay:
            return
        if not self.overlay.visible:
            print("[シングルタップ] バー非表示、無視", flush=True)
            return
        with self.lock:
            if not self.is_recording:
                self._start_recording()
            else:
                self._stop_and_process()

    def _on_button_click(self):
        """バー上のマイクボタン: 録音開始/停止"""
        with self.lock:
            if not self.is_recording:
                self._start_recording()
            else:
                self._stop_and_process()

    def _on_double_tap(self):
        """Right Alt 2回: バー表示/非表示トグル"""
        if not self.overlay:
            return
        # 録音中はバーを隠さない (保護)
        if self.is_recording and self.overlay.visible:
            print("[ダブルタップ] 録音中なので無視", flush=True)
            return
        self.overlay.toggle()
        if self.overlay.visible:
            self.recorder.start_monitoring()
        else:
            self.recorder.stop_monitoring()
        print(f"[オーバーレイ] {'表示' if self.overlay.visible else '非表示'}", flush=True)

    # -------- 録音 --------

    def _start_recording(self):
        if not self.is_ready or self.transcriber is None:
            print("[録音] まだロード中のためスキップ", flush=True)
            if self.overlay:
                self.overlay.set_status("loading...")
            return
        self.is_recording = True
        self._recording_start = time.time()
        self.recorder.start_recording()
        # リアルタイム状態リセット
        self._last_raw_text = ""
        self._last_formatted_text = ""
        if self.overlay:
            if not self.overlay.visible:
                self.overlay.show()
            self.overlay.set_recording(True)
            # preview 状態をリセット
            self.overlay.preview_text = ""
            self.overlay.preview_at = 0.0
            self.overlay.formatted_text = ""
            self.overlay.formatted_at = 0.0
        # リアルタイム preview ワーカー起動
        if self.realtime_config.get("enabled", False):
            self._start_realtime_worker()
        print("\n🎤 録音開始", flush=True)

    def _start_realtime_worker(self):
        """録音中に定期的に Whisper preview を走らせるワーカー"""
        self._realtime_stop_event.clear()
        self._realtime_thread = threading.Thread(
            target=self._realtime_loop, daemon=True,
        )
        self._realtime_thread.start()

    def _stop_realtime_worker(self):
        self._realtime_stop_event.set()
        if self._realtime_thread and self._realtime_thread.is_alive():
            self._realtime_thread.join(timeout=0.5)
        self._realtime_thread = None

    def _llm_background_format(self, raw_text: str):
        """LLM 整形を別スレッドで実行 (preview を止めない)"""
        if self._llm_task_running:
            return  # 既に実行中ならスキップ
        self._llm_task_running = True

        def task():
            try:
                if not self.is_recording:
                    return
                prompt_key = pick_prompt_key(self.config)
                formatted, _ = self.formatter.format(raw_text, prompt_key)
                if not formatted or not self.overlay or not self.is_recording:
                    return
                # 前回と同じならグロー発火しない
                if formatted == self._last_formatted_text:
                    return
                self._last_formatted_text = formatted
                self.overlay.set_formatted_text(formatted)
                print(f"[Preview fmt] {formatted[-30:]}", flush=True)
            except Exception as e:
                print(f"[Realtime] LLM error: {e}", flush=True)
            finally:
                self._llm_task_running = False

        threading.Thread(target=task, daemon=True).start()

    def _realtime_loop(self):
        """
        高頻度 Whisper で preview を更新 (≈0.5秒ごと)。
        LLM 整形はバックグラウンドスレッドで走らせ、preview を止めない。
        """
        interval = self.realtime_config.get("interval_sec", 0.5)
        min_audio = self.realtime_config.get("min_audio_sec", 0.4)
        sr = self.recorder.sample_rate

        while not self._realtime_stop_event.is_set():
            if self._realtime_stop_event.wait(interval):
                break
            if not self.is_recording:
                break
            audio = self.recorder.get_current_audio()
            if len(audio) / sr < min_audio:
                continue

            if self.transcriber is None:
                continue
            # Whisper (raw preview) - ここは同期でOK (≈300ms)
            try:
                raw_text, _ms = self.transcriber.transcribe(audio)
            except Exception as e:
                print(f"[Realtime] Whisper error: {e}", flush=True)
                continue
            if not raw_text:
                continue

            # 前回と同じ raw なら何もしない (無音時の無駄なループ防止)
            if raw_text == self._last_raw_text:
                continue
            self._last_raw_text = raw_text

            if self.overlay:
                self.overlay.set_preview_text(raw_text)
                print(f"[Preview raw] {raw_text[-30:]}", flush=True)

            # LLM 整形をバックグラウンドで (変更があった時だけ)
            self._llm_background_format(raw_text)

    def _stop_and_process(self):
        self.is_recording = False
        # リアルタイムワーカー停止
        self._stop_realtime_worker()

        audio = self.recorder.stop_recording()
        audio_sec = len(audio) / self.recorder.sample_rate
        if self.overlay:
            self.overlay.set_recording(False, "認識中...")
            # preview をクリア
            self.overlay.preview_text = ""
            self.overlay.formatted_text = ""
            self.overlay.preview_at = 0.0
            self.overlay.formatted_at = 0.0
        print(f"⏹ 停止 ({audio_sec:.1f}秒)", flush=True)

        if audio_sec < 0.3:
            print("  → 短すぎ、スキップ", flush=True)
            if self.overlay:
                self.overlay.set_recording(False)
            return

        if self.transcriber is None:
            print("[認識] Transcriber 未ロード、スキップ", flush=True)
            return
        # 認識
        text, whisper_ms = self.transcriber.transcribe(audio)
        print(f"[認識 {whisper_ms:.0f}ms] {text}", flush=True)

        if not text:
            if self.overlay:
                self.overlay.set_recording(False)
            return

        # プロンプト選択
        prompt_key = pick_prompt_key(self.config)
        app_name = get_active_window_process()
        print(f"[アプリ] {app_name} → {prompt_key}", flush=True)

        if self.overlay:
            self.overlay.set_status("整形中...")

        # 短文スキップ
        output_cfg = self.config.get("output", {})
        if (
            output_cfg.get("skip_format_for_short", False)
            and len(text) < output_cfg.get("short_threshold_chars", 8)
        ):
            formatted = text
            print(f"[整形スキップ]", flush=True)
        else:
            formatted, llm_ms = self.formatter.format(text, prompt_key)
            print(f"[整形 {llm_ms:.0f}ms] {formatted}", flush=True)

        # 辞書 hit tracking (最終出力テキストに含まれる term を +1)
        try:
            self.dict_tracker.track_hits(formatted)
        except Exception as e:
            print(f"[DictTracker] track_hits error: {e}", flush=True)

        # 出力
        output_text(formatted, self.config)
        print("✅ 出力完了", flush=True)

        # 辞書 flush (hit があったときだけ書き戻す)
        threading.Thread(
            target=self.dict_tracker.flush, daemon=True
        ).start()

        if self.overlay:
            self.overlay.set_recording(False)

    def _watchdog(self):
        while True:
            time.sleep(1)
            if self.is_recording:
                if time.time() - self._recording_start > self.max_duration:
                    print(f"\n⚠ {self.max_duration}秒超過、自動停止", flush=True)
                    self._schedule_ui(self._on_single_tap)

    # -------- 実行 --------

    def run(self):
        # Tkinter 起動 (root は Toplevel の親に使うだけで非表示)
        self.root = tk.Tk()
        self.root.withdraw()  # 空の "tk" ウィンドウを隠す

        self.overlay = Overlay(
            self.root,
            self.config,
            self.amp_queue,
            on_button_click=self._on_button_click,
            on_close_click=lambda: self._schedule_ui(self._quit),
        )
        # 起動直後: ロード中の状態でバーを即表示
        self.overlay.set_status("起動中")
        self.overlay.show()

        # pynput キーリスナー
        try:
            self.kb_listener = pkb.Listener(
                on_press=self._on_pynput_press,
                on_release=self._on_pynput_release,
                suppress=False,
            )
            self.kb_listener.start()
            print(f"[Listener] pynput 起動 alive={self.kb_listener.is_alive()}", flush=True)
        except Exception as e:
            print(f"[Listener] 起動失敗: {e}", flush=True)
            import traceback
            traceback.print_exc()

        # システムトレイアイコン
        self._setup_tray()

        # ウォッチドッグ
        threading.Thread(target=self._watchdog, daemon=True).start()

        # バックグラウンドで Whisper + Ollama ロード
        threading.Thread(target=self.warmup, daemon=True).start()

        print("\n" + "=" * 60)
        print("🎙 音声入力ツール 起動")
        print("   Right Alt 2回     : バー表示/非表示")
        print("   Right Alt 1回     : 録音開始/停止 (バー表示中のみ)")
        print("   Right Alt + Ctrl  : 終了")
        print("   バー上のマイク   : 録音開始/停止")
        print("   バー上の ×      : バー非表示")
        print("=" * 60 + "\n", flush=True)

        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            pass
        print("\n終了します", flush=True)

    def _setup_tray(self):
        if not HAS_TRAY:
            print("[Tray] pystray 未インストール、スキップ", flush=True)
            return
        try:
            icon_path = Path(__file__).parent / "mic.png"
            if not icon_path.exists():
                print(f"[Tray] アイコン未発見: {icon_path}", flush=True)
                return
            self._tray_icon_normal = Image.open(icon_path)
            rec_path = Path(__file__).parent / "mic_rec.png"
            self._tray_icon_rec = Image.open(rec_path) if rec_path.exists() else self._tray_icon_normal

            menu = pystray.Menu(
                pystray.MenuItem(
                    "バー表示/非表示",
                    lambda icon, item: self._schedule_ui(self._tray_toggle_bar),
                    default=True,
                ),
                pystray.MenuItem(
                    "録音開始/停止",
                    lambda icon, item: self._schedule_ui(self._on_button_click),
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    "終了",
                    lambda icon, item: self._schedule_ui(self._quit),
                ),
            )

            self.tray_icon = pystray.Icon(
                "voice_input",
                self._tray_icon_normal,
                "音声入力ツール",
                menu,
            )
            t = threading.Thread(target=self.tray_icon.run, daemon=True)
            t.start()
            print("[Tray] システムトレイ起動", flush=True)
        except Exception as e:
            print(f"[Tray] 起動失敗: {e}", flush=True)
            import traceback
            traceback.print_exc()

    def _tray_toggle_bar(self):
        if not self.overlay:
            return
        if self.is_recording:
            return  # 録音中はトグル無効
        self.overlay.toggle()
        if self.overlay.visible:
            self.recorder.start_monitoring()
        else:
            self.recorder.stop_monitoring()

    def _quit(self):
        # Ollama の Qwen を即座にアンロード (VRAM 解放)
        try:
            payload = json.dumps({
                "model": self.formatter.llm_config["model"],
                "keep_alive": 0,
            }).encode("utf-8")
            req = urllib.request.Request(
                self.formatter.llm_config["ollama_url"],
                data=payload,
                headers={"Content-Type": "application/json"},
            )
            urllib.request.urlopen(req, timeout=2)
            print("[Quit] Ollama アンロード要求送信", flush=True)
        except Exception as e:
            print(f"[Quit] Ollama アンロード失敗: {e}", flush=True)
        try:
            self.dict_tracker.flush()
        except Exception:
            pass
        try:
            self.recorder.stop_monitoring()
        except Exception:
            pass
        try:
            if hasattr(self, "kb_listener") and self.kb_listener:
                self.kb_listener.stop()
        except Exception:
            pass
        try:
            if hasattr(self, "tray_icon") and self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass
        if self.root:
            self.root.quit()
            self.root.destroy()


def main():
    if not acquire_single_instance():
        print("[ERROR] 既に音声入力ツールが起動しています。重複起動を防止しました。", flush=True)
        sys.exit(1)
    config = load_config()
    app = VoiceInputApp(config)
    app.run()


if __name__ == "__main__":
    main()
