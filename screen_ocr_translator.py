#!/usr/bin/env python3
"""
Screen OCR Translator: press a hotkey, drag a box over anything on screen, and get
the text, a translation, a short summary, an Excel-ready table and the image itself.

README
  Install : py -3.12 -m pip install -r requirements.txt     (Linux: python3 -m pip install -r requirements.txt)
  Setup   : py -3.12 screen_ocr_translator.py --install     (tiny hotkey listener, now and at every login;
            --uninstall removes it). Or run the app directly: py -3.12 screen_ocr_translator.py
  Hotkey  : Ctrl+Shift+T starts the app in capture mode; drag a box, Esc or right-click cancels.
            Closing the window hides it (next capture is fast); "Quit" stops it. --selftest checks logic.
  First run: EasyOCR models (~100 MB + ~20 MB/language) and Argos packs (~100 MB/pair) download once,
             then it works offline. Optional AI summary: install Ollama, run `ollama pull llama3.2`.
  Trouble : Box misaligned -> DPI scaling; use python.org Python. Moved this file? Run --install again.
            Fullscreen games / DRM video capture as black. First capture per language is slow (model load).
            Log: %TEMP%\\screen_ocr_translator.log. Linux: hotkey needs X11; image copy needs xclip/wl-copy.
"""

from __future__ import annotations

import argparse
import ctypes
import functools
import importlib.util
import io
import json
import queue
import re
import shutil
import socket
import statistics
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable, Optional, Sequence

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:  # --selftest still works without Tk
    tk = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]

if TYPE_CHECKING:
    from PIL.Image import Image as PILImage


# =============================================================================
# Constants
# =============================================================================

HOTKEY_KEY = "t"  # used with Ctrl+Shift
HOTKEY_LABEL = "Ctrl+Shift+T"
OLLAMA_MODEL: Optional[str] = "llama3.2"  # set to None to always use "Key sentences"
OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT_S = 120      # CPU-only laptops can take a while, especially the first run
OLLAMA_START_WAIT_S = 15    # how long to wait for an auto-started Ollama server

UI_POLL_MS = 100
MIN_SELECTION_PX = 5
UPSCALE_BELOW_PX = 100
OCR_MIN_BOX_PX = 8  # EasyOCR default 20 drops lone digits like table cells "3"
OVERLAY_BRIGHTNESS = 0.45
THUMBNAIL_SIZE = (360, 220)
TEMP_DIR = Path(tempfile.gettempdir()) / "screen_ocr_translator"
LOG_PATH = Path(tempfile.gettempdir()) / "screen_ocr_translator.log"

# Background listener <-> app communication (localhost only)
APP_PORT = 47821        # running app accepts "capture"; also its single-instance lock
LISTENER_PORT = 47822   # listener accepts "ping" / "stop"; also its single-instance lock
LAUNCH_DEBOUNCE_S = 3.0
AUTOSTART_NAME = "ScreenOCRTranslator"
WINDOWS_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
# Windows' Startup-apps on/off switch for Run entries. An entry without a record here can be
# skipped at login, so --install marks it enabled (0x02 + 11 zero bytes, as Settings writes).
WINDOWS_STARTUP_APPROVED_KEY = r"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
STARTUP_ENABLED = b"\x02" + b"\x00" * 11
LINUX_AUTOSTART_FILE = Path.home() / ".config" / "autostart" / "screen-ocr-translator.desktop"

# Look and feel
FONT_FAMILY = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO_FAMILY = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"
THEME: dict[str, str] = {  # "aurora": deep navy with violet -> pink -> amber accents
    "bg": "#0A0D19", "card": "#131728", "card_border": "#252B4A", "field": "#1A1F37",
    "field_border": "#2E3562", "text": "#EEF0FF", "muted": "#8E94BA", "faint": "#565C85",
    "select": "#43347E", "ok": "#34D399", "busy": "#FBBF24", "err": "#F87171",
    "err_soft": "#3B1B2B", "soft": "#1E2342", "soft_hover": "#2A3160", "link": "#C4B5FD",
    "thumb": "#363D6C", "thumb_hover": "#4E5596",
}
GRADIENT = ("#8B5CF6", "#EC4899", "#F59E0B")
CARD_ACCENTS: dict[str, tuple[str, str]] = {  # icon gradient per card
    "original": ("#22D3EE", "#3B82F6"), "translation": ("#A78BFA", "#7C3AED"),
    "summary": ("#F472B6", "#DB2777"), "table": ("#34D399", "#059669"),
    "image": ("#FBBF24", "#F97316"), "settings": ("#818CF8", "#6366F1"),
}
MIN_PANE_PX = 110  # smallest height a resizable row of boxes can be dragged to
CARD_MARGIN = 9   # shadow room around each card
CARD_RADIUS = 18

LINE_GROUP_FACTOR = 0.5      # same line if centers differ < 0.5 x median box height
PARAGRAPH_GAP_FACTOR = 1.5   # blank line if line centers differ > 1.5 x median line height
TABLE_CELL_GAP_FACTOR = 1.0  # new table cell if horizontal gap > 1.0 x median box height

AUTO_DETECT = "Auto-detect"
DEFAULT_SOURCE = AUTO_DETECT
DEFAULT_TARGET = "English"

# Display name -> (EasyOCR code, Argos code)
LANGUAGES: dict[str, tuple[str, str]] = {
    "English": ("en", "en"),
    "Hindi": ("hi", "hi"),
    "Spanish": ("es", "es"),
    "French": ("fr", "fr"),
    "German": ("de", "de"),
    "Portuguese": ("pt", "pt"),
    "Italian": ("it", "it"),
    "Russian": ("ru", "ru"),
    "Arabic": ("ar", "ar"),
    "Japanese": ("ja", "ja"),
    "Korean": ("ko", "ko"),
    "Chinese (Simplified)": ("ch_sim", "zh"),
    "Turkish": ("tr", "tr"),
    "Urdu": ("ur", "ur"),
}

# Writing systems: EasyOCR reader languages for each, and which script each language uses.
SCRIPT_READER_LANGS: dict[str, list[str]] = {
    "latin": ["en", "es", "fr", "de", "pt", "it", "tr"],
    "cyrillic": ["ru", "en"],
    "devanagari": ["hi", "en"],
    "arabic": ["ar", "ur", "en"],
    "japanese": ["ja", "en"],
    "korean": ["ko", "en"],
    "chinese": ["ch_sim", "en"],
}
SCRIPT_OF_LANGUAGE: dict[str, str] = {
    "English": "latin", "Spanish": "latin", "French": "latin", "German": "latin",
    "Portuguese": "latin", "Italian": "latin", "Turkish": "latin", "Russian": "cyrillic",
    "Hindi": "devanagari", "Arabic": "arabic", "Urdu": "arabic", "Japanese": "japanese",
    "Korean": "korean", "Chinese (Simplified)": "chinese",
}
SCRIPT_CHAR_RANGES: dict[str, str] = {  # letters native to each script (for scoring OCR output)
    "latin": "A-Za-zÀ-ɏ",
    "cyrillic": "Ѐ-ӿ",
    "devanagari": "ऀ-ॿ",
    "arabic": "؀-ۿݐ-ݿ",
    "japanese": "぀-ヿ一-鿿",
    "korean": "가-힯ᄀ-ᇿ",
    "chinese": "一-鿿",
}
LATIN_ACCEPT_CONFIDENCE = 0.65  # Latin read this confidently -> skip trying other scripts
DETECT_SAMPLE_LINES = 4         # scripts compete on this many of the widest text lines
SCRIPT_ACCEPT_SCORE = 0.45      # ends the search early (measured: right script 0.54-0.91,
#                                 wrong scripts never above 0.20)
NON_LATIN_TRY_ORDER = ["devanagari", "cyrillic", "chinese", "japanese", "korean", "arabic"]
#   Arabic last: its reader is ~4x slower than the others on CPU.

SENTENCE_RE = re.compile(r"[^.!?।。！？]+[.!?।。！？]*")
CJK_RANGES = "぀-ヿ㐀-䶿一-鿿豈-﫿"
COMBINING_MARKS = "̀-ͯऀ-෿ً-ٰٟ"  # keeps Hindi/Urdu words whole
CJK_RE = re.compile(f"[{CJK_RANGES}]")
TOKEN_RE = re.compile(f"[{CJK_RANGES}]|(?:[^\\W{CJK_RANGES}]|[{COMBINING_MARKS}])+")
CJK_SENTENCE_ENDS = "。！？"

TranslateFn = Callable[[str, str, str], str]
StatusFn = Callable[[str], None]
EasyOcrResult = Sequence[tuple[Sequence[Sequence[float]], str, float]]


# =============================================================================
# Layout reconstruction (pure Python, unit-tested by --selftest)
# =============================================================================

@dataclass(frozen=True)
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def height(self) -> float:
        return self.y1 - self.y0


def words_from_easyocr(results: EasyOcrResult) -> list[Word]:
    words: list[Word] = []
    for bbox, text, _conf in results:
        text = text.strip()
        if not text:
            continue
        xs = [float(p[0]) for p in bbox]
        ys = [float(p[1]) for p in bbox]
        words.append(Word(min(xs), min(ys), max(xs), max(ys), text))
    return words


def group_lines(words: list[Word]) -> list[list[Word]]:
    """Group boxes into lines by vertical center, then sort each line left to right."""
    if not words:
        return []
    tolerance = LINE_GROUP_FACTOR * statistics.median(w.height for w in words)
    lines: list[list[Word]] = []
    centers: list[float] = []
    for word in sorted(words, key=lambda w: w.yc):
        if lines and abs(word.yc - centers[-1]) < tolerance:
            lines[-1].append(word)
            centers[-1] = statistics.fmean(w.yc for w in lines[-1])
        else:
            lines.append([word])
            centers.append(word.yc)
    return [sorted(line, key=lambda w: w.x0) for line in lines]


def lines_to_text(lines: list[list[Word]]) -> str:
    """Join lines with newlines; insert a blank line at large vertical gaps (paragraphs)."""
    if not lines:
        return ""
    heights = [max(w.y1 for w in line) - min(w.y0 for w in line) for line in lines]
    gap_limit = PARAGRAPH_GAP_FACTOR * statistics.median(heights)
    out: list[str] = []
    prev_center: Optional[float] = None
    for line in lines:
        center = statistics.fmean(w.yc for w in line)
        if prev_center is not None and center - prev_center > gap_limit:
            out.append("")
        out.append(" ".join(w.text for w in line))
        prev_center = center
    return "\n".join(out)


def detect_table(lines: list[list[Word]]) -> Optional[list[list[str]]]:
    """Return rows x columns of cell text if the lines look like a table, else None."""
    if len(lines) < 2:
        return None
    cell_gap = TABLE_CELL_GAP_FACTOR * statistics.median(w.height for line in lines for w in line)

    rows: list[list[Word]] = []
    for line in lines:
        cells: list[Word] = []
        for w in line:
            if cells and w.x0 - cells[-1].x1 <= cell_gap:
                last = cells[-1]
                cells[-1] = Word(last.x0, min(last.y0, w.y0), max(last.x1, w.x1),
                                 max(last.y1, w.y1), f"{last.text} {w.text}")
            else:
                cells.append(w)
        rows.append(cells)

    multi = [r for r in rows if len(r) >= 2]
    if len(multi) < 2 or len(multi) * 2 < len(rows):
        return None

    # Columns = union of overlapping x-ranges of cells from multi-cell rows.
    spans = sorted((c.x0, c.x1) for r in multi for c in r)
    columns: list[list[float]] = [list(spans[0])]
    for x0, x1 in spans[1:]:
        if x0 <= columns[-1][1]:
            columns[-1][1] = max(columns[-1][1], x1)
        else:
            columns.append([x0, x1])
    if len(columns) < 2:
        return None

    def overlap(cell: Word, col: list[float]) -> float:
        return min(cell.x1, col[1]) - max(cell.x0, col[0])  # negative = distance

    table: list[list[str]] = []
    for r in rows:
        row = [""] * len(columns)
        for cell in r:
            idx = max(range(len(columns)), key=lambda i: overlap(cell, columns[i]))
            row[idx] = f"{row[idx]} {cell.text}".strip()
        table.append(row)
    return table


def table_to_tsv(rows: list[list[str]]) -> str:
    return "\n".join("\t".join(cell.replace("\t", " ").replace("\n", " ") for cell in row)
                     for row in rows)


def reconstruct(results: EasyOcrResult) -> tuple[str, Optional[list[list[str]]]]:
    lines = group_lines(words_from_easyocr(results))
    return lines_to_text(lines), detect_table(lines)


# =============================================================================
# Extractive summarizer (pure Python)
# =============================================================================

def split_sentences(text: str) -> list[str]:
    return [" ".join(s.split()) for s in SENTENCE_RE.findall(text) if s.strip()]


def tokenize(text: str) -> list[str]:
    return [t.lower() for t in TOKEN_RE.findall(text)]


def content_tokens(text: str) -> list[str]:
    return [t for t in tokenize(text) if CJK_RE.fullmatch(t) or len(t) > 2]


def join_sentences(sentences: list[str]) -> str:
    out = ""
    for s in sentences:
        if out and out[-1] not in CJK_SENTENCE_ENDS:
            out += " "
        out += s
    return out


def extractive_summary(text: str, max_sentences: int = 3) -> str:
    sentences = split_sentences(text)
    if len(sentences) <= 2:
        return text.strip()
    freq: dict[str, int] = {}
    for tok in content_tokens(text):
        freq[tok] = freq.get(tok, 0) + 1

    def score(sentence: str) -> float:
        toks = content_tokens(sentence)
        return sum(freq[t] for t in toks) / len(toks) if toks else 0.0

    count = min(max_sentences, max(1, (len(sentences) + 1) // 2))
    best = sorted(range(len(sentences)), key=lambda i: score(sentences[i]), reverse=True)[:count]
    return join_sentences([sentences[i] for i in sorted(best)])


# =============================================================================
# Translation logic (pure Python; Argos is plugged in as translate_fn)
# =============================================================================

def required_pairs(src: str, tgt: str) -> list[tuple[str, str]]:
    """Language pairs needed for src -> tgt, pivoting through English if neither is English."""
    if src == tgt:
        return []
    if src != "en" and tgt != "en":
        return [(src, "en"), ("en", tgt)]
    return [(src, tgt)]


def translate_with_pivot(text: str, src: str, tgt: str, translate_fn: TranslateFn) -> str:
    for a, b in required_pairs(src, tgt):  # one separate call per leg
        text = translate_fn(text, a, b)
    return text


def join_broken_lines(lines: list[str]) -> str:
    out = ""
    for line in (ln.strip() for ln in lines):
        if not line:
            continue
        if out and not (CJK_RE.match(out[-1]) and CJK_RE.match(line[0])):
            out += " "
        out += line
    return out


def translate_document(text: str, src: str, tgt: str, join_lines: bool,
                       translate_fn: TranslateFn) -> str:
    """Translate paragraph by paragraph (lines joined) or line by line, keeping blank lines."""
    if src == tgt:
        return text
    result: list[str] = []
    for para in re.split(r"\n\s*\n", text.strip()):
        lines = para.splitlines()
        if join_lines:
            result.append(translate_with_pivot(join_broken_lines(lines), src, tgt, translate_fn))
        else:
            result.append("\n".join(
                translate_with_pivot(ln.strip(), src, tgt, translate_fn) if ln.strip() else ""
                for ln in lines))
    return "\n\n".join(result)


# =============================================================================
# Language auto-detection (pure Python)
# =============================================================================

# Common words that are frequent in one Latin-script language and rare in the others.
LATIN_STOPWORDS: dict[str, set[str]] = {
    "English": {"the", "and", "of", "to", "is", "in", "that", "it", "for", "with", "was", "are",
                "this", "you", "have", "on", "be", "not", "they", "from", "by", "or", "which"},
    "Spanish": {"el", "la", "los", "las", "de", "que", "y", "en", "un", "una", "es", "por", "con",
                "para", "del", "se", "no", "más", "pero", "como", "muy", "también", "está"},
    "French": {"le", "la", "les", "de", "des", "et", "est", "un", "une", "que", "qui", "dans",
               "pour", "pas", "sur", "au", "du", "avec", "ce", "sont", "nous", "vous", "plus"},
    "German": {"der", "die", "das", "und", "ist", "nicht", "ein", "eine", "zu", "den", "mit",
               "von", "sich", "auf", "für", "dem", "ich", "auch", "werden", "sind", "wird",
               "zur", "zum", "im", "vom", "bei", "nach", "aus", "über", "oder", "aber", "wie",
               "wenn", "noch", "nur", "kein", "wurde", "sie", "wir", "ihre", "ihr", "ab"},
    "Portuguese": {"o", "a", "os", "as", "de", "que", "e", "do", "da", "em", "um", "uma", "para",
                   "com", "não", "é", "por", "mais", "dos", "das", "são", "muito", "também",
                   "ao", "aos", "nos", "na", "pelo", "pela", "seu", "sua", "você", "fica"},
    "Italian": {"il", "lo", "la", "gli", "le", "di", "che", "e", "è", "un", "una", "per", "non",
                "con", "del", "della", "sono", "anche", "più", "nel", "molto", "questo",
                "al", "alla", "sulla", "nella", "dei", "nei", "delle", "questa"},
    "Turkish": {"ve", "bir", "bu", "da", "de", "için", "ile", "çok", "ne", "gibi", "daha", "olan",
                "var", "ama", "değil", "olarak", "kadar", "en", "her", "ise", "ben"},
}
# Letters that point to particular languages (weight per occurrence).
LATIN_LETTER_HINTS: dict[str, dict[str, float]] = {
    "ñ": {"Spanish": 3}, "¿": {"Spanish": 3}, "¡": {"Spanish": 3},
    "ã": {"Portuguese": 3}, "õ": {"Portuguese": 3}, "ß": {"German": 3}, "ä": {"German": 2},
    "ğ": {"Turkish": 3}, "ş": {"Turkish": 3}, "ı": {"Turkish": 3}, "œ": {"French": 3},
    "ê": {"French": 1, "Portuguese": 1}, "î": {"French": 1.5}, "û": {"French": 1.5},
    "ô": {"French": 1, "Portuguese": 1}, "ò": {"Italian": 2}, "ì": {"Italian": 2},
    "ç": {"French": 1, "Portuguese": 1, "Turkish": 1}, "ü": {"German": 1, "Turkish": 1},
    "ö": {"German": 1, "Turkish": 1}, "é": {"French": 0.6, "Spanish": 0.3, "Portuguese": 0.3},
    "á": {"Spanish": 0.7, "Portuguese": 0.7}, "í": {"Spanish": 0.7, "Portuguese": 0.5},
    "ó": {"Spanish": 0.7, "Portuguese": 0.5}, "ú": {"Spanish": 0.7, "Portuguese": 0.5},
    "è": {"French": 0.6, "Italian": 0.6}, "à": {"French": 0.6, "Italian": 0.6, "Portuguese": 0.3},
}
# Letter patterns typical of each language; they occur inside almost every word, so even short
# labels like "Einstellungen speichern" get a signal. Spaces mark word starts/ends.
LATIN_NGRAMS: dict[str, dict[str, float]] = {
    "English": {"th": 1.0, "wh": 1.2, "ing": 1.5, "ght": 2, "ly ": 0.8, "ow": 0.5, "ea": 0.4,
                "you": 1, "sh": 0.6, "y ": 0.6},
    "German": {"sch": 1.5, "cht": 2, "ung": 1.5, "eit": 1.2, "ich": 1, "ein": 0.8, "ei": 0.5,
               "ie": 0.4, "tz": 1.5, "chen": 1, "ck": 0.5, "en ": 0.4, "er ": 0.3, "k": 0.25},
    "Spanish": {"ción": 3, "ión": 1, "ll": 0.5, "rr": 0.8, "ía": 1.2, "ado": 0.8, "ido": 0.6,
                "os ": 0.6, "dad": 1.5, "mente": 1.2, "ar ": 0.6},
    "French": {"eau": 2, "oi": 0.8, "ou": 0.4, "qu": 0.5, "ent ": 0.8, "aux ": 1.5, "ais": 0.8,
               "ez ": 1.5, " l'": 0.8, " d'": 1, "qu'": 1.5, "ètre": 1.5, "eur ": 1, "ée": 1.5},
    "Portuguese": {"ção": 3, "ões": 3, "ão": 2, "nh": 1.2, "lh": 1.2, "em ": 0.5, "do ": 0.4,
                   "ade ": 1, "inho": 1.5},
    "Italian": {"zion": 2.5, "ione": 1, "gli": 1.5, "gn": 0.8, "cch": 1.5, "zz": 1.2, "tt": 0.5,
                "oni ": 0.8, "che ": 0.6, "iche ": 1.2, "zi": 0.4, "gg": 1.2},
    "Turkish": {"lar": 1.5, "ler": 1.5, "iyor": 2, "mak": 0.8, "dir": 0.5},
}
LATIN_MIN_SCORE = 0.8  # weaker evidence than this -> English
URDU_ONLY_LETTERS = set("ٹڈڑںےۓھہگچپژک")  # used in Urdu but not in Arabic
KANA_RE = re.compile("[぀-ヿ]")
LATIN_WORD_RE = re.compile(r"[^\W\d_]+")


def detect_latin_language(text: str) -> str:
    """Guess which Latin-script language text is in; English when there are no clues."""
    scores = {lang: 0.0 for lang in LATIN_STOPWORDS}
    for word in LATIN_WORD_RE.findall(text.lower()):
        for lang, words in LATIN_STOPWORDS.items():
            if word in words:
                scores[lang] += 1
    for char in text.lower():
        for lang, weight in LATIN_LETTER_HINTS.get(char, {}).items():
            scores[lang] += weight
    padded = " " + re.sub(r"[^\w']+", " ", text.lower()).strip() + " "
    for lang, grams in LATIN_NGRAMS.items():
        scores[lang] += sum(weight * padded.count(gram) for gram, weight in grams.items())
    best = max(scores, key=lambda lang: scores[lang])
    return best if scores[best] >= LATIN_MIN_SCORE else "English"


def script_score(script: str, text: str, confidence: float) -> float:
    """OCR confidence scaled by how much of the output is letters of that script."""
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return 0.0
    native = re.findall(f"[{SCRIPT_CHAR_RANGES[script]}]", "".join(letters))
    return confidence * len(native) / len(letters)


def choose_script(candidates: dict[str, tuple[str, float]]) -> str:
    """Pick the writing system whose reader produced the most confident native output.
    candidates: script -> (recognized text, mean confidence)."""
    best = max(candidates, key=lambda s: script_score(s, *candidates[s]))
    japanese = candidates.get("japanese")
    if best == "chinese" and japanese and len(KANA_RE.findall(japanese[0])) >= 2:
        return "japanese"  # kanji look alike; kana only occur in Japanese
    return best


def language_for_script(script: str, text: str) -> str:
    """Map a detected writing system (plus the text read with it) to a language name."""
    if script == "latin":
        return detect_latin_language(text)
    if script == "arabic":
        return "Urdu" if any(c in URDU_ONLY_LETTERS for c in text) else "Arabic"
    if script == "japanese":
        return "Japanese" if KANA_RE.search(text) else "Chinese (Simplified)"
    return {"cyrillic": "Russian", "devanagari": "Hindi", "korean": "Korean",
            "chinese": "Chinese (Simplified)"}[script]


def mean_confidence(results: EasyOcrResult) -> float:
    """Confidence averaged over characters, so long words count more than stray marks."""
    total = sum(len(text) for _, text, _ in results)
    return sum(conf * len(text) for _, text, conf in results) / total if total else 0.0


# =============================================================================
# Summary (Ollama first, extractive fallback)
# =============================================================================

def _local_opener() -> urllib.request.OpenerDirector:
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))  # never proxy localhost


def _ollama_up() -> bool:
    try:
        with _local_opener().open(OLLAMA_URL.rsplit("/api/", 1)[0] + "/api/version", timeout=1):
            return True
    except Exception:
        return False


def ensure_ollama_running() -> bool:
    """Start the Ollama server in the background if it is installed but not running."""
    if _ollama_up():
        return True
    exe = shutil.which("ollama")
    default = Path.home() / "AppData" / "Local" / "Programs" / "Ollama" / "ollama.exe"
    if exe is None and sys.platform == "win32" and default.exists():
        exe = str(default)
    if exe is None:
        return False
    kwargs: dict[str, Any] = {"stdin": subprocess.DEVNULL, "stdout": subprocess.DEVNULL,
                              "stderr": subprocess.DEVNULL, "close_fds": True}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x00000008 | 0x00000200  # DETACHED_PROCESS | NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    try:
        subprocess.Popen([exe, "serve"], **kwargs)
    except OSError:
        return False
    deadline = time.monotonic() + OLLAMA_START_WAIT_S
    while time.monotonic() < deadline:
        time.sleep(0.25)
        if _ollama_up():
            return True
    return False


def ollama_summary(text: str, language: str) -> Optional[str]:
    if not OLLAMA_MODEL or not ensure_ollama_running():
        return None
    prompt = (f"Summarize the following text in your own words in 2-3 sentences, written in "
              f"{language}. Do not copy sentences from the text. Output only the summary."
              f"\n\n{text}")
    body = json.dumps({"model": OLLAMA_MODEL, "prompt": prompt, "stream": False}).encode("utf-8")
    request = urllib.request.Request(OLLAMA_URL, data=body,
                                     headers={"Content-Type": "application/json"})
    opener = _local_opener()
    try:
        with opener.open(request, timeout=OLLAMA_TIMEOUT_S) as response:
            summary = json.loads(response.read().decode("utf-8")).get("response", "").strip()
    except Exception:
        return None
    return summary or None


def summarize(text: str, language: str) -> tuple[str, str]:
    if not text.strip():
        return "", ""
    ai = ollama_summary(text, language)
    if ai:
        return ai, "AI summary"
    return extractive_summary(text), "Key sentences"


# =============================================================================
# OCR engine (EasyOCR, lazy and cached per language)
# =============================================================================

def cuda_available() -> bool:
    try:
        import torch
        return bool(torch.cuda.is_available())
    except Exception:
        return False


class OcrEngine:
    """One EasyOCR reader per writing system, loaded lazily. The Latin reader also owns the
    text-box detector; the others only recognize text inside boxes it found."""

    def __init__(self) -> None:
        self._readers: dict[str, Any] = {}
        self._lock = threading.Lock()

    def is_loaded(self, script: str) -> bool:
        return script in self._readers

    def _reader(self, script: str) -> Any:
        with self._lock:
            if script not in self._readers:
                import easyocr
                self._readers[script] = easyocr.Reader(
                    SCRIPT_READER_LANGS[script], gpu=cuda_available(), verbose=False,
                    detector=script == "latin")
            return self._readers[script]

    @staticmethod
    def _prepare(image: PILImage) -> Any:
        import numpy as np
        from PIL import Image

        img = image.convert("RGB")
        if img.height < UPSCALE_BELOW_PX:
            img = img.resize((img.width * 2, img.height * 2), Image.Resampling.LANCZOS)
        return np.asarray(img)

    def _detect(self, pixels: Any) -> tuple[list, list]:
        horizontal, free = self._reader("latin").detect(pixels, min_size=OCR_MIN_BOX_PX)
        return horizontal[0], free[0]

    def _recognize(self, script: str, pixels: Any, boxes: tuple[list, list]) -> EasyOcrResult:
        from easyocr.utils import reformat_input
        _, grey = reformat_input(pixels)
        return self._reader(script).recognize(grey, horizontal_list=boxes[0], free_list=boxes[1])

    def read(self, image: PILImage, language: str) -> tuple[str, Optional[list[list[str]]]]:
        pixels = self._prepare(image)
        boxes = self._detect(pixels)
        if not boxes[0] and not boxes[1]:
            return "", None
        return reconstruct(self._recognize(SCRIPT_OF_LANGUAGE[language], pixels, boxes))

    def read_auto(self, image: PILImage, status: StatusFn
                  ) -> tuple[str, Optional[list[list[str]]], Optional[str]]:
        """Read text in an unknown language. Returns (text, table, detected language name)."""
        pixels = self._prepare(image)
        boxes = self._detect(pixels)
        if not boxes[0] and not boxes[1]:
            return "", None, None
        results = {"latin": self._recognize("latin", pixels, boxes)}
        latin_text = " ".join(t for _, t, _ in results["latin"])
        script = "latin"
        if script_score("latin", latin_text, mean_confidence(results["latin"])) < LATIN_ACCEPT_CONFIDENCE:
            # Not clearly Latin: every script competes on a few lines, the winner reads them all.
            sample = (sorted(boxes[0], key=lambda b: b[1] - b[0], reverse=True)[:DETECT_SAMPLE_LINES],
                      boxes[1][:DETECT_SAMPLE_LINES])
            candidates = {"latin": (latin_text, mean_confidence(results["latin"]))}
            for name in NON_LATIN_TRY_ORDER:  # cheapest/likeliest first; stop once one is sure
                if not self.is_loaded(name):
                    status(f"Detecting language... loading the {name} reader (first time only)")
                read = self._recognize(name, pixels, sample)
                candidates[name] = (" ".join(t for _, t, _ in read), mean_confidence(read))
                best = max(script_score(n, *c) for n, c in candidates.items() if n != "latin")
                if name != "chinese" and best >= SCRIPT_ACCEPT_SCORE:
                    break  # (never right after Chinese: Japanese, tried next, may own the kanji)
            script = choose_script(candidates)
            if script not in results:
                results[script] = self._recognize(script, pixels, boxes)
        text, table = reconstruct(results[script])
        return text, table, language_for_script(script, text)


# =============================================================================
# Translator (Argos, auto-installs packs)
# =============================================================================

class PackNotFoundError(Exception):
    pass


class ArgosTranslator:
    """Locks are per language pair, so a slow pack download (100-300 MB) only blocks work that
    needs that same pack; translating with another, installed pair goes ahead meanwhile."""

    def __init__(self) -> None:
        self._guard = threading.Lock()          # protects _pair_locks
        self._pair_locks: dict[tuple[str, str], threading.Lock] = {}
        self._index_lock = threading.Lock()
        self._index_updated = False
        self._translate_lock = threading.Lock()  # one translation at a time (CPU-bound anyway)

    def _lock_for(self, pair: tuple[str, str]) -> threading.Lock:
        with self._guard:
            return self._pair_locks.setdefault(pair, threading.Lock())

    @staticmethod
    def _installed(a: str, b: str) -> bool:
        import argostranslate.translate as tr
        if hasattr(tr.get_installed_languages, "cache_clear"):
            tr.get_installed_languages.cache_clear()
        langs = {lang.code: lang for lang in tr.get_installed_languages()}
        return a in langs and b in langs and langs[a].get_translation(langs[b]) is not None

    def ensure_pairs(self, pairs: list[tuple[str, str]], status: StatusFn) -> None:
        for a, b in pairs:
            with self._lock_for((a, b)):
                if self._installed(a, b):
                    continue
                import argostranslate.package as pkg
                status(f"Downloading the {a}->{b} translation pack (first time only, 100-300 MB)...")
                with self._index_lock:
                    if not self._index_updated:
                        pkg.update_package_index()
                        self._index_updated = True
                match = next((p for p in pkg.get_available_packages()
                              if p.from_code == a and p.to_code == b), None)
                if match is None:
                    raise PackNotFoundError(f"No offline translation pack exists for {a} -> {b}.")
                pkg.install_from_path(match.download())

    def translate(self, text: str, a: str, b: str) -> str:
        import argostranslate.translate as tr
        with self._translate_lock:
            return tr.translate(text, a, b)


# =============================================================================
# Screen capture, clipboard image, temp PNG files
# =============================================================================

@dataclass
class Capture:
    image: PILImage
    left: int
    top: int


def grab_monitor_under_cursor() -> Capture:
    import mss
    from PIL import Image
    from pynput.mouse import Controller

    x, y = Controller().position
    with mss.mss() as sct:
        monitors = sct.monitors[1:] or sct.monitors
        mon = next((m for m in monitors
                    if m["left"] <= x < m["left"] + m["width"]
                    and m["top"] <= y < m["top"] + m["height"]), monitors[0])
        shot = sct.grab(mon)
        image = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
    return Capture(image, int(mon["left"]), int(mon["top"]))


def _windows_copy_dib(image: PILImage) -> None:
    from ctypes import wintypes

    buf = io.BytesIO()
    image.convert("RGB").save(buf, "BMP")
    data = buf.getvalue()[14:]  # strip BITMAPFILEHEADER -> CF_DIB

    k32 = ctypes.WinDLL("kernel32")  # private instances: don't clobber argtypes used by pynput
    u32 = ctypes.WinDLL("user32")
    k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k32.GlobalAlloc.restype = wintypes.HANDLE
    k32.GlobalLock.argtypes = [wintypes.HANDLE]
    k32.GlobalLock.restype = ctypes.c_void_p
    k32.GlobalUnlock.argtypes = [wintypes.HANDLE]
    k32.GlobalFree.argtypes = [wintypes.HANDLE]
    u32.OpenClipboard.argtypes = [wintypes.HWND]
    u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    u32.SetClipboardData.restype = wintypes.HANDLE

    handle = k32.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
    if not handle:
        raise OSError("GlobalAlloc failed")
    ctypes.memmove(k32.GlobalLock(handle), data, len(data))
    k32.GlobalUnlock(handle)
    for _ in range(10):
        if u32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        k32.GlobalFree(handle)
        raise OSError("Clipboard is busy")
    try:
        u32.EmptyClipboard()
        if not u32.SetClipboardData(8, handle):  # CF_DIB; clipboard owns handle on success
            k32.GlobalFree(handle)
            raise OSError("SetClipboardData failed")
    finally:
        u32.CloseClipboard()


def save_png(image: PILImage) -> Path:
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S") + f"_{int(time.time() * 1000) % 1000:03d}"
    path = TEMP_DIR / f"capture_{stamp}.png"
    image.save(path)
    return path


def reveal_file(path: Path) -> None:
    if sys.platform == "win32":
        subprocess.Popen(["explorer", "/select,", str(path)])
    elif shutil.which("xdg-open"):
        subprocess.Popen(["xdg-open", str(path.parent)])


def copy_image_to_clipboard(image: PILImage) -> str:
    """Copy image to the system clipboard; returns a status message."""
    if sys.platform == "win32":
        _windows_copy_dib(image)
        return "Image copied. Paste with Ctrl+V."
    buf = io.BytesIO()
    image.save(buf, "PNG")
    for cmd in (["wl-copy", "--type", "image/png"],
                ["xclip", "-selection", "clipboard", "-t", "image/png", "-i"]):
        if shutil.which(cmd[0]):
            subprocess.run(cmd, input=buf.getvalue(), check=True, timeout=5)
            return "Image copied. Paste with Ctrl+V."
    path = save_png(image)
    return f"No xclip/wl-copy found; saved instead: {path}"


# =============================================================================
# Background worker (serial job queue on a daemon thread)
# =============================================================================

class Worker:
    def __init__(self, name: str) -> None:
        self._jobs: queue.Queue[Callable[[], None]] = queue.Queue()
        threading.Thread(target=self._run, name=name, daemon=True).start()

    def submit(self, job: Callable[[], None]) -> None:
        self._jobs.put(job)

    def _run(self) -> None:
        while True:
            job = self._jobs.get()
            try:
                job()
            except Exception:
                traceback.print_exc()


# =============================================================================
# Background listener, single instance, autostart
# =============================================================================

def send_command(port: int, command: str) -> bool:
    """Send a one-word command to a local instance; False if nothing is listening."""
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=1) as conn:
            conn.sendall(command.encode("ascii"))
        return True
    except OSError:
        return False


def gui_python() -> str:
    """pythonw.exe on Windows (no console window), else the current interpreter."""
    exe = Path(sys.executable)
    if sys.platform == "win32":
        windowed = exe.with_name("pythonw.exe")
        return str(windowed if windowed.exists() else exe)
    return str(exe)


def launch_detached(*args: str) -> None:
    script = Path(__file__).resolve()
    log = open(LOG_PATH, "a", encoding="utf-8")
    kwargs: dict[str, Any] = {"stdin": subprocess.DEVNULL, "stdout": log, "stderr": subprocess.STDOUT,
                              "cwd": str(script.parent), "close_fds": True}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x00000008 | 0x00000200  # DETACHED_PROCESS | NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen([gui_python(), str(script), *args], **kwargs)
    log.close()


def make_hotkey_listener(callback: Callable[[], None]) -> Any:
    """Ctrl+Shift+HOTKEY_KEY matched by key code. pynput's GlobalHotKeys misses it on Windows,
    where Ctrl turns the 't' character into '\\x14'. The callback runs on pynput's thread."""
    from pynput import keyboard

    ctrl_keys = {keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r}
    shift_keys = {keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r}
    chars = {HOTKEY_KEY, HOTKEY_KEY.upper(), chr(ord(HOTKEY_KEY.upper()) - 64)}  # t, T, Ctrl+T
    pressed: set[str] = set()

    def name(key: Any) -> Optional[str]:
        if key in ctrl_keys:
            return "ctrl"
        if key in shift_keys:
            return "shift"
        vk, char = getattr(key, "vk", None), getattr(key, "char", None)
        if (sys.platform == "win32" and vk == ord(HOTKEY_KEY.upper())) or char in chars:
            return "key"
        return None

    def on_press(key: Any) -> None:
        part = name(key)
        if part == "key" and {"ctrl", "shift"} <= pressed and "key" not in pressed:
            callback()  # "key" not yet pressed: ignore auto-repeat while held
        if part:
            pressed.add(part)

    def on_release(key: Any) -> None:
        pressed.discard(name(key) or "")

    return keyboard.Listener(on_press=on_press, on_release=on_release)


def run_listener() -> int:
    """Tiny always-on process: owns the hotkey, starts the app or tells it to capture."""
    try:
        guard = socket.create_server(("127.0.0.1", LISTENER_PORT))
    except OSError:
        print("Listener is already running.")
        return 0
    last_launch = [0.0]

    def trigger() -> None:
        if send_command(APP_PORT, "capture"):
            return
        if time.monotonic() - last_launch[0] < LAUNCH_DEBOUNCE_S:
            return  # app is still starting up
        last_launch[0] = time.monotonic()
        launch_detached("--capture")

    hotkeys = make_hotkey_listener(lambda: threading.Thread(target=trigger, daemon=True).start())
    hotkeys.start()
    print(f"Listener running. {HOTKEY_LABEL} starts the app in capture mode.", flush=True)
    while True:
        conn, _ = guard.accept()
        with conn:
            try:
                data = conn.recv(64).strip()
            except OSError:
                continue
        if data == b"stop":
            hotkeys.stop()
            return 0


def _autostart_command() -> str:
    return " ".join(f'"{part}"' for part in (gui_python(), str(Path(__file__).resolve()), "--listener"))


def install_autostart() -> int:
    if sys.platform == "win32":
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WINDOWS_RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, AUTOSTART_NAME, 0, winreg.REG_SZ, _autostart_command())
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, WINDOWS_STARTUP_APPROVED_KEY, 0,
                                winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, AUTOSTART_NAME, 0, winreg.REG_BINARY, STARTUP_ENABLED)
    else:
        LINUX_AUTOSTART_FILE.parent.mkdir(parents=True, exist_ok=True)
        LINUX_AUTOSTART_FILE.write_text(
            "[Desktop Entry]\nType=Application\nName=Screen OCR Translator\n"
            f"Exec={_autostart_command()}\nX-GNOME-Autostart-enabled=true\n", encoding="utf-8")
    if not send_command(LISTENER_PORT, "ping"):
        launch_detached("--listener")
    print("Installed. The hotkey listener is running now and will start at every login.\n"
          "Press Ctrl+Shift+T anywhere to capture. Undo with --uninstall.")
    return 0


def uninstall_autostart() -> int:
    if sys.platform == "win32":
        import winreg
        for key_path in (WINDOWS_RUN_KEY, WINDOWS_STARTUP_APPROVED_KEY):
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                    winreg.KEY_SET_VALUE) as key:
                    winreg.DeleteValue(key, AUTOSTART_NAME)
            except FileNotFoundError:
                pass
    else:
        LINUX_AUTOSTART_FILE.unlink(missing_ok=True)
    stopped = send_command(LISTENER_PORT, "stop")
    print("Uninstalled: no longer starts at login." + (" Listener stopped." if stopped else ""))
    return 0


# =============================================================================
# Theme: Pillow-rendered shapes (anti-aliased rounded corners, gradients, shadows)
# =============================================================================

def hex_rgb(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)


def mix(a: str, b: str, t: float) -> str:
    """Blend color a toward b by t (0..1)."""
    ra, rb = hex_rgb(a), hex_rgb(b)
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(ra, rb))


def gradient_image(w: int, h: int, colors: Sequence[str]) -> PILImage:
    from PIL import Image
    w, h = max(w, 1), max(h, 1)
    stops = [hex_rgb(c) for c in colors]
    if len(stops) == 1:
        return Image.new("RGB", (w, h), stops[0])
    row = Image.new("RGB", (w, 1))
    px = row.load()
    segments = len(stops) - 1
    for x in range(w):
        t = x / max(w - 1, 1) * segments
        i = min(int(t), segments - 1)
        f = t - i
        a, b = stops[i], stops[i + 1]
        px[x, 0] = (round(a[0] + (b[0] - a[0]) * f), round(a[1] + (b[1] - a[1]) * f),
                    round(a[2] + (b[2] - a[2]) * f))
    return row.resize((w, h))


@functools.lru_cache(maxsize=64)
def _corner_mask(radius: int) -> PILImage:
    """Anti-aliased top-left quarter circle, rendered once per radius."""
    from PIL import Image, ImageDraw
    s = 4
    big = Image.new("L", (radius * s, radius * s), 0)
    ImageDraw.Draw(big).ellipse((0, 0, 2 * radius * s - 1, 2 * radius * s - 1), fill=255)
    return big.resize((radius, radius), Image.Resampling.LANCZOS)


def rounded_mask(w: int, h: int, radius: int) -> PILImage:
    """Rounded-rectangle alpha mask built from cached corners (cheap at any size)."""
    from PIL import Image
    w, h = max(w, 1), max(h, 1)
    r = max(0, min(radius, w // 2, h // 2))
    mask = Image.new("L", (w, h), 255)
    if r:
        corner = _corner_mask(r)
        mask.paste(corner, (0, 0))
        mask.paste(corner.transpose(Image.Transpose.FLIP_LEFT_RIGHT), (w - r, 0))
        mask.paste(corner.transpose(Image.Transpose.FLIP_TOP_BOTTOM), (0, h - r))
        mask.paste(corner.transpose(Image.Transpose.ROTATE_180), (w - r, h - r))
    return mask


def nine_slice(template: PILImage, border: int, w: int, h: int,
               vertical: bool = True) -> PILImage:
    """Stretch a small template to w x h, keeping its `border`-px edges and corners intact.
    vertical=False stretches only horizontally (template height must equal h)."""
    from PIL import Image
    tw, th, b = template.width, template.height, border
    w = max(w, 2 * b + 1)
    h = max(h, 2 * b + 1) if vertical else th
    out = Image.new(template.mode, (w, h))
    xs = [(0, b, 0, b), (b, tw - b, b, w - b), (tw - b, tw, w - b, w)]  # src x0,x1 -> dst x0,x1
    ys = ([(0, b, 0, b), (b, th - b, b, h - b), (th - b, th, h - b, h)] if vertical
          else [(0, th, 0, th)])
    for sx0, sx1, dx0, dx1 in xs:
        for sy0, sy1, dy0, dy1 in ys:
            piece = template.crop((sx0, sy0, sx1, sy1))
            if piece.size != (dx1 - dx0, dy1 - dy0):  # edges are flat color: nearest is exact
                piece = piece.resize((dx1 - dx0, dy1 - dy0), Image.Resampling.NEAREST)
            out.paste(piece, (dx0, dy0))
    return out


def paint(base: PILImage, fill: str | Sequence[str], box: tuple[int, int, int, int],
          radius: int) -> None:
    """Paste a rounded rectangle (solid or gradient) onto base."""
    from PIL import Image
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    if w <= 0 or h <= 0:
        return
    src = Image.new("RGB", (w, h), fill) if isinstance(fill, str) else gradient_image(w, h, fill)
    base.paste(src, (x0, y0), rounded_mask(w, h, max(0, min(radius, w // 2, h // 2))))


def soft_shadow(base: PILImage, box: tuple[int, int, int, int], radius: int, blur: float,
                strength: int = 120, color: str = "#000000") -> None:
    from PIL import Image, ImageDraw, ImageFilter
    mask = Image.new("L", base.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle(box, radius=radius, fill=strength)
    base.paste(Image.new("RGB", base.size, color), (0, 0), mask.filter(ImageFilter.GaussianBlur(blur)))


CARD_SLICE = CARD_MARGIN + CARD_RADIUS + 6  # covers margin, corner and shadow offset


@functools.lru_cache(maxsize=4)
def _card_template(bg: str) -> PILImage:
    """Small fully-drawn card (shadow + border + fill); nine_slice() stretches it."""
    from PIL import Image
    m, r = CARD_MARGIN, CARD_RADIUS
    size = 2 * CARD_SLICE + 2
    base = Image.new("RGB", (size, size), bg)
    soft_shadow(base, (m, m + 4, size - m, size - m + 3), r, m / 2)
    paint(base, mix(THEME["card_border"], "#FFFFFF", 0.03), (m, m, size - m, size - m), r)
    paint(base, THEME["card"], (m + 1, m + 1, size - m - 1, size - m - 1), r - 1)
    return base


def card_image(w: int, h: int, bg: str) -> PILImage:
    return nine_slice(_card_template(bg), CARD_SLICE, w, h)


def pill_image(w: int, h: int, fill: str | Sequence[str], bg: str,
               border: Optional[str] = None, radius: Optional[int] = None) -> PILImage:
    from PIL import Image
    r = h // 2 if radius is None else radius
    if isinstance(fill, str) and w > 2 * r + 4:  # solid shapes: draw small, stretch the middle
        template = pill_image(2 * r + 4, h, fill, bg, border, r)
        return nine_slice(template, r + 1, w, h, vertical=h > 2 * r + 2)
    base = Image.new("RGB", (w, h), bg)
    if border:
        paint(base, border, (0, 0, w, h), r)
        paint(base, fill, (1, 1, w - 1, h - 1), max(0, r - 1))
    else:
        paint(base, fill, (0, 0, w, h), r)
    return base


def pill_rgba(w: int, h: int, fill: str | Sequence[str], radius: Optional[int] = None) -> PILImage:
    """Rounded shape with transparent corners (for drawing over arbitrary pixels)."""
    from PIL import Image
    src = Image.new("RGB", (w, h), fill) if isinstance(fill, str) else gradient_image(w, h, fill)
    src = src.convert("RGBA")
    src.putalpha(rounded_mask(w, h, h // 2 if radius is None else radius))
    return src


def icon_image(size: int, colors: Sequence[str]) -> PILImage:
    from PIL import Image, ImageDraw
    s = 4
    big = Image.new("L", (size * s, size * s), 0)
    ImageDraw.Draw(big).ellipse((0, 0, size * s - 1, size * s - 1), fill=255)
    img = gradient_image(size, size, colors).convert("RGBA")
    img.putalpha(big.resize((size, size), Image.Resampling.LANCZOS))
    return img


def set_dark_title_bar(window: tk.Misc) -> None:
    if sys.platform != "win32":
        return
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value),  # IMMERSIVE_DARK_MODE
                                                   ctypes.sizeof(value))
        caption = ctypes.c_int(int.from_bytes(bytes(reversed(hex_rgb(THEME["bg"]))), "big"))
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(caption),  # CAPTION_COLOR
                                                   ctypes.sizeof(caption))
    except Exception:
        pass


def apply_theme(root: tk.Tk) -> None:
    t = THEME
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=t["bg"], foreground=t["text"], font=(FONT_FAMILY, 10))
    style.configure("Field.TCombobox", fieldbackground=t["field"], background=t["field"],
                    foreground=t["text"], arrowcolor=t["link"], bordercolor=t["field"],
                    lightcolor=t["field"], darkcolor=t["field"], padding=(2, 3), arrowsize=15,
                    relief="flat")
    style.map("Field.TCombobox", fieldbackground=[("readonly", t["field"])],
              background=[("readonly", t["field"]), ("active", t["field"])],
              foreground=[("readonly", t["text"])], bordercolor=[("focus", t["field"])],
              lightcolor=[("focus", t["field"])], darkcolor=[("focus", t["field"])],
              arrowcolor=[("active", "#FFFFFF")],
              selectbackground=[("readonly", t["field"])], selectforeground=[("readonly", t["text"])])
    root.option_add("*TCombobox*Listbox.background", t["field"])
    root.option_add("*TCombobox*Listbox.foreground", t["text"])
    root.option_add("*TCombobox*Listbox.selectBackground", GRADIENT[0])
    root.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")
    root.option_add("*TCombobox*Listbox.font", (FONT_FAMILY, 10))
    root.option_add("*TCombobox*Listbox.borderWidth", 0)
    root.option_add("*TCombobox*Listbox.relief", "flat")
    for orient, sticky in (("Vertical", "ns"), ("Horizontal", "ew")):
        name = f"Slim.{orient}.TScrollbar"
        style.layout(name, [(f"{orient}.Scrollbar.trough", {"sticky": sticky, "children": [
            (f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        style.configure(name, troughcolor=t["card"], background=t["thumb"], bordercolor=t["card"],
                        lightcolor=t["thumb"], darkcolor=t["thumb"], arrowsize=7, gripcount=0,
                        relief="flat")
        style.map(name, background=[("active", t["thumb_hover"])],
                  lightcolor=[("active", t["thumb_hover"])], darkcolor=[("active", t["thumb_hover"])])


def round_rect(canvas: tk.Canvas, x0: float, y0: float, x1: float, y1: float, r: float,
               **kwargs: Any) -> int:
    points = [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
              x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]
    return canvas.create_polygon(points, smooth=True, **kwargs)


# =============================================================================
# Custom widgets (Canvas + Pillow images)
# =============================================================================

PILL_VARIANTS: dict[str, dict[str, Any]] = {
    "primary": {"fill": GRADIENT[:2], "text": "#FFFFFF",
                "hover": (mix(GRADIENT[0], "#FFFFFF", 0.18), mix(GRADIENT[1], "#FFFFFF", 0.18))},
    "light": {"fill": "#FFFFFF", "hover": "#F3EEFF", "text": "#3B1D8F"},
    "soft": {"fill": THEME["soft"], "hover": THEME["soft_hover"], "text": THEME["text"],
             "border": THEME["field_border"]},
    "ghost": {"fill": None, "hover": THEME["soft"], "text": THEME["link"]},
    "danger": {"fill": THEME["soft"], "hover": THEME["err_soft"], "text": THEME["err"],
               "border": THEME["field_border"]},
    "badge": {"fill": "#2D1A3D", "hover": None, "text": "#F9A8D4"},
}


class PillButton(tk.Canvas):
    def __init__(self, parent: tk.Misc, text: str, command: Optional[Callable[[], None]] = None,
                 variant: str = "soft", bg: str = THEME["card"], size: int = 10, height: int = 34,
                 padx: int = 18, min_text: str = "") -> None:
        from tkinter import font as tkfont
        super().__init__(parent, height=height, bg=bg, highlightthickness=0, bd=0,
                         cursor="hand2" if command else "arrow")
        self.command = command
        self.variant = PILL_VARIANTS[variant]
        self.text = text
        self._bg = bg
        self._height = height
        self._padx = padx
        self._min_text = min_text
        self._font = (FONT_FAMILY, size, "bold")
        self._measure = tkfont.Font(family=FONT_FAMILY, size=size, weight="bold").measure
        self._width = 0
        self._hover = self._pressed = False
        self.enabled = True
        self._cache: dict[tuple, Any] = {}
        self._img = self.create_image(0, 0, anchor="nw")
        self._label = self.create_text(0, height / 2, text=text, font=self._font)
        self.bind("<Enter>", lambda _e: self._set_hover(True))
        self.bind("<Leave>", lambda _e: self._set_hover(False))
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.set_text(text)

    def _render(self) -> None:
        from PIL import ImageTk
        v = self.variant
        fill = v["hover"] if (self._hover and self.enabled and v.get("hover")) else v["fill"]
        fill = fill or self._bg
        key = (self._width, str(fill), self._bg)
        if key not in self._cache:
            self._cache[key] = ImageTk.PhotoImage(
                pill_image(self._width, self._height, fill, self._bg, v.get("border")))
        self.itemconfigure(self._img, image=self._cache[key])
        self.itemconfigure(self._label, fill=v["text"] if self.enabled else THEME["faint"])
        self.coords(self._label, self._width / 2, self._height / 2 + (1 if self._pressed else 0))

    def set_text(self, text: str) -> None:
        self.text = text
        widest = max(self._measure(text), self._measure(self._min_text))
        self._width = widest + 2 * self._padx
        self.configure(width=self._width)
        self.itemconfigure(self._label, text=text)
        self._render()

    def set_bg(self, bg: str) -> None:
        if bg != self._bg:
            self._bg = bg
            self.configure(bg=bg)
            self._render()

    def set_enabled(self, enabled: bool) -> None:
        self.enabled = enabled
        self.configure(cursor="hand2" if enabled and self.command else "arrow")
        self._render()

    def _set_hover(self, hover: bool) -> None:
        self._hover = hover
        if not hover:
            self._pressed = False
        self._render()

    def _on_press(self, _event: Any) -> None:
        if self.enabled and self.command:
            self._pressed = True
            self._render()

    def _on_release(self, _event: Any) -> None:
        fire = self._pressed and self._hover
        self._pressed = False
        self._render()
        if fire and self.enabled and self.command:
            self.command()


class Toggle(tk.Canvas):
    """Animated on/off switch bound to a BooleanVar."""

    W, H, KNOB = 46, 26, 20

    def __init__(self, parent: tk.Misc, variable: tk.BooleanVar, command: Callable[[], None],
                 bg: str = THEME["card"]) -> None:
        from PIL import ImageTk
        super().__init__(parent, width=self.W, height=self.H, bg=bg, highlightthickness=0, bd=0,
                         cursor="hand2")
        self.variable = variable
        self.command = command
        self._on = ImageTk.PhotoImage(pill_image(self.W, self.H, GRADIENT[:2], bg))
        self._off = ImageTk.PhotoImage(pill_image(self.W, self.H, THEME["soft_hover"], bg))
        self._knob_img = ImageTk.PhotoImage(icon_image(self.KNOB, ["#FFFFFF"]))
        self._track = self.create_image(0, 0, anchor="nw")
        self._knob = self.create_image(0, self.H / 2, image=self._knob_img)
        self._x = self._target()
        self._draw()
        self.bind("<Button-1>", lambda _e: self.toggle())

    def _target(self) -> float:
        return self.W - self.H / 2 if self.variable.get() else self.H / 2

    def _draw(self) -> None:
        self.itemconfigure(self._track, image=self._on if self.variable.get() else self._off)
        self.coords(self._knob, self._x, self.H / 2)

    def toggle(self) -> None:
        self.variable.set(not self.variable.get())
        self._animate()
        self.command()

    def _animate(self) -> None:
        target = self._target()
        self._x += (target - self._x) * 0.45
        if abs(target - self._x) < 0.6:
            self._x = target
        self._draw()
        if self._x != target:
            self.after(14, self._animate)


class Field(tk.Canvas):
    """Rounded input shell around a borderless ttk.Combobox."""

    H = 40

    def __init__(self, parent: tk.Misc, variable: tk.StringVar, values: Sequence[str],
                 command: Callable[[], None], width: int = 190, bg: str = THEME["card"]) -> None:
        from PIL import ImageTk
        super().__init__(parent, width=width, height=self.H, bg=bg, highlightthickness=0, bd=0)
        self._img = ImageTk.PhotoImage(pill_image(width, self.H, THEME["field"], bg,
                                                  THEME["field_border"], radius=12))
        self.create_image(0, 0, image=self._img, anchor="nw")
        self.combo = ttk.Combobox(self, textvariable=variable, values=list(values), state="readonly",
                                  style="Field.TCombobox", font=(FONT_FAMILY, 11))
        self.create_window(12, self.H / 2, anchor="w", window=self.combo, width=width - 20)

        def on_select(_event: Any) -> None:
            self.combo.selection_clear()
            command()

        self.combo.bind("<<ComboboxSelected>>", on_select)


class Card(tk.Canvas):
    """Rounded, shadowed panel. Put content in .body, header buttons in .actions."""

    def __init__(self, parent: tk.Misc, title: Optional[str] = None, accent: str = "settings",
                 glyph: str = "", height: int = 200) -> None:
        super().__init__(parent, bg=THEME["bg"], highlightthickness=0, bd=0, height=height)
        c = THEME["card"]
        self.header = tk.Frame(self, bg=c)
        self.body = tk.Frame(self, bg=c)
        self.body.columnconfigure(0, weight=1)
        self.body.rowconfigure(0, weight=1)
        self.actions = tk.Frame(self.header, bg=c)
        self.actions.pack(side="right")
        self._has_header = title is not None
        if title is not None:
            from PIL import ImageTk
            icon = tk.Canvas(self.header, width=28, height=28, bg=c, highlightthickness=0, bd=0)
            self._icon = ImageTk.PhotoImage(icon_image(28, CARD_ACCENTS[accent]))
            icon.create_image(14, 14, image=self._icon)
            icon.create_text(14, 14, text=glyph, fill="#FFFFFF", font=(FONT_FAMILY, 10, "bold"))
            icon.pack(side="left")
            tk.Label(self.header, text=title, bg=c, fg=THEME["text"],
                     font=(FONT_FAMILY, 11, "bold")).pack(side="left", padx=(10, 0))
        self._bg_item = self.create_image(0, 0, anchor="nw")
        self._header_win = self.create_window(0, 0, anchor="nw", window=self.header)
        self._body_win = self.create_window(0, 0, anchor="nw", window=self.body)
        if not self._has_header:
            self.itemconfigure(self._header_win, state="hidden")
        self._photo: Any = None
        self._pending = False
        self.bind("<Configure>", lambda _e: self._schedule())

    def _schedule(self) -> None:
        if not self._pending:
            self._pending = True
            self.after(25, self._render)

    def _render(self) -> None:
        from PIL import ImageTk
        self._pending = False
        w, h = self.winfo_width(), self.winfo_height()
        if w < 60 or h < 40:
            return
        self._photo = ImageTk.PhotoImage(card_image(w, h, THEME["bg"]))
        self.itemconfigure(self._bg_item, image=self._photo)
        pad = CARD_MARGIN + 18
        top = CARD_MARGIN + 14
        if self._has_header:
            self.coords(self._header_win, pad, top)
            self.itemconfigure(self._header_win, width=w - 2 * pad)
            top += self.header.winfo_reqheight() + 10
        self.coords(self._body_win, pad, top)
        self.itemconfigure(self._body_win, width=w - 2 * pad,
                           height=max(10, h - top - CARD_MARGIN - 16))

    def fit_height(self) -> None:
        self.update_idletasks()
        extra = (self.header.winfo_reqheight() + 10) if self._has_header else 0
        self.configure(height=self.body.winfo_reqheight() + extra + 2 * CARD_MARGIN + 30)


class Banner(tk.Canvas):
    """Gradient hero header with title, hotkey chip and the New capture button."""

    H = 112

    def __init__(self, parent: tk.Misc, on_capture: Callable[[], None]) -> None:
        super().__init__(parent, height=self.H, bg=THEME["bg"], highlightthickness=0, bd=0)
        self._bg_item = self.create_image(0, 0, anchor="nw")
        self._title = self.create_text(0, 0, anchor="w", fill="#FFFFFF",
                                       text="Screen OCR Translator", font=(FONT_FAMILY, 21, "bold"))
        self._subtitle = self.create_text(0, 0, anchor="w", fill="#FDF4FF",
                                          text="Capture  •  Translate  •  Summarize  •  100% offline",
                                          font=(FONT_FAMILY, 10))
        self._chip_text = self.create_text(0, 0, text="Ctrl + Shift + T", fill="#FFFFFF",
                                           font=(FONT_FAMILY, 9, "bold"))
        self.button = PillButton(self, "✦  New capture", on_capture, variant="light",
                                 bg=GRADIENT[2], size=10, height=40, padx=20)
        self._button_win = self.create_window(0, 0, anchor="e", window=self.button)
        self._photo: Any = None
        self._pending = False
        self.bind("<Configure>", lambda _e: self._schedule())

    def _schedule(self) -> None:
        if not self._pending:
            self._pending = True
            self.after(25, self._render)

    def _render(self) -> None:
        from PIL import Image, ImageDraw, ImageFilter, ImageTk
        self._pending = False
        w, h, m = self.winfo_width(), self.H, CARD_MARGIN
        if w < 200:
            return
        base = Image.new("RGB", (w, h), THEME["bg"])
        soft_shadow(base, (m + 10, m + 10, w - m - 10, h - m + 6), 22, 10, 150, GRADIENT[0])
        bw, bh = w - 2 * m, h - 2 * m
        art = gradient_image(bw, bh, GRADIENT)
        glow = Image.new("L", (bw, bh), 0)  # soft light blobs for an "aurora" feel
        d = ImageDraw.Draw(glow)
        d.ellipse((-bw * 0.05, -bh * 1.2, bw * 0.45, bh * 0.9), fill=110)
        d.ellipse((bw * 0.55, bh * 0.2, bw * 1.05, bh * 2.0), fill=80)
        art.paste(Image.new("RGB", (bw, bh), "#FFFFFF"), (0, 0), glow.filter(ImageFilter.GaussianBlur(bh * 0.6)))
        shade = Image.new("L", (bw, bh), 0)
        ImageDraw.Draw(shade).rectangle((0, bh * 0.55, bw, bh), fill=60)
        art.paste(Image.new("RGB", (bw, bh), "#1A0B3D"), (0, 0), shade.filter(ImageFilter.GaussianBlur(bh * 0.4)))
        base.paste(art, (m, m), rounded_mask(bw, bh, 22))

        button_x = w - m - 22
        button_w = int(self.button.cget("width"))
        sample = base.getpixel((min(w - 1, button_x - button_w // 2), h // 2))
        self.button.set_bg("#%02x%02x%02x" % sample)
        cx0, _, cx1, _ = self.bbox(self._chip_text)
        chip_w, chip_h = cx1 - cx0 + 30, 32
        chip_x1 = button_x - button_w - 16
        chip_box = (chip_x1 - chip_w, h // 2 - chip_h // 2, chip_x1, h // 2 + chip_h // 2)
        chip = base.crop(chip_box)
        chip = Image.blend(chip, Image.new("RGB", chip.size, "#FFFFFF"), 0.22)
        base.paste(chip, chip_box[:2], rounded_mask(chip_w, chip_h, chip_h // 2))

        self._photo = ImageTk.PhotoImage(base)
        self.itemconfigure(self._bg_item, image=self._photo)
        _, ty0, _, ty1 = self.bbox(self._title)
        _, sy0, _, sy1 = self.bbox(self._subtitle)
        block = (ty1 - ty0) + (sy1 - sy0) + 2
        top = h / 2 - block / 2
        self.coords(self._title, m + 30, top + (ty1 - ty0) / 2)
        self.coords(self._subtitle, m + 32, top + (ty1 - ty0) + 2 + (sy1 - sy0) / 2)
        self.coords(self._chip_text, chip_x1 - chip_w / 2, h / 2)
        self.coords(self._button_win, button_x, h / 2)


class Shimmer(tk.Canvas):
    """Thin animated gradient bar shown while work is in progress."""

    H, SEGMENT = 3, 300

    def __init__(self, parent: tk.Misc) -> None:
        from PIL import ImageTk
        super().__init__(parent, height=self.H, bg=THEME["bg"], highlightthickness=0, bd=0)
        seg = gradient_image(self.SEGMENT, self.H, GRADIENT).convert("RGBA")
        alpha = seg.getchannel("A")
        for x in range(self.SEGMENT):  # fade both ends
            a = int(255 * (1 - abs(x / (self.SEGMENT - 1) * 2 - 1)) ** 0.7)
            for y in range(self.H):
                alpha.putpixel((x, y), a)
        seg.putalpha(alpha)
        self._seg = ImageTk.PhotoImage(seg)
        self._item = self.create_image(-self.SEGMENT, 0, image=self._seg, anchor="nw", state="hidden")
        self._x = -self.SEGMENT
        self._running = False

    def start(self) -> None:
        if not self._running:
            self._running = True
            self.itemconfigure(self._item, state="normal")
            self._tick()

    def stop(self) -> None:
        self._running = False
        self.itemconfigure(self._item, state="hidden")

    def _tick(self) -> None:
        if not self._running:
            return
        self._x += 9
        if self._x > self.winfo_width():
            self._x = -self.SEGMENT
        self.coords(self._item, self._x, 0)
        self.after(16, self._tick)


class PulseDot(tk.Canvas):
    """Status light; gently pulses while busy."""

    S = 18

    def __init__(self, parent: tk.Misc, bg: str = THEME["bg"]) -> None:
        super().__init__(parent, width=self.S, height=self.S, bg=bg, highlightthickness=0, bd=0)
        self._bg = bg
        self._halo = self.create_oval(0, 0, 0, 0, outline="")
        self._dot = self.create_oval(0, 0, 0, 0, outline="")
        self._color = THEME["muted"]
        self._phase = 0.0
        self._busy = False
        self._draw(4.0, 0.0)

    def set(self, kind: str) -> None:
        self._color = {"busy": THEME["busy"], "ok": THEME["ok"], "err": THEME["err"]}.get(
            kind, THEME["muted"])
        was_busy, self._busy = self._busy, kind == "busy"
        if self._busy and not was_busy:
            self._pulse()
        elif not self._busy:
            self._draw(4.0, 3.0 if kind in ("ok", "err") else 0.0)

    def _draw(self, radius: float, halo: float) -> None:
        c = self.S / 2
        self.itemconfigure(self._dot, fill=self._color)
        self.itemconfigure(self._halo, fill=mix(self._color, self._bg, 0.7))
        self.coords(self._dot, c - radius, c - radius, c + radius, c + radius)
        self.coords(self._halo, c - radius - halo, c - radius - halo, c + radius + halo, c + radius + halo)

    def _pulse(self) -> None:
        if not self._busy:
            return
        import math
        self._phase += 0.22
        wave = (math.sin(self._phase) + 1) / 2
        self._draw(3.5 + wave, 1.0 + 3.0 * wave)
        self.after(40, self._pulse)


# =============================================================================
# GUI
# =============================================================================

class App:
    def __init__(self, root: tk.Tk, server: Optional[socket.socket] = None) -> None:
        self.root = root
        self.ui_queue: queue.Queue[Callable[[], None]] = queue.Queue()
        self.ocr = OcrEngine()
        self.translator = ArgosTranslator()
        self.heavy = Worker("ocr-translate")
        self.light = Worker("summary")
        self.capture_busy = threading.Event()
        self._capture_lock = threading.Lock()
        self.server = server
        self.screenshot: Optional[PILImage] = None
        self._preview: Optional[PILImage] = None  # small copy of screenshot for the Image card
        self.table_tsv = ""
        self.ocr_job = 0
        self.translate_job = 0
        self.original_placeholder = False  # Original box shows a hint, not OCR text
        self.detected_source: Optional[str] = None  # language found by Auto-detect
        self.thumbnail_ref: Any = None
        self.hotkey: Any = None
        self._flash_text: dict[int, str] = {}
        apply_theme(root)
        self._build_ui()
        if server is not None:
            threading.Thread(target=self._serve, args=(server,), daemon=True).start()
        self._prewarm()
        root.after(UI_POLL_MS, self._poll)

    def _prewarm(self) -> None:
        """Load the OCR model and translation engine (~15-25 s on a laptop CPU) in the
        background at startup, so they're ready by the time the user finishes selecting."""
        def warm_translation() -> None:
            try:
                import argostranslate.translate  # noqa: F401  (heavy import, cached afterwards)
            except Exception:
                traceback.print_exc()

        threading.Thread(target=warm_translation, name="prewarm-translate", daemon=True).start()
        self.heavy.submit(lambda: self.ocr._reader("latin"))  # first in the OCR queue

    # ---- UI construction -----------------------------------------------------

    def _build_ui(self) -> None:
        t, r = THEME, self.root
        r.title("Screen OCR Translator")
        r.geometry("1180x940")
        r.minsize(940, 780)
        r.configure(bg=t["bg"])
        r.protocol("WM_DELETE_WINDOW", r.withdraw)
        r.columnconfigure(0, weight=1, uniform="col")
        r.columnconfigure(1, weight=1, uniform="col")

        Banner(r, self.new_capture).grid(row=0, column=0, columnspan=2, sticky="ew",
                                         padx=12, pady=(10, 0))
        self.shimmer = Shimmer(r)
        self.shimmer.grid(row=1, column=0, columnspan=2, sticky="ew", padx=40, pady=(2, 2))

        controls = Card(r, height=100)
        controls.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12)
        self._build_controls(controls.body)
        controls.fit_height()

        # Drag the gaps between boxes to resize them (vertical: rows, horizontal: columns).
        rows = self._panes(r, "vertical")
        rows.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=12)
        r.rowconfigure(3, weight=1)
        top, bottom = self._panes(rows, "horizontal"), self._panes(rows, "horizontal")

        card, self.original_text = self._text_card(top, "Original text", "original", "Aa", False)
        top.add(card, minsize=220, stretch="always")
        self.original_text.bind("<Key>", self._clear_original_placeholder, add="+")
        card, self.translation_text = self._text_card(top, "Translation", "translation", "文", False)
        top.add(card, minsize=220, stretch="always")

        summary, self.summary_text = self._text_card(rows, "Summary", "summary", "≡", True)
        self.method_badge = PillButton(summary.header, "", variant="badge", size=8, height=22,
                                       padx=10)

        card, self.table_text = self._text_card(bottom, "Table for Excel", "table", "XL", True,
                                                copy=False, mono=True, wrap="none")
        self.table_copy_btn = PillButton(card.actions, "Copy for Excel", self._copy_table,
                                         variant="ghost", size=9, height=28, padx=12,
                                         min_text="Copied ✓")
        self.table_copy_btn.pack(side="right")
        bottom.add(card, minsize=220, stretch="always")

        card = Card(bottom, "Image", "image", "◆")
        bottom.add(card, minsize=220, stretch="always")
        rows.add(top, minsize=MIN_PANE_PX, height=250, stretch="always")
        rows.add(summary, minsize=MIN_PANE_PX, height=190, stretch="always")
        rows.add(bottom, minsize=MIN_PANE_PX, height=210, stretch="always")
        self.save_btn = PillButton(card.actions, "Save PNG", self._save_image, variant="ghost",
                                   size=9, height=28, padx=12)
        self.save_btn.pack(side="right")
        self.copy_image_btn = PillButton(card.actions, "Copy image", self._copy_image,
                                         variant="ghost", size=9, height=28, padx=12,
                                         min_text="Copied ✓")
        self.copy_image_btn.pack(side="right", padx=(0, 2))
        self.thumb_canvas = tk.Canvas(card.body, bg=t["card"], highlightthickness=0, bd=0,
                                      width=1, height=1)
        self.thumb_canvas.grid(row=0, column=0, sticky="nsew")
        self._thumb_bg = self.thumb_canvas.create_image(0, 0, anchor="nw")
        self._thumb_hint = self.thumb_canvas.create_text(
            0, 0, text="No capture yet\nPress Ctrl+Shift+T", fill=t["faint"], justify="center",
            font=(FONT_FAMILY, 10))
        self.thumb_canvas.bind("<Configure>", lambda _e: self._schedule_thumbnail())
        self._thumbnail_pending = False

        bar = tk.Frame(r, bg=t["bg"])
        bar.grid(row=4, column=0, columnspan=2, sticky="ew", padx=26, pady=(2, 12))
        self.status_dot = PulseDot(bar)
        self.status_dot.pack(side="left", padx=(0, 8))
        self.status_var = tk.StringVar()
        tk.Label(bar, textvariable=self.status_var, bg=t["bg"], fg=t["muted"],
                 font=(FONT_FAMILY, 10)).pack(side="left")
        PillButton(bar, "Quit", self.quit, variant="danger", bg=t["bg"], size=9, height=32,
                   padx=20).pack(side="right")
        tk.Label(bar, text="Drag the gaps between boxes to resize   ·   Close = hide   ·   "
                           "Quit = stop", bg=t["bg"], fg=t["faint"],
                 font=(FONT_FAMILY, 9)).pack(side="right", padx=14)

        self._set_status("Ready. Press Ctrl+Shift+T to capture.", "idle")
        self._set_summary("", "")
        self._set_table(None)

    def _build_controls(self, body: tk.Frame) -> None:
        t = THEME
        self.source_var = tk.StringVar(value=DEFAULT_SOURCE)
        self.target_var = tk.StringVar(value=DEFAULT_TARGET)
        self.join_var = tk.BooleanVar(value=True)

        def labeled(parent: tk.Frame, label: str, var: tk.StringVar, values: Sequence[str],
                    command: Callable[[], None]) -> tuple[tk.Frame, tk.Label]:
            frame = tk.Frame(parent, bg=t["card"])
            caption = tk.Label(frame, text=label, bg=t["card"], fg=t["muted"],
                               font=(FONT_FAMILY, 8, "bold"))
            caption.pack(anchor="w", padx=2)
            Field(frame, var, values, command).pack(anchor="w", pady=(5, 0))
            return frame, caption

        frame, self.source_caption = labeled(body, "TEXT IN IMAGE", self.source_var,
                                             [AUTO_DETECT, *LANGUAGES], self.request_ocr)
        frame.pack(side="left")
        tk.Label(body, text="➜", bg=t["card"], fg=t["link"],
                 font=(FONT_FAMILY, 16)).pack(side="left", padx=8, pady=(20, 0))
        labeled(body, "TRANSLATE TO", self.target_var, list(LANGUAGES),
                self.request_translation)[0].pack(side="left")
        PillButton(body, "↻  Translate again", self.request_translation, variant="soft",
                   size=10, height=40, padx=16).pack(side="right", pady=(20, 0))
        toggle_row = tk.Frame(body, bg=t["card"])
        toggle_row.pack(side="left", padx=(20, 0), pady=(20, 0))
        toggle = Toggle(toggle_row, self.join_var, self.request_translation)
        toggle.pack(side="left")
        label = tk.Label(toggle_row, text="Join broken lines before translating", bg=t["card"],
                         fg=t["text"], font=(FONT_FAMILY, 10), cursor="hand2")
        label.pack(side="left", padx=(10, 0))
        label.bind("<Button-1>", lambda _e: toggle.toggle())

    @staticmethod
    def _panes(parent: tk.Misc, orient: str) -> tk.PanedWindow:
        """Container whose children can be resized by dragging the gap between them.
        Live (opaque) resizing re-lays-out every box per mouse move and lags in Tk, so a
        glowing guide line follows the mouse and the boxes resize once on release."""
        cursor = "sb_v_double_arrow" if orient == "vertical" else "sb_h_double_arrow"
        panes = tk.PanedWindow(parent, orient=orient, bg=THEME["bg"], bd=0, sashwidth=8,
                               sashrelief="flat", sashpad=0, opaqueresize=False, showhandle=False,
                               sashcursor=cursor)
        try:
            panes.configure(proxybackground=GRADIENT[1], proxyborderwidth=0, proxyrelief="flat")
        except tk.TclError:  # Tk older than 8.6.5: default grey guide line
            pass
        return panes

    def _text_card(self, parent: tk.Misc, title: str, accent: str, glyph: str, readonly: bool,
                   copy: bool = True, mono: bool = False,
                   wrap: str = "word") -> tuple[Card, tk.Text]:
        t = THEME
        card = Card(parent, title, accent, glyph)
        font = (MONO_FAMILY, 10) if mono else (FONT_FAMILY, 11)
        text = tk.Text(card.body, wrap=wrap, undo=not readonly, font=font, width=1, height=1,
                       bg=t["card"], fg=t["text"], insertbackground=t["link"],
                       selectbackground=t["select"], selectforeground="#FFFFFF",
                       relief="flat", bd=0, highlightthickness=0, padx=2, pady=2,
                       spacing1=2, spacing3=4)
        text.tag_configure("muted", foreground=t["muted"])
        scroll = ttk.Scrollbar(card.body, orient="vertical", command=text.yview,
                               style="Slim.Vertical.TScrollbar")
        text.configure(yscrollcommand=self._auto_hide(scroll))
        text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns", padx=(10, 0))
        if wrap == "none":
            hscroll = ttk.Scrollbar(card.body, orient="horizontal", command=text.xview,
                                    style="Slim.Horizontal.TScrollbar")
            text.configure(xscrollcommand=self._auto_hide(hscroll))
            hscroll.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        if readonly:
            text.configure(state="disabled", cursor="arrow")
        if copy:
            button = PillButton(card.actions, "Copy", variant="ghost", size=9, height=28, padx=12,
                                min_text="Copied ✓")
            button.command = lambda: self._copy_text(text, button)
            button.set_enabled(True)
            button.pack(side="right")
        return card, text

    @staticmethod
    def _auto_hide(scrollbar: ttk.Scrollbar) -> Callable[[str, str], None]:
        """Scroll command that shows the scrollbar only when the content overflows."""
        def update(first: str, last: str) -> None:
            if float(first) <= 0.0 and float(last) >= 1.0:
                scrollbar.grid_remove()
            else:
                scrollbar.grid()
            scrollbar.set(first, last)
        return update

    # ---- thread-safe UI plumbing -------------------------------------------

    def post(self, fn: Callable[[], None]) -> None:
        self.ui_queue.put(fn)

    def status(self, message: str, kind: str = "busy") -> None:
        """Safe to call from any thread."""
        self.post(lambda: self._set_status(message, kind))

    def _set_status(self, message: str, kind: str = "busy") -> None:
        self.status_var.set(message)
        self.status_dot.set(kind)
        if kind == "busy":
            self.shimmer.start()
        else:
            self.shimmer.stop()

    def _poll(self) -> None:
        try:
            while True:
                fn = self.ui_queue.get_nowait()
                try:
                    fn()
                except Exception:
                    traceback.print_exc()
        except queue.Empty:
            pass
        self.root.after(UI_POLL_MS, self._poll)

    def _serve(self, server: socket.socket) -> None:
        """Accept "capture" commands from the background listener (or a second launch)."""
        while True:
            try:
                conn, _ = server.accept()
            except OSError:
                return
            with conn:
                try:
                    data = conn.recv(64).strip()
                except OSError:
                    continue
            if data == b"capture":
                self._on_hotkey()

    @staticmethod
    def _set_text(widget: tk.Text, value: str, muted: bool = False) -> None:
        readonly = str(widget.cget("state")) == "disabled"
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value, ("muted",) if muted else ())
        if readonly:
            widget.configure(state="disabled")
        else:
            widget.edit_reset()  # Ctrl+Z shouldn't bring back placeholders

    def _set_original_placeholder(self, message: str) -> None:
        self.original_placeholder = True
        self._set_text(self.original_text, message, muted=True)

    def _clear_original_placeholder(self, _event: Any = None) -> None:
        if self.original_placeholder:
            self.original_placeholder = False
            self._set_text(self.original_text, "")

    def _set_summary(self, text: str, method: str, muted: bool = False) -> None:
        if method:
            self.method_badge.set_text(method.upper())
            self.method_badge.pack(side="left", padx=(12, 0))
        else:
            self.method_badge.pack_forget()
        self._set_text(self.summary_text, text, muted=muted)

    def _set_table(self, rows: Optional[list[list[str]]]) -> None:
        self.table_tsv = table_to_tsv(rows) if rows else ""
        if rows and self.table_tsv:
            from tkinter import font as tkfont
            measure = tkfont.Font(font=self.table_text.cget("font")).measure
            stops, x = [], 0
            for col in range(len(rows[0]) - 1):  # tab stop after each column's widest cell
                x += max(measure(row[col]) for row in rows) + measure("    ")
                stops.append(x)
            self.table_text.configure(tabs=tuple(stops))
            self._set_text(self.table_text, self.table_tsv)
        else:
            self._set_text(self.table_text, "No table detected in this selection.", muted=True)
        self.table_copy_btn.set_enabled(bool(self.table_tsv))

    def _flash(self, button: PillButton, message: str) -> None:
        original = self._flash_text.setdefault(id(button), button.text)
        button.set_text(message)
        self.root.after(1300, lambda: button.set_text(original))

    # ---- hotkey and overlay -----------------------------------------------

    def start_hotkey(self) -> None:
        self.hotkey = make_hotkey_listener(self._on_hotkey)
        self.hotkey.start()

    def _on_hotkey(self) -> None:
        """Runs on pynput's or the server's thread: return immediately."""
        with self._capture_lock:
            if self.capture_busy.is_set():
                return
            self.capture_busy.set()
        threading.Thread(target=self._grab_for_overlay, daemon=True).start()

    def new_capture(self) -> None:
        self.root.withdraw()  # keep this window out of the screenshot
        self.root.after(250, self._on_hotkey)

    def _grab_for_overlay(self) -> None:
        try:
            capture = grab_monitor_under_cursor()
        except Exception as exc:
            self.capture_busy.clear()
            self.status(f"Screen capture failed: {exc}", "err")
            return
        self.post(lambda: self._show_overlay(capture))

    def _show_overlay(self, capture: Capture) -> None:
        from PIL import ImageEnhance, ImageTk

        accent, accent2 = GRADIENT[0], GRADIENT[1]
        width, height = capture.image.size
        overlay = tk.Toplevel(self.root)
        overlay.overrideredirect(True)
        overlay.geometry(f"{width}x{height}+{capture.left}+{capture.top}")
        overlay.attributes("-topmost", True)
        canvas = tk.Canvas(overlay, width=width, height=height, highlightthickness=0, bd=0,
                           cursor="crosshair", bg="black")
        canvas.pack(fill="both", expand=True)
        dim = ImageTk.PhotoImage(ImageEnhance.Brightness(capture.image).enhance(OVERLAY_BRIGHTNESS))
        canvas.create_image(0, 0, image=dim, anchor="nw")
        spot = canvas.create_image(0, 0, anchor="nw", state="hidden")  # undimmed selection
        glow = canvas.create_rectangle(0, 0, 0, 0, outline=mix(accent, "#000000", 0.35), width=6,
                                       state="hidden")
        rect = canvas.create_rectangle(0, 0, 0, 0, outline=accent2, width=2, state="hidden")
        ants = canvas.create_rectangle(0, 0, 0, 0, outline="#FFFFFF", width=1, dash=(6, 6),
                                       state="hidden")
        handles = [canvas.create_oval(0, 0, 0, 0, fill="#FFFFFF", outline=accent, width=2,
                                      state="hidden") for _ in range(4)]
        size_text = canvas.create_text(0, 0, anchor="nw", fill="#FFFFFF", state="hidden",
                                       font=(FONT_FAMILY, 9, "bold"))
        cx = width // 2
        hint_text = canvas.create_text(cx, 54, fill="#FFFFFF", font=(FONT_FAMILY, 11, "bold"),
                                       text="✦  Drag over the text you want   ·   "
                                            "Esc or right-click to cancel")
        tx0, ty0, tx1, ty1 = canvas.bbox(hint_text)
        hint_photo = ImageTk.PhotoImage(pill_rgba(tx1 - tx0 + 60, ty1 - ty0 + 24, GRADIENT))
        hint = [canvas.create_image(cx, 54, image=hint_photo), hint_text]
        canvas.tag_raise(hint_text)
        drag: dict[str, Any] = {"start": None, "box": None, "pending": False, "dash": 0,
                                "size_bg": None, "refs": [dim, hint_photo]}

        def box_from(event: Any) -> tuple[int, int, int, int]:
            x0, y0 = drag["start"]
            return (max(0, min(x0, event.x)), max(0, min(y0, event.y)),
                    min(width, max(x0, event.x)), min(height, max(y0, event.y)))

        def set_hidden(items: Sequence[int], hidden: bool) -> None:
            for item in items:
                canvas.itemconfigure(item, state="hidden" if hidden else "normal")

        def refresh() -> None:
            drag["pending"] = False
            if drag["box"] is None:
                return
            x0, y0, x1, y1 = drag["box"]
            if x1 - x0 >= 2 and y1 - y0 >= 2:
                photo = ImageTk.PhotoImage(capture.image.crop((x0, y0, x1, y1)))
                drag["refs"] = [dim, hint_photo, photo]
                canvas.itemconfigure(spot, image=photo, state="normal")
                canvas.coords(spot, x0, y0)
            for item in (glow, rect, ants):
                canvas.coords(item, x0, y0, x1, y1)
            for item, (hx, hy) in zip(handles, ((x0, y0), (x1, y0), (x0, y1), (x1, y1))):
                canvas.coords(item, hx - 4, hy - 4, hx + 4, hy + 4)
            label_y = y1 + 10 if y1 + 34 < height else y0 - 30
            canvas.itemconfigure(size_text, text=f"{x1 - x0} × {y1 - y0} px")
            set_hidden([glow, rect, ants, size_text, *handles], False)
            canvas.coords(size_text, x0 + 10, label_y + 4)
            bx0, by0, bx1, by1 = canvas.bbox(size_text)  # bbox() is None while hidden
            if drag["size_bg"]:
                canvas.delete(drag["size_bg"])
            drag["size_bg"] = round_rect(canvas, bx0 - 10, by0 - 4, bx1 + 10, by1 + 4, 9,
                                         fill="#1B1433", outline=accent)
            canvas.tag_raise(size_text)

        def march() -> None:
            if not overlay.winfo_exists():
                return
            drag["dash"] = (drag["dash"] + 2) % 12
            canvas.itemconfigure(ants, dashoffset=drag["dash"])
            overlay.after(70, march)

        def reset() -> None:
            drag["start"] = drag["box"] = None
            set_hidden([spot, glow, rect, ants, size_text, *handles], True)
            if drag["size_bg"]:
                canvas.delete(drag["size_bg"])
                drag["size_bg"] = None
            set_hidden(hint, False)

        def close(_event: Any = None) -> None:
            overlay.destroy()
            self.capture_busy.clear()

        def on_press(event: Any) -> None:
            drag["start"] = (event.x, event.y)
            drag["box"] = (event.x, event.y, event.x, event.y)
            set_hidden(hint, True)
            refresh()

        def on_drag(event: Any) -> None:
            if drag["start"] is None:
                return
            drag["box"] = box_from(event)
            if not drag["pending"]:
                drag["pending"] = True
                canvas.after(15, refresh)

        def on_release(event: Any) -> None:
            if drag["start"] is None:
                return
            box = box_from(event)
            if box[2] - box[0] < MIN_SELECTION_PX or box[3] - box[1] < MIN_SELECTION_PX:
                reset()  # too small: ignore, let the user drag again
                return
            close()
            self.on_capture(capture.image.crop(box))  # crop the pre-overlay screenshot

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        canvas.bind("<Button-3>", close)
        overlay.bind("<Escape>", close)
        overlay.lift()
        overlay.focus_force()
        overlay.after(50, overlay.focus_force)
        march()

    # ---- pipeline: capture -> OCR -> translate -> summary -------------------

    def on_capture(self, image: PILImage) -> None:
        self.screenshot = image
        self._show_window()
        self._set_thumbnail(image)
        self.request_ocr()

    def request_ocr(self) -> None:
        if self.screenshot is None:
            return
        self.ocr_job += 1
        job = self.ocr_job
        name = self.source_var.get()
        auto = name == AUTO_DETECT
        script = "latin" if auto else SCRIPT_OF_LANGUAGE[name]
        image = self.screenshot
        self._set_detected(None)
        loading = "" if self.ocr.is_loaded(script) else "\nThe first capture also loads the OCR model."
        self._set_original_placeholder(f"Reading text from the image... please wait.{loading}")
        self._set_text(self.translation_text, "Waiting for the text...", muted=True)
        self._set_summary("", "")
        self._set_table(None)
        self._set_status("Detecting language and reading text..." if auto else "Reading text...")

        def work() -> None:
            if job != self.ocr_job:
                return
            if not self.ocr.is_loaded("latin") or not self.ocr.is_loaded(script):
                self.status("Loading OCR model... first use downloads it")
            try:
                if auto:
                    text, table, detected = self.ocr.read_auto(image, self.status)
                else:
                    (text, table), detected = self.ocr.read(image, name), None
            except Exception as exc:
                message = f"OCR failed: {exc}"
                self.post(lambda: self._ocr_failed(job, message))
                return
            self.post(lambda: self._ocr_done(job, text, table, detected))

        self.heavy.submit(work)

    def _ocr_failed(self, job: int, message: str) -> None:
        if job == self.ocr_job:
            self._set_original_placeholder("No text could be read.")
            self._set_text(self.translation_text, f"⚠ {message}")
            self._set_status(message, "err")

    def _set_detected(self, language: Optional[str]) -> None:
        """Remember the auto-detected language and show it above the Text in image picker."""
        self.detected_source = language
        if language:
            short = language.split(" (")[0].upper()  # fits the column: "DETECTED: PORTUGUESE"
            self.source_caption.configure(text=f"DETECTED: {short}", fg=THEME["link"])
        else:
            self.source_caption.configure(text="TEXT IN IMAGE", fg=THEME["muted"])

    def _source_language(self) -> Optional[str]:
        name = self.source_var.get()
        return self.detected_source if name == AUTO_DETECT else name

    def _ocr_done(self, job: int, text: str, table: Optional[list[list[str]]],
                  detected: Optional[str] = None) -> None:
        if job != self.ocr_job:
            return
        self._set_detected(detected)
        self._set_table(table)
        if not text.strip():
            self._set_original_placeholder("No text found in this selection.")
            self._set_text(self.translation_text, "")
            self._set_status("No text found in the selection.", "idle")
            return
        self.original_placeholder = False
        self._set_text(self.original_text, text)
        self.request_translation()

    def request_translation(self) -> None:
        source = self._source_language()
        if self.original_placeholder or source is None:  # nothing real to translate yet
            return
        self.translate_job += 1
        job = self.translate_job
        text = self.original_text.get("1.0", "end-1c")
        src = LANGUAGES[source][1]
        target_name = self.target_var.get()
        tgt = LANGUAGES[target_name][1]
        join_lines = self.join_var.get()
        self._set_summary("", "")
        if not text.strip():
            self._set_text(self.translation_text, "")
            return
        self._set_text(self.translation_text, "Translating...", muted=True)
        self._set_status("Translating...")

        def work() -> None:
            if job != self.translate_job:
                return
            ok = False

            def pack_status(message: str) -> None:  # explain big downloads in the Translation box
                hint = ("\n\nWrong language? Pick the right one under 'Text in image' — "
                        "that translates right away; this download keeps going in the background.")
                self.status(message)
                self.post(lambda: job == self.translate_job and self._set_text(
                    self.translation_text, message + hint, muted=True))

            try:
                self.translator.ensure_pairs(required_pairs(src, tgt), pack_status)
                if job != self.translate_job:
                    return  # superseded while the pack downloaded
                self.status("Translating...")
                result = translate_document(text, src, tgt, join_lines, self.translator.translate)
                ok = True
            except PackNotFoundError as exc:
                result = f"⚠ {exc}\nPick another 'Translate to' language."
            except Exception as exc:
                result = (f"⚠ Translation failed: {exc}\n"
                          "The first use of a language pair needs internet to download its pack.")
            self.post(lambda: self._translation_done(job, result, ok, target_name))

        # Own thread, not the OCR queue: a pack download must never hold up re-reading the
        # image or translating with another language (stale jobs are dropped by job id).
        threading.Thread(target=work, name="translate", daemon=True).start()

    def _translation_done(self, job: int, result: str, ok: bool, language: str) -> None:
        if job != self.translate_job:
            return
        self._set_text(self.translation_text, result)
        if not ok:
            self._set_status("Translation unavailable.", "err")
            return
        if OLLAMA_MODEL:
            self._set_summary("Writing AI summary... this can take 15-60 seconds on a laptop CPU.",
                              "", muted=True)
            self._set_status("Writing AI summary...")
        else:
            self._set_summary("Summarizing...", "", muted=True)
            self._set_status("Summarizing...")

        def work() -> None:
            if job != self.translate_job:
                return
            summary, method = summarize(result, language)
            self.post(lambda: self._summary_done(job, summary, method))

        self.light.submit(work)

    def _summary_done(self, job: int, summary: str, method: str) -> None:
        if job == self.translate_job:
            self._set_summary(summary, method)
            self._set_status("Done", "ok")

    # ---- window, clipboard, images -------------------------------------------

    def _show_window(self) -> None:
        r = self.root
        was_hidden = r.state() == "withdrawn"
        if was_hidden:
            try:
                r.attributes("-alpha", 0.0)
            except tk.TclError:
                was_hidden = False
        r.deiconify()
        r.update_idletasks()
        if r.state() == "iconic":  # background-launched processes can open minimized
            r.state("normal")
        set_dark_title_bar(r)
        r.lift()
        r.attributes("-topmost", True)
        r.after(300, lambda: r.attributes("-topmost", False))
        r.focus_force()
        if was_hidden:
            self._fade_in(1)

    def _fade_in(self, step: int, steps: int = 9) -> None:
        self.root.attributes("-alpha", min(1.0, step / steps))
        if step < steps:
            self.root.after(15, lambda: self._fade_in(step + 1, steps))

    def _set_thumbnail(self, image: PILImage) -> None:
        preview = image.convert("RGB")  # shrink once; resizes then work on this small copy
        preview.thumbnail((THUMBNAIL_SIZE[0] * 2, THUMBNAIL_SIZE[1] * 2))
        self._preview = preview
        self.root.update_idletasks()
        self._render_thumbnail()

    def _schedule_thumbnail(self) -> None:
        if not self._thumbnail_pending:
            self._thumbnail_pending = True
            self.root.after(40, self._render_thumbnail)

    def _render_thumbnail(self) -> None:
        """Rounded preview well with the capture fitted inside."""
        from PIL import Image, ImageTk
        self._thumbnail_pending = False
        w, h = self.thumb_canvas.winfo_width(), self.thumb_canvas.winfo_height()
        if w < 40 or h < 30:
            return
        well = pill_image(w, h, THEME["field"], THEME["card"], THEME["field_border"], radius=14)
        if self._preview is not None:
            thumb = self._preview.copy()
            thumb.thumbnail((w - 24, h - 24))
            tw, th = thumb.size
            well.paste(thumb, ((w - tw) // 2, (h - th) // 2), rounded_mask(tw, th, 8))
        self.thumbnail_ref = ImageTk.PhotoImage(well)
        self.thumb_canvas.itemconfigure(self._thumb_bg, image=self.thumbnail_ref)
        self.thumb_canvas.coords(self._thumb_hint, w / 2, h / 2)
        self.thumb_canvas.itemconfigure(self._thumb_hint,
                                        state="hidden" if self._preview is not None else "normal")

    def _copy_text(self, widget: tk.Text, button: PillButton) -> None:
        if widget is self.original_text and self.original_placeholder:
            return
        self._copy_string(widget.get("1.0", "end-1c"))
        self._flash(button, "Copied ✓")

    def _copy_table(self) -> None:
        if self.table_tsv:
            self._copy_string(self.table_tsv)
            self._flash(self.table_copy_btn, "Copied ✓")
            self._set_status("Table copied. Paste into Excel with Ctrl+V.", "ok")

    def _copy_string(self, value: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(value)

    def _copy_image(self) -> None:
        if self.screenshot is None:
            return
        try:
            self._set_status(copy_image_to_clipboard(self.screenshot), "ok")
            self._flash(self.copy_image_btn, "Copied ✓")
        except Exception as exc:
            self._set_status(f"Copy image failed: {exc}", "err")

    def _save_image(self) -> None:
        if self.screenshot is None:
            return
        try:
            path = save_png(self.screenshot)
            reveal_file(path)
            self._set_status(f"Saved {path.name} (temporary, deleted when you quit)", "ok")
        except Exception as exc:
            self._set_status(f"Save failed: {exc}", "err")

    def quit(self) -> None:
        if self.hotkey is not None:
            self.hotkey.stop()
        if self.server is not None:
            self.server.close()
        shutil.rmtree(TEMP_DIR, ignore_errors=True)
        self.root.destroy()


# =============================================================================
# Self-test (no GUI, no hotkey, no downloads)
# =============================================================================

def run_selftest() -> int:
    failures = 0

    def check(name: str, condition: bool, detail: Any = "") -> None:
        nonlocal failures
        print(f"[{'PASS' if condition else 'FAIL'}] {name}"
              + ("" if condition else f"\n       got: {ascii(detail)}"))
        failures += not condition

    def box(x0: float, y0: float, x1: float, y1: float, text: str) -> tuple:
        return ([[x0, y0], [x1, y0], [x1, y1], [x0, y1]], text, 0.99)

    # 1. Line grouping + paragraph break (boxes shuffled, slight vertical jitter).
    fake = [box(120, 12, 220, 32, "world"), box(10, 136, 180, 156, "continues here"),
            box(10, 10, 110, 30, "Hello"), box(10, 110, 250, 130, "New paragraph"),
            box(10, 38, 200, 58, "second line")]
    text, table = reconstruct(fake)
    check("line grouping keeps layout and paragraph break",
          text == "Hello world\nsecond line\n\nNew paragraph\ncontinues here", text)
    check("plain text is not mistaken for a table", table is None, table)

    # 2. Table detection with an empty cell -> TSV for Excel.
    fake_table = [box(10, 10, 80, 30, "Item"), box(200, 10, 260, 30, "Qty"),
                  box(350, 10, 420, 30, "Price"), box(10, 40, 90, 60, "Apple"),
                  box(205, 40, 230, 60, "3"), box(350, 40, 400, 60, "1.20"),
                  box(10, 70, 95, 90, "Banana"), box(350, 70, 400, 90, "0.50")]
    _, rows = reconstruct(fake_table)
    tsv = table_to_tsv(rows) if rows else ""
    check("table -> tab-separated rows, empty cell kept",
          tsv == "Item\tQty\tPrice\nApple\t3\t1.20\nBanana\t\t0.50", tsv)

    # 3. Extractive summary, English.
    english = ("Solar power is growing quickly around the world. Many countries now install more "
               "solar panels than ever before. The weather was pleasant yesterday. Cheaper solar "
               "panels make solar power attractive for homes and countries. My cat likes to sleep. "
               "Experts expect solar power to keep growing as panels get cheaper.")
    summary = extractive_summary(english)
    picked = split_sentences(summary)
    originals = split_sentences(english)
    in_order = [originals.index(s) for s in picked if s in originals]
    check("English summary: 3 sentences, original order, off-topic dropped",
          len(picked) == 3 and len(in_order) == 3 and in_order == sorted(in_order)
          and "cat" not in summary and "weather" not in summary, summary)
    check("short text (<=2 sentences) returned as is",
          extractive_summary("One sentence. Two sentences.") == "One sentence. Two sentences.")

    # 4. Extractive summary, Chinese (each CJK character is one token).
    chinese = ("太阳能发电在全球快速增长。许多国家安装了大量太阳能电池板。昨天天气很好。"
               "便宜的太阳能电池板让太阳能更受欢迎。专家预计太阳能将继续增长。")
    zh_summary = extractive_summary(chinese)
    zh_picked = split_sentences(zh_summary)
    zh_all = split_sentences(chinese)
    zh_order = [zh_all.index(s) for s in zh_picked if s in zh_all]
    check("CJK tokenizer: one token per character", tokenize("太阳能") == ["太", "阳", "能"])
    check("Chinese summary: 3 sentences, original order, off-topic dropped",
          len(zh_picked) == 3 and len(zh_order) == 3 and zh_order == sorted(zh_order)
          and "天气" not in zh_summary, zh_summary)

    # 5. Hindi sentence splitting on the danda.
    hindi = "यह पहला वाक्य है। यह दूसरा वाक्य है। क्या यह तीसरा वाक्य है?"
    hi_sentences = split_sentences(hindi)
    check("Hindi splitter: 3 sentences on '।' and '?'",
          len(hi_sentences) == 3 and hi_sentences[0] == "यह पहला वाक्य है।", hi_sentences)
    check("Hindi tokenizer keeps vowel signs inside words",
          tokenize("हिंदी भाषा") == ["हिंदी", "भाषा"], tokenize("हिंदी भाषा"))

    # 6. Manual pivot logic with a mocked translator.
    calls: list[tuple[str, str]] = []

    def mock(t: str, a: str, b: str) -> str:
        calls.append((a, b))
        return f"[{a}>{b}]{t}"

    out = translate_with_pivot("x", "hi", "ja", mock)
    check("pivot hi->ja = two calls via English",
          out == "[en>ja][hi>en]x" and calls == [("hi", "en"), ("en", "ja")], (out, calls))
    calls.clear()
    out = translate_with_pivot("x", "en", "hi", mock)
    check("direct en->hi = one call", out == "[en>hi]x" and calls == [("en", "hi")], (out, calls))
    calls.clear()
    out = translate_document("same\ntext", "hi", "hi", True, mock)
    check("source == target returns text unchanged", out == "same\ntext" and not calls, out)
    doc = "line one\nline two\n\nsecond para"
    out = translate_document(doc, "hi", "en", True, mock)
    check("join on: one call per paragraph",
          out == "[hi>en]line one line two\n\n[hi>en]second para", out)
    out = translate_document(doc, "hi", "en", False, mock)
    check("join off: one call per line, blank line kept",
          out == "[hi>en]line one\n[hi>en]line two\n\n[hi>en]second para", out)

    # 7. Language table.
    check("language table: 14 languages, Chinese = ch_sim/zh",
          len(LANGUAGES) == 14 and LANGUAGES["Chinese (Simplified)"] == ("ch_sim", "zh"))
    check("every language has a writing system with a reader",
          all(SCRIPT_OF_LANGUAGE.get(n) in SCRIPT_READER_LANGS for n in LANGUAGES))
    check("auto-detect tries every non-Latin script, Japanese right after Chinese",
          sorted(NON_LATIN_TRY_ORDER) == sorted(set(SCRIPT_READER_LANGS) - {"latin"})
          and NON_LATIN_TRY_ORDER.index("japanese") == NON_LATIN_TRY_ORDER.index("chinese") + 1)

    # 8. Auto-detect: Latin-script languages from paragraphs.
    paragraphs = {
        "English": "The weather is nice today and the children are playing in the park with their dog.",
        "Spanish": "El tiempo está muy bien hoy y los niños juegan en el parque con su perro.",
        "French": "Le temps est très beau aujourd'hui et les enfants jouent dans le parc avec leur chien.",
        "German": "Das Wetter ist heute sehr schön und die Kinder spielen mit ihrem Hund im Park.",
        "Portuguese": "O tempo está muito bom hoje e as crianças estão brincando no parque com o cão.",
        "Italian": "Il tempo è molto bello oggi e i bambini giocano nel parco con il loro cane.",
        "Turkish": "Bugün hava çok güzel ve çocuklar köpekleriyle birlikte parkta oynuyorlar.",
    }
    wrong = {lang: detect_latin_language(p) for lang, p in paragraphs.items()
             if detect_latin_language(p) != lang}
    check("auto-detect: 7 Latin-script languages from a sentence each", not wrong, wrong)
    shorts = {"Thank you very much for the help": "English", "¿Dónde está la estación?": "Spanish",
              "Straße gesperrt": "German", "Hoşgeldiniz, değil mi?": "Turkish"}
    wrong = {t: detect_latin_language(t) for t, lang in shorts.items()
             if detect_latin_language(t) != lang}
    check("auto-detect: short phrases with telltale words/letters", not wrong, wrong)
    check("auto-detect: no clues falls back to English", detect_latin_language("OK 2024") == "English")
    german_labels = ["Einstellungen speichern", "Abbrechen", "Weiter zur Kasse",
                     "Wetterbericht Montag: Regen, Wind, Temperaturen bis 12 Grad",
                     "Kostenloser Versand ab 50 Euro", "Vielen Dank für Ihren Einkauf"]
    wrong = {t: detect_latin_language(t) for t in german_labels
             if detect_latin_language(t) != "German"}
    check("auto-detect: short German labels/headlines (once misread as Spanish)", not wrong, wrong)

    # 9. Auto-detect: script -> language, and choosing a script from competing OCR outputs.
    check("auto-detect: Urdu vs Arabic by Urdu-only letters",
          language_for_script("arabic", "یہ ایک کتاب ہے") == "Urdu"
          and language_for_script("arabic", "هذا كتاب جميل") == "Arabic")
    check("auto-detect: Japanese (kana) vs Chinese (no kana)",
          language_for_script("japanese", "これは本です") == "Japanese"
          and language_for_script("japanese", "这是一本书") == "Chinese (Simplified)")
    check("auto-detect: other scripts map to their language",
          [language_for_script(s, "") for s in ("cyrillic", "devanagari", "korean")]
          == ["Russian", "Hindi", "Korean"])
    hindi_page = {"latin": ("Ya ek kitab hai", 0.21), "devanagari": ("यह एक किताब है", 0.88),
                  "cyrillic": ("уа ек", 0.15), "arabic": ("یه", 0.12)}
    check("choose_script: confident native output wins (Hindi page)",
          choose_script(hindi_page) == "devanagari", choose_script(hindi_page))
    english_seen_by_cjk = {"latin": ("Hello world", 0.62), "chinese": ("Hello world", 0.95)}
    check("choose_script: Latin letters from a CJK reader don't count as Chinese",
          choose_script(english_seen_by_cjk) == "latin", choose_script(english_seen_by_cjk))
    japanese_page = {"chinese": ("私は学生です", 0.80), "japanese": ("私は学生です", 0.78)}
    check("choose_script: kana make it Japanese even if the Chinese reader scored higher",
          choose_script(japanese_page) == "japanese", choose_script(japanese_page))
    check("mean_confidence weights by text length",
          abs(mean_confidence([([], "aaaa", 1.0), ([], "b", 0.0)]) - 0.8) < 1e-9)

    print(f"\n{'ALL TESTS PASSED' if not failures else f'{failures} TEST(S) FAILED'}")
    return 1 if failures else 0


# =============================================================================
# Entry point
# =============================================================================

def enable_dpi_awareness() -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor aware
        except Exception:
            pass


def run_gui(start_capture: bool) -> int:
    if tk is None:
        print("Tkinter is missing. Install Python from python.org (Windows) or python3-tk (Linux).")
        return 1
    missing = [m for m in ("mss", "PIL", "easyocr", "argostranslate", "pynput", "numpy")
               if importlib.util.find_spec(m) is None]
    if missing:
        print(f"Missing packages: {', '.join(missing)}\n"
              f"Install with: {Path(sys.executable).name} -m pip install -r requirements.txt")
        return 1
    try:
        server = socket.create_server(("127.0.0.1", APP_PORT))
    except OSError:  # already running: ask that instance to capture instead
        send_command(APP_PORT, "capture")
        return 0

    enable_dpi_awareness()  # before any Tk window
    shutil.rmtree(TEMP_DIR, ignore_errors=True)  # leftovers from a crashed session
    root = tk.Tk()
    root.withdraw()
    app = App(root, server)
    if send_command(LISTENER_PORT, "ping"):
        print("Running. The background listener handles Ctrl+Shift+T.", flush=True)
    else:
        try:
            app.start_hotkey()
            print("Running. Press Ctrl+Shift+T to capture. Use the Quit button to exit.", flush=True)
        except Exception as exc:
            app._set_status(f"Hotkey unavailable: {exc}", "err")
            root.deiconify()
    if start_capture:
        root.after(100, app._on_hotkey)
    root.mainloop()
    return 0


def redirect_output_if_windowless() -> None:
    """pythonw has no console: send prints and tracebacks to the log file."""
    if sys.stdout is not None:
        return
    try:
        if LOG_PATH.exists() and LOG_PATH.stat().st_size > 1_000_000:
            LOG_PATH.unlink()
    except OSError:
        pass
    log = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stderr = log


def main(argv: Optional[list[str]] = None) -> int:
    redirect_output_if_windowless()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Screen OCR + offline translator")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--selftest", action="store_true",
                      help="run logic tests without GUI, hotkey or downloads")
    mode.add_argument("--capture", action="store_true", help="start straight in capture mode")
    mode.add_argument("--listener", action="store_true",
                      help="run only the tiny background hotkey listener")
    mode.add_argument("--install", action="store_true",
                      help="start the listener now and at every login")
    mode.add_argument("--uninstall", action="store_true", help="undo --install")
    args = parser.parse_args(argv)
    if args.selftest:
        return run_selftest()
    if args.listener:
        return run_listener()
    if args.install:
        return install_autostart()
    if args.uninstall:
        return uninstall_autostart()
    return run_gui(start_capture=args.capture)


if __name__ == "__main__":
    sys.exit(main())
