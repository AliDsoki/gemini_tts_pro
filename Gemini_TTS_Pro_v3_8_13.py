# -*- coding: utf-8 -*-
"""
🌟 Gemini TTS Pro v3.8.13 — محوّل النصوص إلى كلام (استوديو إنتاج احترافي متكامل)
================================================================================
تطبيق استوديو متكامل وتحكم فائق لتوليد الكلام العربي والدولي بدقة عالية
باستخدام Google Gemini TTS API.

✨ الميزات الحصرية المدمجة في الإصدار v3.8.13:
- معالجة تلقائية وحل جذري لأخطاء 400 Developer Instruction لجميع النماذج المخصصة للصوت.
- إدارة استباقية للمفاتيح والكوتة (Adaptive RPM Pacing + Round-Robin Rotation).
- طرق تقسيم نص متعددة مدمجة: بالمدة الزمنية (الثواني)، بعد الكلمات، أو بعد الجمل.
- معالجة وتشكيل الرموز التلقائي (Auto-clean URL/Symbols & Diacritics Removal).
- نقل الإيقاع العاطفي والسياق بين المقاطع (Context Carryover).
- توازي متكيف (Adaptive Parallel Workers): خفض وزيادة العمال ذاتياً عند ضغط الشبكة.
- وضع الإعادة اللانهائية (Retry Failed Forever): إعادة الأجزاء الفاشلة لطابور المعالجة حتى اكتمال المشروع.
- محرك دمج محلي سريع ومستقل (Pure Python Native WAV Merger) مع فواصل صمت مخصصة.
- تقسيم الناتج النهائي المدمج تلقائياً (WAV/Audio Splitter) حسب الأجزاء أو المدة أو الحجم مع تداخل زمني.
- استئناف المشاريع المتوقفة تلقائياً عبر نقطة الفحص (Atomic Checkpoints).
- خيار إيقاف تشغيل الكمبيوتر تلقائياً عند اكتمال المشروع والدمج.
- واجهة ثنائية المظهر (ليلي داكن ونهاري فاتح) فائقة الوضوح والتباين لجميع النوافذ والحوارات.
"""

import os
import sys
import re
import json
import time
import wave
import base64
import shutil
import subprocess
import threading
import traceback
import hashlib
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict, Any, Tuple, Set, Callable
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from enum import Enum

# فرض ترميز UTF-8 لبيئة بايثون
os.environ["PYTHONUTF8"] = "1"

try:
    from google import genai
    from google.genai import types
    from google.genai import errors as genai_errors
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pypdf import PdfReader
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QPushButton, QComboBox, QSlider, QSpinBox,
    QLineEdit, QTextEdit, QFileDialog, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QSplitter, QDialog, QCheckBox,
    QMenuBar, QMenu, QStatusBar, QMessageBox, QSystemTrayIcon,
    QStyle, QDialogButtonBox, QFormLayout, QScrollArea, QAbstractItemView,
    QPlainTextEdit, QGridLayout, QListWidget, QListWidgetItem, QFrame, QStackedWidget,
    QTabWidget, QTabBar
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QObject, QMutex, QMutexLocker, QTimer, QSize
)
from PyQt6.QtGui import (
    QAction, QFont, QColor, QKeySequence, QTextCursor, QBrush, QShortcut
)


# ═══════════════════════════════════════════════════════════════
# CONSTANTS & CONFIGURATION DEFAULTS
# ═══════════════════════════════════════════════════════════════

APP_NAME = "Gemini TTS Pro"
APP_VERSION = "3.8.13"

APP_DIR = Path(__file__).resolve().parent
CONFIG_DIR = Path.home() / ".gemini_tts_pro"
CONFIG_FILE = CONFIG_DIR / "config.json"
VOICE_PREVIEWS_DIR = CONFIG_DIR / "voice_previews"
TOKENS_FILE_NAME = "tokens.txt"
KEYS_STATE_FILE = "keys_state.json"
GLOBAL_KEYS_STATE_FILE = "keys_global_state.json"

CHARS_PER_SECOND = 15
MAX_CHUNK_CHARS = 4000
MAX_LOG_LINES = 1000
DEFAULT_SAMPLE_RATE = 24000
STOP_CHECK_INTERVAL = 0.1

DEFAULT_MODELS = [
    {"id": "gemini-2.5-flash-preview-tts", "label": "Gemini 2.5 Flash TTS (سريع وخفيف)", "family": "2.5"},
    {"id": "gemini-2.5-pro-preview-tts", "label": "Gemini 2.5 Pro TTS (أعلى جودة وتعبير)", "family": "2.5"},
]

VOICES = [
    "Sulafat", "Gacrux", "Charon", "Umbriel", "Vindemiatrix", "Algenib", "Despina", "Algieba", "Schedar", "Achernar",
    "Rasalgethi", "Sadaltager", "Iapetus", "Erinome", "Kore", "Orus", "Alnilam", "Achird", "Autonoe", "Zephyr",
    "Pulcherrima", "Zubenelgenubi", "Callirrhoe", "Aoede", "Enceladus", "Leda", "Sadachbia", "Laomedeia", "Puck", "Fenrir"
]

VOICE_CATALOG = [
    # ── أولاً: الأفضل للروايات والقصص والأداء الدافئ ──
    ("Sulafat", "أنثى — دافئ (Warm) [روايات]"),
    ("Gacrux", "أنثى — ناضج (Mature) [روايات]"),
    ("Charon", "ذكر — إخباري هادئ (Informative) [روايات]"),
    ("Umbriel", "ذكر — مسترخٍ (Easy-going) [روايات]"),
    ("Vindemiatrix", "أنثى — لطيف (Gentle) [روايات]"),
    ("Algenib", "ذكر — أجشّ/درامي (Gravelly) [روايات]"),
    ("Despina", "أنثى — ناعم (Smooth) [روايات]"),
    ("Algieba", "ذكر — ناعم (Smooth) [روايات]"),
    ("Schedar", "ذكر — متزن الإيقاع (Even) [روايات]"),
    ("Achernar", "أنثى — هادئ (Soft) [روايات]"),
    # ── ثانياً: الأفضل للكتب (غير روائية، إخبارية/تعليمية، وضوح وثقة) ──
    ("Rasalgethi", "ذكر — إخباري (Informative) [كتب/وثائقي]"),
    ("Sadaltager", "ذكر — مطّلع (Knowledgeable) [كتب/أكاديمي]"),
    ("Iapetus", "ذكر — واضح (Clear) [كتب/وثائقي]"),
    ("Erinome", "أنثى — واضح (Clear) [كتب/أكاديمي]"),
    ("Kore", "أنثى — حازم/واثق (Firm) [كتب/رسمي]"),
    ("Orus", "ذكر — حازم (Firm) [كتب/رسمي]"),
    ("Alnilam", "ذكر — حازم (Firm) [كتب/رسمي]"),
    ("Achird", "ذكر — ودود (Friendly) [كتب/شروحات]"),
    ("Autonoe", "أنثى — مشرق (Bright) [كتب/تعليمي]"),
    ("Zephyr", "أنثى — مشرق (Bright) [كتب/تعليمي]"),
    # ── ثالثاً: باقي الأصوات (إعلانات، حوارات، شخصيات حيوية) ──
    ("Pulcherrima", "أنثى — جريء/تعبيري (Forward) [حوارات]"),
    ("Zubenelgenubi", "ذكر — غير رسمي (Casual) [حوارات]"),
    ("Callirrhoe", "أنثى — مسترخٍ (Easy-going) [إعلانات]"),
    ("Aoede", "أنثى — خفيف (Breezy) [إعلانات]"),
    ("Enceladus", "ذكر — نفسي/همسي (Breathy) [شخصيات]"),
    ("Leda", "أنثى — شبابي (Youthful) [شخصيات]"),
    ("Sadachbia", "ذكر — حيوي (Lively) [حماس/إعلانات]"),
    ("Laomedeia", "أنثى — متحمس (Upbeat) [حماس/إعلانات]"),
    ("Puck", "ذكر — متحمس (Upbeat) [شخصيات حيوية]"),
    ("Fenrir", "ذكر — سريع الانفعال (Excitable) [دراما/شخصيات]"),
]

READING_STYLES = {
    "رواية (افتراضي)": (
        "أنت راوٍ وممثل صوتي سينمائي محترف فائق البراعة متخصص في روايات الأدب العربي والدراما. "
        "مطلوب منك أداء النص صوتياً بأعلى درجات التعبير والتشخيص الحي ليعكس المواقف والأحاسيس بدقة متناهية:\n"
        "1. تمثيل المواقف والأحاسيس: اشعر بالموقف الدرامي؛ إذا كان المشهد يحمل توتراً أو خوفاً أو غضباً اجعل نبرتك مشحونة ومضطربة، وإذا كان حزناً أو انكساراً اجعل الإلقاء خافتاً ومُثقلاً بالمشاعر، وإذا كان فرحاً أو حماساً املأ الصوت بالحيوية والإشراق.\n"
        "2. تباين الحوارات وطريقة الكلام: فرّق بوضوح تام بين صوت السارد (الراوي الرصين والمتأني) وبين أصوات الشخصيات أثناء الحوارات؛ غيّر طبقة صوتك وسرعة إلقائك وطريقة كلامك لتناسب طبيعة كل شخصية وحالتها النفسية في اللحظة.\n"
        "3. الإيقاع والوقفات العاطفية: تمهّل والتقط أنفاسك عند التأملات واللحظات المصيرية لتعزيز التشويق، وأسرع بإيقاعك في مشاهد الحركة والمواجهات.\n"
        "4. جودة النطق: حافظ على فصحى عربية سليمة ومخارج حروف واضحة جداً مع تدفق عاطفي طبيعي يأسر المستمع."
    ),
    "إخباري رسمي": (
        "أنت مذيع أخبار محترف في قناة عربية عالمية. "
        "اقرأ النص بأسلوب إخباري رصين، بلغة عربية فصحى متقنة، ومخارج حروف واضحة جداً، "
        "مع إيصال الأهمية والجدية في النبرة دون انفعال مبالغ فيه."
    ),
    "وثائقي هادئ": (
        "أنت معلّق صوتي في فيلم وثائقي استقصائي. "
        "اقرأ النص بنبرة عميقة وهادئة وتأملية، تمهّل عند المعلومات الهامة والحقائق العلمية، "
        "واملأ الصوت بالهيبة والوضوح التام."
    ),
    "تعليمي وأكاديمي": (
        "أنت أستاذ ومحاضر متميز. "
        "اقرأ النص بأسلوب تعليمي واضح، قسّم الجمل بطريقة تساعد على الفهم والاستيعاب، "
        "وأعطِ نبرة مشجعة، واضحة ومبسطة للمتلقي."
    ),
    "إعلاني وتحميسي": (
        "أنت معلّق صوتي للإعلانات التجارية والمقاطع الحماسية. "
        "اقرأ النص بنبرة مليئة بالحيوية والنشاط، مع التركيز على الكلمات المفتاحية وجذب الانتباه."
    ),
    "درامي وتمثيلي": (
        "أنت ممثل صوتي مسرحي. "
        "جسّد المشاعر والنص بكل حيوية، وتفاعل مع الحوارات والتناقضات العاطفية بمرونة وواقعية."
    ),
}

UI = {
    "app_title": f"{APP_NAME} v{APP_VERSION} — استوديو تحويل النصوص إلى كلام",
    "source_text": "📄 مصدر النص والمعاينة",
    "browse": "استعراض...",
    "file_path": "مسار الملف النصي أو المستند:",
    "text_stats": "إحصائيات النص:",
    "chars": "حرف",
    "words": "كلمة",
    "tokens_est": "توكن تقريبي",
    "est_duration": "المدة التقديرية:",
    "minutes": "دقيقة",
    "seconds": "ثانية",
    "preview_text": "🔍 معاينة النص كاملًا",
    "tts_settings": "🎙️ إعدادات الصوت والأداء الصوتي",
    "model": "النموذج الصوتي:",
    "voice": "نبرة الصوت:",
    "preview_voice": "🔊 تجربة الصوت",
    "speed": "سرعة الإلقاء:",
    "lang_hint": "توجيه اللغة:",
    "chunking": "✂️ هندسة وتقسيم النص (Chunking)",
    "split_method_label": "طريقة تقسيم النص:",
    "chunk_duration": "مدة المقطع (ثانية):",
    "chunk_words_label": "عدد الكلمات للمقطع:",
    "chunk_sentences_label": "عدد الجمل للمقطع:",
    "preview_chunks": "📋 جدول ومعاينة المقاطع",
    "context_carry": "✨ نقل الإيقاع والسياق بين المقاطع (Context Carryover)",
    "auto_clean_symbols": "🧹 تنظيف الرموز والروابط والفواصل المكررة تلقائياً",
    "remove_diacritics": "🔤 إزالة التشكيل والحركات لمنع أخطاء نطق الحركات",
    "api_keys": "🔑 إدارة مفاتيح API والكوتة",
    "output": "📁 مسار ومخرجات المشروع",
    "output_folder": "مجلد الإخراج الرئيسي:",
    "project_name": "اسم المشروع:",
    "auto_merge": "⚡ دمج المقاطع تلقائياً فور اكتمال المشروع",
    "audio_bitrate_label": "ضغط وإعادة ترميز الصوت المدمج (MP3):",
    "start": "▶ بدء المعالجة",
    "pause": "⏸ إيقاف مؤقت",
    "stop": "⏹ إيقاف",
    "merge_available": "🔀 دمج المتاح الآن",
    "retry_failed_all": "🔁 إعادة محاولة الفاشلة",
    "open_folder": "📂 مجلد الكتاب",
    "open_root_folder": "📂 المسار الرئيسي",
    "current_chunk": "حالة العمليات:",
    "log_panel": "📝 سجل العمليات والمراقبة الحية",
    "export_log": "💾 تصدير السجل",
    "settings": "⚙️ الإعدادات المتقدمة",
    "about": "ℹ️ حول التطبيق",
    "status_ready": "جاهز للعمل",
    "ffmpeg_missing": "ℹ️ تنبيه: ffmpeg غير مثبت. سيتم استخدام محرك الدمج المدمج السريع (Pure Python WAV Merger).",
    "no_api_keys": "الرجاء إدخال مفتاح API واحد على الأقل للبدء.",
    "no_text": "الرجاء تحميل ملف نصي أولاً أو كتابة نص.",
    "no_output": "الرجاء تحديد مجلد الإخراج.",
    "no_project": "الرجاء إدخال اسم للمشروع.",
    "confirm_stop": "هل تريد إيقاف المعالجة؟ يمكنك استئناف المقاطع المتبقية لاحقاً بنفس اسم المشروع.",
    "merge_success": "✅ تم دمج الملف الصوتي النهائي بنجاح!\nالمسار: ",
    "consecutive_fail": "تنبيه: فشل {0} مقاطع متتالية بسبب ضغط الشبكة أو الكوتة. هل تريد المتابعة أم التوقف مؤقتاً؟",
    "open_file": "فتح ملف نصي أو مستند",
    "file_filter": "جميع الملفات المدعومة (*.txt *.docx *.pdf);;ملفات نصية (*.txt);;مستندات Word (*.docx);;ملفات PDF (*.pdf);;الكل (*.*)",
    "recent_files": "الملفات الأخيرة",
    "max_workers": "أقصى عدد عمال متزامنين (Parallel Workers):",
    "retry_count": "عدد المحاولات التلقائية لكل مقطع:",
    "settings_title": f"الإعدادات المتقدمة — {APP_NAME} v{APP_VERSION}",
    "about_text": f"🌟 {APP_NAME} v{APP_VERSION}\nالإصدار الاحترافي الشامل\n\nتطبيق استوديو متقدم لتحويل النصوص العربية والدولية إلى كلام طبيعي فائق الدقة\nباستخدام تقنيات Google Gemini TTS API.\n\nتطوير وتصميم احترافي مع إدارة ذكية للمفاتيح والكوتة، وتقسيم متعدد الطرق، ودمج محلي سريع.",
    "models_management": "🧠 النماذج المتاحة",
    "fetch_models": "🔄 جلب وتحديث من API",
    "select_models": "📋 تخصيص النماذج",
    "models_status": "الحالة:",
    "add_keys": "➕ إضافة مفاتيح",
    "keys_status": "📊 تفاصيل المفاتيح",
    "validate_keys": "✓ فحص وتنشيط",
    "blocked_keys": "المحظور حالياً:",
}


# ═══════════════════════════════════════════════════════════════
# DATACLASSES
# ═══════════════════════════════════════════════════════════════

@dataclass
class ChunkInfo:
    """Represents a single text chunk for TTS processing."""
    index: int
    text: str
    char_count: int
    estimated_duration: float
    status: str = "pending"
    output_file: str = ""
    time_taken: float = 0.0
    error_msg: str = ""
    context_hint: str = ""


@dataclass
class BookJob:
    """Represents a self-contained book project in the multi-book queue (integrated from PDF Master Pro)."""
    file_path: str
    project_name: str
    status: str = "pending"       # 'pending', 'processing', 'done', 'failed'
    source_text: str = ""
    chunks: List[ChunkInfo] = field(default_factory=list)
    output_final_wav: str = ""
    error_msg: str = ""


@dataclass
class ProjectConfig:
    """Stores project-level configuration integrated from v3.8.13."""
    source_file: str = ""
    model_id: str = "gemini-2.5-flash-preview-tts"
    voice: str = "Kore"
    speed: float = 1.0
    lang_hint: str = "ar"
    
    # Chunking configuration (integrated from v3.8.13)
    split_method: str = "seconds"       # 'seconds', 'words', 'sentences'
    chunk_duration: int = 60            # seconds target
    chunk_words: int = 150              # words target
    chunk_sentences: int = 10           # sentences target
    context_carry: bool = True
    auto_clean_symbols: bool = True     # clean URLs/symbols automatically
    remove_diacritics: bool = False     # strip tashkeel if preferred
    
    # Output & Merging configuration
    output_folder: str = ""
    project_name: str = "project"
    auto_merge: bool = True
    silence_pad_ms: int = 300
    audio_bitrate: str = "copy"         # 'copy' (WAV direct), '48k', '64k', '96k', '128k'
    keep_temp_files: bool = True        # keep chunk_xxxx.wav after merge
    output_split_method: str = "none"   # 'none', 'parts', 'duration', 'size'
    output_split_parts: int = 5
    output_split_minutes: int = 60
    output_split_size_mb: int = 50
    output_split_overlap_sec: int = 10
    
    # Concurrency & Reliability configuration
    max_workers: int = 3
    retry_count: int = 4
    max_retry_rounds: int = 3           # جولات الإعادة التلقائية للدفعات الفاشلة
    adaptive_parallel: bool = True      # auto-scale workers on rate limit
    retry_failed_forever: bool = False  # loop on failed chunks infinitely
    shutdown_after_queue: bool = False  # auto shutdown PC on completion
    
    # UI & Theme
    dark_theme: bool = True
    recent_files: list = field(default_factory=list)
    all_fetched_models: list = field(default_factory=list)
    enabled_model_ids: list = field(default_factory=list)
    reading_style: str = "رواية (افتراضي)"
    window_width: int = 1420
    window_height: int = 860


# ═══════════════════════════════════════════════════════════════
# THEME STYLESHEETS & DIALOG THEME HELPER
# ═══════════════════════════════════════════════════════════════

def get_app_stylesheet(dark_theme: bool = True) -> str:
    """Returns crystal-clear, high-contrast stylesheet for both dark and light modes."""
    if dark_theme:
        return """
            QWidget, QDialog, QMainWindow {
                background-color: #1e1e2e;
                color: #f8f9fa;
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
            QGroupBox {
                background-color: #1e1e2e;
                border: 1px solid #45475a;
                border-radius: 8px;
                margin-top: 1.1em;
                font-weight: bold;
                color: #89b4fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 6px;
                color: #89b4fa;
            }
            QLabel, QCheckBox, QRadioButton {
                color: #f8f9fa;
                font-size: 13px;
                background-color: transparent;
            }
            QLineEdit, QSpinBox, QComboBox, QTextEdit, QPlainTextEdit {
                background-color: #313244;
                border: 1px solid #45475a;
                border-radius: 6px;
                padding: 6px;
                color: #ffffff;
                font-size: 13px;
                selection-background-color: #89b4fa;
                selection-color: #11111b;
            }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #89b4fa;
                background-color: #3b3d54;
            }
            QComboBox::drop-down {
                border: none;
                width: 24px;
            }
            QComboBox QAbstractItemView {
                background-color: #313244;
                color: #ffffff;
                border: 1px solid #89b4fa;
                selection-background-color: #89b4fa;
                selection-color: #11111b;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                background-color: #45475a;
                border-radius: 3px;
            }
            QPushButton {
                background-color: #313244;
                border: 1px solid #45475a;
                border-radius: 6px;
                padding: 7px 16px;
                font-weight: 600;
                color: #f8f9fa;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #45475a;
                border-color: #89b4fa;
                color: #ffffff;
            }
            QPushButton:pressed {
                background-color: #585b70;
            }
            QPushButton#btn_start {
                background-color: #a6e3a1;
                color: #11111b;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#btn_start:hover {
                background-color: #94e28f;
                border-color: #a6e3a1;
            }
            QPushButton#btn_pause {
                background-color: #f9e2af;
                color: #11111b;
                font-weight: bold;
            }
            QPushButton#btn_stop {
                background-color: #f38ba8;
                color: #11111b;
                font-weight: bold;
            }
            QProgressBar {
                border: 1px solid #45475a;
                border-radius: 6px;
                text-align: center;
                background-color: #313244;
                color: #ffffff;
                height: 20px;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #89b4fa;
                border-radius: 5px;
            }
            QTableWidget {
                background-color: #181825;
                alternate-background-color: #181825;
                border: 1px solid #45475a;
                gridline-color: #313244;
                color: #ffffff;
                font-size: 13px;
            }
            QTableWidget::item {
                background-color: #181825;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #313244;
                color: #89b4fa;
                padding: 7px;
                font-weight: bold;
                border: none;
                border-bottom: 1px solid #45475a;
            }
            QTableWidget::item:selected {
                background-color: #45475a;
                color: #ffffff;
            }
            QSplitter::handle {
                background-color: #45475a;
                width: 2px;
            }
            QScrollBar:vertical {
                background: #1e1e2e;
                width: 12px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: #45475a;
                border-radius: 6px;
                min-height: 20px;
            }
            QMenuBar {
                background-color: #181825;
                color: #f8f9fa;
                border-bottom: 1px solid #313244;
            }
            QMenuBar::item:selected {
                background-color: #313244;
                color: #89b4fa;
            }
            QMenu {
                background-color: #313244;
                color: #f8f9fa;
                border: 1px solid #45475a;
            }
            QMenu::item:selected {
                background-color: #45475a;
                color: #89b4fa;
            }
            QStatusBar {
                background-color: #181825;
                color: #a6adc8;
                border-top: 1px solid #313244;
            }
            QTabWidget::pane {
                border: 1px solid #45475a;
                border-radius: 6px;
            }
            QTabBar::tab {
                background: #313244;
                color: #f8f9fa;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background: #89b4fa;
                color: #11111b;
                font-weight: bold;
            }
        """
    else:
        return """
            QWidget, QDialog, QMainWindow {
                background-color: #f8f9fa;
                color: #212529;
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
            QGroupBox {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                margin-top: 1.1em;
                font-weight: bold;
                color: #0d6efd;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 6px;
                color: #0d6efd;
            }
            QLabel, QCheckBox, QRadioButton {
                color: #212529;
                font-size: 13px;
                background-color: transparent;
            }
            QLineEdit, QSpinBox, QComboBox, QTextEdit, QPlainTextEdit {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                border-radius: 6px;
                padding: 6px;
                color: #212529;
                font-size: 13px;
                selection-background-color: #0d6efd;
                selection-color: #ffffff;
            }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #0d6efd;
                background-color: #ffffff;
            }
            QComboBox::drop-down {
                border: none;
                width: 24px;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #212529;
                border: 1px solid #0d6efd;
                selection-background-color: #0d6efd;
                selection-color: #ffffff;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                background-color: #e9ecef;
                border-radius: 3px;
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                border-radius: 6px;
                padding: 7px 16px;
                font-weight: 600;
                color: #212529;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #0d6efd;
                color: #0d6efd;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
            QPushButton#btn_start {
                background-color: #198754;
                color: #ffffff;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#btn_start:hover {
                background-color: #157347;
                border-color: #198754;
            }
            QPushButton#btn_pause {
                background-color: #ffc107;
                color: #000000;
                font-weight: bold;
            }
            QPushButton#btn_stop {
                background-color: #dc3545;
                color: #ffffff;
                font-weight: bold;
            }
            QProgressBar {
                border: 1px solid #ced4da;
                border-radius: 6px;
                text-align: center;
                background-color: #e9ecef;
                color: #212529;
                height: 20px;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #0d6efd;
                border-radius: 5px;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                gridline-color: #e9ecef;
                color: #212529;
                font-size: 13px;
            }
            QHeaderView::section {
                background-color: #e9ecef;
                color: #0d6efd;
                padding: 7px;
                font-weight: bold;
                border: none;
                border-bottom: 1px solid #ced4da;
            }
            QTableWidget::item:selected {
                background-color: #0d6efd;
                color: #ffffff;
            }
            QSplitter::handle {
                background-color: #ced4da;
                width: 2px;
            }
            QScrollBar:vertical {
                background: #f8f9fa;
                width: 12px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: #ced4da;
                border-radius: 6px;
                min-height: 20px;
            }
            QMenuBar {
                background-color: #ffffff;
                color: #212529;
                border-bottom: 1px solid #ced4da;
            }
            QMenuBar::item:selected {
                background-color: #e9ecef;
                color: #0d6efd;
            }
            QMenu {
                background-color: #ffffff;
                color: #212529;
                border: 1px solid #ced4da;
            }
            QMenu::item:selected {
                background-color: #e9ecef;
                color: #0d6efd;
            }
            QStatusBar {
                background-color: #ffffff;
                color: #6c757d;
                border-top: 1px solid #ced4da;
            }
            QTabWidget::pane {
                border: 1px solid #ced4da;
                border-radius: 6px;
            }
            QTabBar::tab {
                background: #e9ecef;
                color: #212529;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background: #0d6efd;
                color: #ffffff;
                font-weight: bold;
            }
        """


def apply_dialog_theme(dialog: QDialog, dark_theme: bool = True) -> None:
    """Explicitly applies the crystal-clear theme to dialogs."""
    dialog.setStyleSheet(get_app_stylesheet(dark_theme))


# ═══════════════════════════════════════════════════════════════
# QUOTA & KEY HELPERS
# ═══════════════════════════════════════════════════════════════

class QuotaKind(Enum):
    PER_MINUTE = 1
    PER_DAY = 2
    UNKNOWN = 3


def today_key() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def seconds_until_midnight() -> float:
    now = datetime.now()
    midnight = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return (midnight - now).total_seconds()


def parse_429(exc) -> Tuple[QuotaKind, float]:
    msg = str(exc).lower()
    if "per_day" in msg or "per day" in msg or "quota exceeded for quota metric 'generate content api requests per day'" in msg:
        return QuotaKind.PER_DAY, seconds_until_midnight()
    if "per_minute" in msg or "per minute" in msg or "rate limit" in msg:
        match = re.search(r"retry after (\d+)s", msg)
        if not match:
            match = re.search(r"try again in (\d+)s", msg)
        if match:
            secs = float(match.group(1)) + 5.0
            return QuotaKind.PER_MINUTE, secs
        return QuotaKind.PER_MINUTE, 30.0
    if "quota" in msg and "exceeded" in msg:
        if "day" in msg:
            return QuotaKind.PER_DAY, seconds_until_midnight()
        if "minute" in msg:
            return QuotaKind.PER_MINUTE, 30.0
        # إذا تجاوز حد الكوتة الفعلي الكبير (دون تحديد ثوانٍ أو دقيقة)، نعتبره 10 دقائق (600 ثانية)
        return QuotaKind.PER_MINUTE, 600.0
    return QuotaKind.UNKNOWN, 30.0


def is_quota_error(exc) -> bool:
    msg = str(exc).lower()
    if "429" in msg or "resource_exhausted" in msg or "quota" in msg:
        return True
    if hasattr(exc, "status_code") and getattr(exc, "status_code") == 429:
        return True
    return False


def is_developer_instruction_error(exc) -> bool:
    """Checks if error is caused by unsupported system_instruction on TTS models."""
    msg = str(exc).lower()
    return ("400" in msg or "invalid_argument" in msg) and (
        "developer instruction" in msg or "system_instruction" in msg or "system instruction" in msg or "not enabled for models" in msg or "not supported" in msg
    )


def is_retryable_error(exc) -> bool:
    msg = str(exc).lower()
    retryable_codes = ["500", "502", "503", "504", "timeout", "connection", "reset by peer", "unavailable", "overloaded"]
    return any(code in msg for code in retryable_codes)


def interruptible_sleep(seconds: float, stop_event: threading.Event) -> None:
    end = time.time() + seconds
    while time.time() < end:
        if stop_event.is_set():
            raise InterruptedError("توقف")
        remaining = end - time.time()
        if remaining <= 0:
            break
        step = min(0.3, remaining) if remaining > 2.0 else min(0.1, remaining)
        time.sleep(max(0.0, step))


def atomic_write_text(path: Path, content: str, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        with open(tmp, "w", encoding=encoding) as f:
            f.write(content)
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                pass
        for attempt in range(5):
            try:
                os.replace(tmp, path)
                return
            except PermissionError:
                if attempt < 4:
                    time.sleep(0.1 * (attempt + 1))
                else:
                    raise
    except Exception:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
        with open(path, "w", encoding=encoding) as f:
            f.write(content)


def play_or_open_file(filepath: str) -> bool:
    """Open or play audio/document file using system default application across OS platforms."""
    if not os.path.exists(filepath):
        return False
    try:
        if sys.platform == "win32":
            try:
                import winsound
                if filepath.lower().endswith(".wav"):
                    winsound.PlaySound(filepath, winsound.SND_FILENAME | winsound.SND_ASYNC)
                    return True
            except Exception:
                pass
            os.startfile(filepath)
            return True
        elif sys.platform == "darwin":
            subprocess.Popen(["open", filepath])
            return True
        else:
            subprocess.Popen(["xdg-open", filepath])
            return True
    except Exception:
        return False


def clean_arabic_text(text: str, auto_clean: bool = True, remove_diacritics: bool = False) -> str:
    """Normalizes and cleans Arabic text before chunking."""
    if not text:
        return ""
    t = text
    if remove_diacritics:
        t = re.sub(r'[\u064B-\u065F\u0670]', '', t)
    if auto_clean:
        t = re.sub(r'https?://\S+|www\.\S+', '', t)
        t = re.sub(r'[!！]{2,}', '!', t)
        t = re.sub(r'[?؟]{2,}', '؟', t)
        t = re.sub(r'[.…]\{2,}', '…', t)
        t = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', t)
        t = re.sub(r'[ \t]+', ' ', t)
        t = re.sub(r'\n\s*\n+', '\n\n', t)
    return t.strip()


# ═══════════════════════════════════════════════════════════════
# GlobalKeyManager — إدارة المفاتيح وتوزيع الحمل الذكي
# ═══════════════════════════════════════════════════════════════

class GlobalKeyManager:
    """Manages API keys with proactive RPM pacing, round-robin rotation, and 10-minute daily audio tracking."""

    def __init__(self, global_state_path: Path, rpm_limit_per_key: int = 14):
        self.path = global_state_path
        self._lock = threading.RLock()
        self._data: Dict[str, Dict[str, Dict[str, Any]]] = {}
        self._daily_audio_durations: Dict[str, Dict[str, Dict[str, float]]] = {}
        self._busy: Dict[str, Set[str]] = {}
        self._last_used: Dict[Tuple[str, str], float] = {}
        self._request_timestamps: Dict[Tuple[str, str], List[float]] = {}
        self._success_counts: Dict[str, int] = {}
        self._fail_counts: Dict[str, int] = {}
        self._rpm_limit = rpm_limit_per_key
        self.load()

    def load(self):
        with self._lock:
            self._data = {}
            self._daily_audio_durations = {}
            if not self.path.exists():
                return
            try:
                saved = json.loads(self.path.read_text(encoding="utf-8", errors="replace"))
                self._data = saved.get("cooldowns", {}) or {}
                self._daily_audio_durations = saved.get("daily_audio_durations", {}) or {}
                self._success_counts = saved.get("success_counts", {}) or {}
                self._fail_counts = saved.get("fail_counts", {}) or {}
                self._cleanup_locked()
            except Exception:
                pass

    def save(self):
        with self._lock:
            try:
                payload = {
                    "saved_at": datetime.now().isoformat(),
                    "cooldowns": self._data,
                    "daily_audio_durations": self._daily_audio_durations,
                    "success_counts": self._success_counts,
                    "fail_counts": self._fail_counts,
                }
                atomic_write_text(self.path, json.dumps(payload, ensure_ascii=False, indent=2))
            except Exception:
                pass

    def unblock_all_keys(self, api_key: Optional[str] = None) -> int:
        """Manually clears all cooldowns and daily audio limits across all keys or for a specific key."""
        with self._lock:
            count = 0
            if api_key:
                if api_key in self._data:
                    count += len(self._data[api_key])
                    self._data.pop(api_key, None)
                today = today_key()
                if today in self._daily_audio_durations and api_key in self._daily_audio_durations[today]:
                    self._daily_audio_durations[today].pop(api_key, None)
            else:
                for k in list(self._data.keys()):
                    count += len(self._data[k])
                self._data = {}
                self._daily_audio_durations = {}
            self.save()
            return count

    def get_daily_audio_seconds(self, api_key: str, model_id: str) -> float:
        with self._lock:
            today = today_key()
            return self._daily_audio_durations.get(today, {}).get(api_key, {}).get(model_id, 0.0)

    def mark_chunk_audio_success(self, api_key: str, model_id: str, audio_duration_sec: float):
        with self._lock:
            today = today_key()
            if today not in self._daily_audio_durations:
                self._daily_audio_durations = {today: {}}
            key_dict = self._daily_audio_durations[today].setdefault(api_key, {})
            current = key_dict.get(model_id, 0.0) + max(0.0, audio_duration_sec)
            key_dict[model_id] = current
            self._success_counts[api_key] = self._success_counts.get(api_key, 0) + 1
            
            # حظر المفتاح على هذا النموذج فقط عند الوصول إلى 10 دقائق (600 ثانية) من الصوت المولد اليوم
            if current >= 600.0:
                if api_key not in self._data:
                    self._data[api_key] = {}
                self._data[api_key][model_id] = {
                    "until": time.time() + seconds_until_midnight(),
                    "kind": "rpd",
                    "day": today,
                    "reason": f"اكتمل حد 10 دقائق صالحة اليوم على النموذج ({int(current/60)}د)"
                }
            self.save()

    def _cleanup_locked(self) -> int:
        count = 0
        now = time.time()
        for k in list(self._data.keys()):
            mc = self._data[k]
            for mid in list(mc.keys()):
                info = mc[mid]
                kind = info.get("kind")
                if kind == "rpd":
                    if info.get("day", "") != today_key():
                        mc.pop(mid, None)
                        count += 1
                else:
                    if now >= info.get("until", 0.0):
                        mc.pop(mid, None)
                        count += 1
            if not mc:
                self._data.pop(k, None)
                
        for key_pair in list(self._request_timestamps.keys()):
            self._request_timestamps[key_pair] = [
                ts for ts in self._request_timestamps[key_pair] if now - ts < 60.0
            ]
        return count

    def cleanup_and_save(self) -> int:
        with self._lock:
            n = self._cleanup_locked()
            if n > 0:
                self.save()
            return n

    def is_blocked_locked(self, api_key: str, model_id: str) -> bool:
        mc = self._data.get(api_key, {})
        info = mc.get(model_id)
        if not info:
            return False
        if info.get("kind") == "rpd":
            return info.get("day", "") == today_key()
        return time.time() < info.get("until", 0.0)

    def is_blocked(self, api_key: str, model_id: str) -> bool:
        with self._lock:
            self._cleanup_locked()
            return self.is_blocked_locked(api_key, model_id)

    def mark_blocked(self, api_key: str, model_id: str, kind: QuotaKind, delay: float):
        with self._lock:
            # 🌟 شرط الحظر الجديد على النموذج (10 دقائق أو حظر يومي صريح):
            today = today_key()
            daily_sec = self._daily_audio_durations.get(today, {}).get(api_key, {}).get(model_id, 0.0)

            if kind == QuotaKind.PER_DAY or daily_sec >= 590.0 or delay >= seconds_until_midnight() * 0.4:
                if api_key not in self._data:
                    self._data[api_key] = {}
                self._data[api_key][model_id] = {
                    "until": time.time() + seconds_until_midnight(),
                    "kind": "rpd",
                    "day": today,
                    "reason": f"حظر يومي على النموذج (أنتج اليوم {int(daily_sec)}s / 600s)"
                }
                self._fail_counts[api_key] = self._fail_counts.get(api_key, 0) + 1
                self.save()
                return

            if delay >= 600.0:
                if api_key not in self._data:
                    self._data[api_key] = {}
                self._data[api_key][model_id] = {
                    "until": time.time() + min(delay, 600.0),
                    "kind": "rpm",
                    "day": today,
                    "reason": f"تهدئة كوتة ({int(delay)}ث)"
                }
                self.save()
            else:
                # تهدئة قصيرة لا تتطلب وضع المفتاح في قائمة الحظر الطويلة، يستمر المفتاح مفتوحاً ومتاحاً للعمل
                pass
            self._fail_counts[api_key] = self._fail_counts.get(api_key, 0) + 1

    def mark_success(self, api_key: str):
        with self._lock:
            self._success_counts[api_key] = self._success_counts.get(api_key, 0) + 1

    def get_available_keys(self, all_keys: List[str], model_id: str) -> List[str]:
        with self._lock:
            self._cleanup_locked()
            avail = []
            for k in all_keys:
                if not self.is_blocked_locked(k, model_id):
                    avail.append(k)
            return avail

    def secs_until_any_available(self, all_keys: List[str], model_id: str) -> float:
        with self._lock:
            self._cleanup_locked()
            now = time.time()
            min_wait = float("inf")
            for k in all_keys:
                mc = self._data.get(k, {})
                info = mc.get(model_id)
                if not info:
                    return 0.0
                if info.get("kind") == "rpd":
                    if info.get("day", "") != today_key():
                        return 0.0
                    w = seconds_until_midnight()
                else:
                    w = max(0.0, info.get("until", 0.0) - now)
                if w < min_wait:
                    min_wait = w
            return min_wait if min_wait != float("inf") else 0.0

    def acquire_key(self, all_keys: List[str], model_id: str,
                    preferred_key: Optional[str] = None,
                    timeout: float = 300.0,
                    stop_event: Optional[threading.Event] = None) -> Optional[str]:
        """Acquire an available key for the given model with adaptive RPM pacing & round-robin."""
        end = time.time() + max(timeout, 0.0)
        while True:
            with self._lock:
                self._cleanup_locked()
                busy = self._busy.setdefault(model_id, set())
                candidates = [k for k in all_keys
                              if not self.is_blocked_locked(k, model_id) and k not in busy]
                if preferred_key and preferred_key in candidates:
                    busy.add(preferred_key)
                    self._request_timestamps.setdefault((model_id, preferred_key), []).append(time.time())
                    self._last_used[(model_id, preferred_key)] = time.time()
                    return preferred_key
                
                now = time.time()
                ready_candidates = []
                min_rpm_wait = float("inf")
                for k in candidates:
                    ts_list = self._request_timestamps.get((model_id, k), [])
                    if len(ts_list) < self._rpm_limit:
                        ready_candidates.append(k)
                    else:
                        oldest = ts_list[0]
                        w = max(0.1, 60.0 - (now - oldest))
                        if w < min_rpm_wait:
                            min_rpm_wait = w

                if ready_candidates:
                    ready_candidates.sort(key=lambda k: (
                        len(self._request_timestamps.get((model_id, k), [])),
                        self._last_used.get((model_id, k), 0.0)
                    ))
                    chosen = ready_candidates[0]
                    busy.add(chosen)
                    self._request_timestamps.setdefault((model_id, chosen), []).append(now)
                    self._last_used[(model_id, chosen)] = now
                    return chosen

            if stop_event and stop_event.is_set():
                return None
            if time.time() >= end:
                return None
            time.sleep(min(0.2, min_rpm_wait if min_rpm_wait != float("inf") else 0.4))

    def release_key(self, api_key: Optional[str], model_id: str):
        if not api_key:
            return
        with self._lock:
            self._busy.setdefault(model_id, set()).discard(api_key)
            self._last_used[(model_id, api_key)] = time.time()

    def get_key_status(self, api_key: str) -> Dict[str, Any]:
        with self._lock:
            self._cleanup_locked()
            mc = self._data.get(api_key, {})
            now = time.time()
            blocked_in = {}
            for mid, info in mc.items():
                reason = info.get("reason", "")
                if info.get("kind") == "rpd":
                    if info.get("day", "") == today_key():
                        blocked_in[mid] = f"محظور يومياً على هذا النموذج ({reason or 'حتى منتصف الليل'})"
                else:
                    rem = info.get("until", 0.0) - now
                    if rem > 0:
                        blocked_in[mid] = f"تهدئة ({reason or f'متبقي {int(rem)}ث'})"
            
            today = today_key()
            dur_dict = self._daily_audio_durations.get(today, {}).get(api_key, {})
            dur_summary = " | ".join(f"{mid}: {sec/60:.1f}د/10د" for mid, sec in dur_dict.items() if sec > 0)
            
            return {
                "key": api_key,
                "masked": self.mask_key(api_key),
                "blocked_models": blocked_in,
                "dur_summary": dur_summary or "0.0د اليوم",
                "success": self._success_counts.get(api_key, 0),
                "fail": self._fail_counts.get(api_key, 0),
                "is_fully_blocked": len(blocked_in) > 0,
            }

    def get_blocked_summary(self) -> str:
        with self._lock:
            self._cleanup_locked()
            if not self._data:
                return "⏸️ الحظر: لا يوجد مفاتيح محظورة حالياً"
            count = len(self._data)
            return f"⏸️ الحظر: {count} مفتاح محظور مؤقتاً (سيتم التبديل تلقائياً)"

    @staticmethod
    def mask_key(k: str) -> str:
        if not k or len(k) < 12:
            return "..."
        return f"{k[:6]}...{k[-4:]}"


# ═══════════════════════════════════════════════════════════════
# CheckpointManager — حفظ نقاط التوقف والاستئناف التلقائي
# ═══════════════════════════════════════════════════════════════

class CheckpointManager:
    """Manages project checkpoint state for seamless pause/resume."""

    def __init__(self, checkpoint_path: str):
        self._path = checkpoint_path
        self._data: Dict[str, Any] = {}
        self._lock = QMutex()
        self._load()

    def _load(self) -> None:
        if os.path.exists(self._path):
            try:
                with open(self._path, "r", encoding="utf-8") as f:
                    self._data = json.load(f)
            except Exception:
                self._data = {}

    def _save(self) -> None:
        try:
            os.makedirs(os.path.dirname(self._path), exist_ok=True)
            atomic_write_text(Path(self._path), json.dumps(self._data, ensure_ascii=False, indent=2))
        except Exception:
            pass

    def get_chunk_status(self, chunk_index: int) -> Optional[Dict]:
        with QMutexLocker(self._lock):
            return self._data.get(str(chunk_index))

    def mark_chunk(self, chunk_index: int, status: str, file_path: str = "",
                   chars: int = 0, time_taken: float = 0.0, error: str = "") -> None:
        with QMutexLocker(self._lock):
            self._data[str(chunk_index)] = {
                "status": status,
                "file": file_path,
                "chars": chars,
                "time": round(time_taken, 2),
                "error": error,
                "updated_at": datetime.now().isoformat(),
            }
            self._save()

    def get_completed_indices(self) -> List[int]:
        with QMutexLocker(self._lock):
            result = []
            for k, v in self._data.items():
                fpath = v.get("file", "")
                if v.get("status") == "done" and fpath and os.path.exists(fpath):
                    try:
                        if os.path.getsize(fpath) > 500:
                            with wave.open(fpath, "rb") as wf:
                                if wf.getnframes() > 0:
                                    result.append(int(k))
                    except Exception:
                        pass
            return sorted(result)

    def clear(self) -> None:
        with QMutexLocker(self._lock):
            self._data = {}
            if os.path.exists(self._path):
                try:
                    os.remove(self._path)
                except OSError:
                    pass


# ═══════════════════════════════════════════════════════════════
# ChunkManager — تقسيم وتفكيك النص الذكي (متعدد الطرق)
# ═══════════════════════════════════════════════════════════════

class ChunkManager:
    """Splits text into naturally paced chunks with smart boundary detection for Arabic."""
    SENTENCE_ENDINGS = re.compile(r'[.!?؟…\n]+')

    def __init__(self, text: str, split_method: str = "seconds",
                 chunk_duration: int = 60, chunk_words: int = 150, chunk_sentences: int = 10,
                 context_carry: bool = True, auto_clean: bool = True, remove_diacritics: bool = False):
        self._text = clean_arabic_text(text, auto_clean=auto_clean, remove_diacritics=remove_diacritics)
        self._split_method = split_method
        self._chunk_duration = chunk_duration
        self._chunk_words = chunk_words
        self._chunk_sentences = chunk_sentences
        self._context_carry = context_carry
        self._chunks: List[ChunkInfo] = []
        self._split()

    def _split_sentences(self, text: str) -> List[str]:
        parts = self.SENTENCE_ENDINGS.split(text)
        separators = self.SENTENCE_ENDINGS.findall(text)
        sentences = []
        for i, part in enumerate(parts):
            s = part.strip()
            if not s:
                continue
            if i < len(separators):
                s += separators[i].strip()
            sentences.append(s)
        return sentences

    def _split(self) -> None:
        sentences = self._split_sentences(self._text)
        if not sentences:
            if self._text.strip():
                self._chunks = [ChunkInfo(
                    index=0, text=self._text.strip(),
                    char_count=len(self._text.strip()),
                    estimated_duration=len(self._text.strip()) / CHARS_PER_SECOND)]
            return

        chunks = []
        current_sentences = []
        current_char_len = 0
        current_word_len = 0
        prev_last_two = []

        target_chars = min(self._chunk_duration * CHARS_PER_SECOND, MAX_CHUNK_CHARS)

        for sentence in sentences:
            sentence_len = len(sentence)
            sentence_words = len(sentence.split())
            
            if sentence_len > MAX_CHUNK_CHARS:
                if current_sentences:
                    chunk_text = " ".join(current_sentences)
                    prev_last_two = current_sentences[-2:] if len(current_sentences) >= 2 else list(current_sentences)
                    chunks.append(ChunkInfo(index=len(chunks), text=chunk_text,
                                            char_count=len(chunk_text),
                                            estimated_duration=len(chunk_text) / CHARS_PER_SECOND))
                    current_sentences = []
                    current_char_len = 0
                    current_word_len = 0
                words = sentence.split()
                sub = []
                sub_len = 0
                for w in words:
                    if sub_len + len(w) + 1 > MAX_CHUNK_CHARS and sub:
                        chunk_text = " ".join(sub)
                        chunks.append(ChunkInfo(index=len(chunks), text=chunk_text,
                                                char_count=len(chunk_text),
                                                estimated_duration=len(chunk_text) / CHARS_PER_SECOND))
                        sub = []
                        sub_len = 0
                    sub.append(w)
                    sub_len += len(w) + 1
                if sub:
                    current_sentences = sub
                    current_char_len = sub_len
                    current_word_len = len(sub)
                continue

            should_split = False
            if self._split_method == "sentences":
                if len(current_sentences) + 1 > self._chunk_sentences and current_sentences:
                    should_split = True
            elif self._split_method == "words":
                if current_word_len + sentence_words > self._chunk_words and current_sentences:
                    should_split = True
            else:  # seconds
                if current_char_len + sentence_len + 1 > target_chars and current_sentences:
                    should_split = True

            if current_char_len + sentence_len + 1 > MAX_CHUNK_CHARS and current_sentences:
                should_split = True

            if should_split:
                chunk_text = " ".join(current_sentences)
                context_hint = " ".join(prev_last_two) if self._context_carry and prev_last_two and chunks else ""
                prev_last_two = current_sentences[-2:] if len(current_sentences) >= 2 else list(current_sentences)
                chunks.append(ChunkInfo(index=len(chunks), text=chunk_text,
                                        char_count=len(chunk_text),
                                        estimated_duration=len(chunk_text) / CHARS_PER_SECOND,
                                        context_hint=context_hint))
                current_sentences = []
                current_char_len = 0
                current_word_len = 0

            current_sentences.append(sentence)
            current_char_len += sentence_len + 1
            current_word_len += sentence_words

        if current_sentences:
            chunk_text = " ".join(current_sentences)
            context_hint = " ".join(prev_last_two) if self._context_carry and prev_last_two and chunks else ""
            chunks.append(ChunkInfo(index=len(chunks), text=chunk_text,
                                    char_count=len(chunk_text),
                                    estimated_duration=len(chunk_text) / CHARS_PER_SECOND,
                                    context_hint=context_hint))
        self._chunks = chunks

    @property
    def chunks(self) -> List[ChunkInfo]:
        return self._chunks

    @property
    def total_chunks(self) -> int:
        return len(self._chunks)


# ═══════════════════════════════════════════════════════════════
# TTSWorkerThread — خيط المعالجة المتزامنة للمقاطع
# ═══════════════════════════════════════════════════════════════

class TTSWorkerThread(QThread):
    """Processes text chunks concurrently across available API keys using ThreadPoolExecutor."""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    chunk_started = pyqtSignal(int)
    chunk_done = pyqtSignal(int, str)
    chunk_failed = pyqtSignal(int, str)
    log_message = pyqtSignal(str, str)
    finished_signal = pyqtSignal(str)
    error = pyqtSignal(str)
    key_used = pyqtSignal(str)

    def __init__(self, chunks: List[ChunkInfo], model_id: str, voice: str,
                 speed: float, lang_hint: str, output_dir: str,
                 project_name: str, api_keys: List[str], max_workers: int,
                 retry_count: int, reading_style: str = "رواية (افتراضي)",
                 global_key_mgr: Optional[GlobalKeyManager] = None,
                 adaptive_parallel: bool = True, retry_failed_forever: bool = False):
        super().__init__()
        self._chunks = chunks
        self._model_id = model_id
        self._voice = voice
        self._speed = speed
        self._lang_hint = lang_hint
        self._output_dir = output_dir
        self._project_name = project_name
        self._api_keys = api_keys
        self._max_workers = min(max_workers, max(1, len(api_keys) * 2)) if api_keys else max_workers
        self._active_workers = self._max_workers
        self._retry_count = retry_count
        self._reading_style = reading_style
        self._global_key_mgr = global_key_mgr
        self._adaptive_parallel = adaptive_parallel
        self._retry_failed_forever = retry_failed_forever
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()
        self._total = len(chunks)
        self._done_count = 0
        self._consecutive_failures = 0
        self._checkpoint = CheckpointManager(
            os.path.join(output_dir, f"{project_name}_checkpoint.json"))

    def pause(self):
        self._pause_event.clear()

    def resume(self):
        self._pause_event.set()

    def stop(self):
        self._stop_event.set()
        self._pause_event.set()

    def is_paused(self) -> bool:
        return not self._pause_event.is_set()

    def _wait_if_paused(self) -> bool:
        while not self._pause_event.is_set():
            if self._stop_event.is_set():
                return True
            time.sleep(0.1)
        return self._stop_event.is_set()

    def _ts(self) -> str:
        return datetime.now().strftime("%H:%M:%S")

    def _process_single_chunk(self, chunk: ChunkInfo) -> Tuple[int, bool, str]:
        if self._stop_event.is_set():
            return (chunk.index, False, "stopped")

        cp = self._checkpoint.get_chunk_status(chunk.index)
        if cp and cp.get("status") == "done" and cp.get("file") and os.path.exists(cp["file"]):
            return (chunk.index, True, cp["file"])

        if self._wait_if_paused():
            return (chunk.index, False, "stopped")

        self.chunk_started.emit(chunk.index)
        output_path = os.path.join(self._output_dir, f"{self._project_name}_chunk_{chunk.index:04d}.wav")
        self._checkpoint.mark_chunk(chunk.index, "processing", chars=chunk.char_count)

        last_error = ""
        attempt = 0
        while True:
            if self._stop_event.is_set() or self._wait_if_paused():
                return (chunk.index, False, "stopped")

            api_key = self._global_key_mgr.acquire_key(
                self._api_keys, self._model_id,
                stop_event=self._stop_event, timeout=120)

            if api_key is None:
                last_error = "جميع المفاتيح محظورة أو مستنفدة حالياً"
                self.log_message.emit(f"[{self._ts()}] ⚠ لا مفتاح متاح — انتظار...", "#fab387")
                wait_secs = self._global_key_mgr.secs_until_any_available(self._api_keys, self._model_id)
                if wait_secs > 0:
                    self.log_message.emit(f"[{self._ts()}] ⏳ انتظار {wait_secs:.0f} ثانية حتى يتوفر مفتاح...", "#89b4fa")
                    try:
                        interruptible_sleep(min(wait_secs + 3, 120), self._stop_event)
                    except InterruptedError:
                        return (chunk.index, False, "stopped")
                if not self._retry_failed_forever and attempt >= self._retry_count:
                    break
                continue

            try:
                self.key_used.emit(GlobalKeyManager.mask_key(api_key))
                start_time = time.time()
                self._call_gemini_tts(chunk.text, api_key, output_path, chunk.context_hint)
                elapsed = time.time() - start_time

                self._global_key_mgr.release_key(api_key, self._model_id)
                audio_dur = chunk.estimated_duration
                try:
                    if os.path.exists(output_path):
                        with wave.open(output_path, "rb") as wf:
                            audio_dur = wf.getnframes() / wf.getframerate()
                except Exception:
                    pass
                self._global_key_mgr.mark_chunk_audio_success(api_key, self._model_id, audio_dur)
                self._checkpoint.mark_chunk(chunk.index, "done", output_path, chunk.char_count, elapsed)
                self._consecutive_failures = 0
                
                if self._adaptive_parallel and self._active_workers < self._max_workers:
                    self._active_workers = min(self._max_workers, self._active_workers + 1)
                    
                self.log_message.emit(f"[{self._ts()}] ✓ المقطع {chunk.index + 1} اكتمل بنجاح ({elapsed:.1f}ث)", "#a6e3a1")
                return (chunk.index, True, output_path)

            except Exception as e:
                attempt += 1
                last_error = str(e)
                self._global_key_mgr.release_key(api_key, self._model_id)

                if is_quota_error(e):
                    kind, delay = parse_429(e)
                    self._global_key_mgr.mark_blocked(api_key, self._model_id, kind, delay)
                    if kind == QuotaKind.PER_DAY or delay >= 600.0:
                        kind_label = "يومي" if kind == QuotaKind.PER_DAY else "10 دقائق (حد الكوتة)"
                        self.log_message.emit(f"[{self._ts()}] 🔄 حظر كوتة ({kind_label}) على {GlobalKeyManager.mask_key(api_key)} — التبديل لمفتاح آخر", "#fab387")
                    else:
                        self.log_message.emit(f"[{self._ts()}] ⚡ ضغط شبكة/معدل مؤقت على {GlobalKeyManager.mask_key(api_key)} — المفتاح مفتوح، تهدئة قصيرة...", "#89b4fa")
                    
                    if self._adaptive_parallel and self._active_workers > 1:
                        self._active_workers = max(1, self._active_workers // 2)
                        self.log_message.emit(f"[{self._ts()}] ⚡ توازي متكيف: تم خفض العمال إلى {self._active_workers}", "#fab387")
                        
                    time.sleep(min(3.0, delay if delay < 600.0 else 3.0))
                elif is_retryable_error(e):
                    backoff = min(60, 2 ** attempt)
                    self.log_message.emit(f"[{self._ts()}] ⚠ المقطع {chunk.index + 1} خطأ شبكة — محاولة {attempt} بعد {backoff}ث", "#fab387")
                    try:
                        interruptible_sleep(backoff, self._stop_event)
                    except InterruptedError:
                        return (chunk.index, False, "stopped")
                else:
                    self.log_message.emit(f"[{self._ts()}] ⚠ المقطع {chunk.index + 1} محاولة {attempt}/{self._retry_count}: {last_error[:80]}", "#fab387")
                    backoff = min(60, 2 ** attempt)
                    try:
                        interruptible_sleep(backoff, self._stop_event)
                    except InterruptedError:
                        return (chunk.index, False, "stopped")

                if not self._retry_failed_forever and attempt >= self._retry_count:
                    break

        self._checkpoint.mark_chunk(chunk.index, "failed", "", chunk.char_count, 0.0, last_error)
        self._consecutive_failures += 1
        self.log_message.emit(f"[{self._ts()}] ✗ المقطع {chunk.index + 1} فشل بعد المحاولات: {last_error[:80]}", "#f38ba8")
        return (chunk.index, False, last_error)

    def _call_gemini_tts(self, text: str, api_key: str, output_path: str, context_hint: str = "") -> None:
        if not GENAI_AVAILABLE:
            raise ImportError("مكتبة google-genai غير مثبتة في النظام")

        style_instruction = READING_STYLES.get(self._reading_style, READING_STYLES["رواية (افتراضي)"])
        if context_hint and context_hint.strip():
            style_instruction = f"{style_instruction}\n\n[توجيه سياقي عاطفي متصل من المقطع السابق لضبط وتيرة الإلقاء فقط، لا تقرأ هذا النص السابق بصوتك: \"{context_hint.strip()}\"]"

        client = genai.Client(api_key=api_key)
        speech_cfg = types.SpeechConfig(
            voice_config=types.VoiceConfig(
                prebuilt_voice_config=types.PrebuiltVoiceConfig(
                    voice_name=self._voice)))

        # نماذج الصوت المتخصصة مثل preview-tts لا تدعم system_instruction وتتطلب إرسال النص مع speech_config مباشرة
        use_sys_inst = not ("preview-tts" in self._model_id.lower() or "tts" in self._model_id.lower())
        response = None

        if use_sys_inst:
            try:
                response = client.models.generate_content(
                    model=self._model_id,
                    contents=text,
                    config=types.GenerateContentConfig(
                        system_instruction=style_instruction,
                        response_modalities=["AUDIO"],
                        speech_config=speech_cfg,
                    ),
                )
            except Exception as e:
                if is_developer_instruction_error(e):
                    use_sys_inst = False
                else:
                    raise

        if not use_sys_inst or response is None:
            response = client.models.generate_content(
                model=self._model_id,
                contents=text,
                config=types.GenerateContentConfig(
                    response_modalities=["AUDIO"],
                    speech_config=speech_cfg,
                ),
            )

        audio_data = response.candidates[0].content.parts[0].inline_data.data
        sample_rate = DEFAULT_SAMPLE_RATE
        try:
            mime = response.candidates[0].content.parts[0].inline_data.mime_type
            if "44100" in (mime or ""):
                sample_rate = 44100
        except (AttributeError, IndexError):
            pass
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with wave.open(output_path, "wb") as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(sample_rate)
            wf.writeframes(audio_data)

    def run(self) -> None:
        self.status.emit("جاري المعالجة المتزامنة...")
        self.log_message.emit(f"[{self._ts()}] 🚀 بدء المعالجة: {self._total} مقطع متصل (عمال متزامنين: {self._active_workers})", "#89b4fa")

        pending_chunks = []
        for chunk in self._chunks:
            cp = self._checkpoint.get_chunk_status(chunk.index)
            if cp and cp.get("status") == "done" and cp.get("file") and os.path.exists(cp["file"]):
                self._done_count += 1
                self.chunk_done.emit(chunk.index, cp["file"])
            else:
                pending_chunks.append(chunk)

        if self._done_count > 0:
            self.log_message.emit(f"[{self._ts()}] ♻ استئناف المشروع: {self._done_count} مقطع مكتمل مسبقاً", "#a6e3a1")
            self.progress.emit(int(self._done_count / self._total * 100))

        if not pending_chunks:
            self.finished_signal.emit("جميع المقاطع مكتملة مسبقاً وجاهزة!")
            return

        try:
            with ThreadPoolExecutor(max_workers=self._max_workers) as executor:
                futures = {executor.submit(self._process_single_chunk, c): c for c in pending_chunks}
                for future in as_completed(futures):
                    if self._stop_event.is_set():
                        executor.shutdown(wait=False, cancel_futures=True)
                        self.status.emit("تم الإيقاف")
                        self.log_message.emit(f"[{self._ts()}] ⏹ تم إيقاف المعالجة بناءً على أمر المستخدم", "#fab387")
                        return
                    idx, success, result = future.result()
                    if success:
                        self._done_count += 1
                        self.chunk_done.emit(idx, result)
                    else:
                        self.chunk_failed.emit(idx, result)
                    self.progress.emit(int(self._done_count / self._total * 100))
                    if self._consecutive_failures >= 3:
                        self.error.emit(UI["consecutive_fail"].format(self._consecutive_failures))
                        self._consecutive_failures = 0
        except Exception as e:
            self.error.emit(f"خطأ غير متوقع في المعالجة: {str(e)}")
            return

        self.progress.emit(100)
        done_count = len(self._checkpoint.get_completed_indices())
        msg = f"اكتملت المعالجة: {done_count}/{self._total} مقطع بنجاح"
        self.finished_signal.emit(msg)


# ═══════════════════════════════════════════════════════════════
# MergeWorker & AudioSplitter — خيط الدمج وتقسيم الملف النهائي
# ═══════════════════════════════════════════════════════════════

def split_wav_file(input_wav: str, method: str, parts: int = 5,
                   duration_min: int = 60, size_mb: int = 50,
                   overlap_sec: int = 10, target_dir: str = "") -> List[str]:
    """Splits master WAV file into parts with seamless cross-overlap inside project subfolder."""
    if not os.path.exists(input_wav) or method == "none":
        return [input_wav] if os.path.exists(input_wav) else []
        
    part_files = []
    with wave.open(input_wav, "rb") as win:
        params = win.getparams()
        total_frames = params.nframes
        framerate = params.framerate
        nchannels = params.nchannels
        sampwidth = params.sampwidth
        
        frames_per_part = total_frames
        if method == "parts" and parts > 0:
            frames_per_part = max(100, (total_frames + parts - 1) // parts)
        elif method == "duration" and duration_min > 0:
            frames_per_part = max(100, int(framerate * duration_min * 60))
        elif method == "size" and size_mb > 0:
            bytes_per_frame = max(1, nchannels * sampwidth)
            frames_per_part = max(100, int(size_mb * 1024 * 1024 / bytes_per_frame))
            
        overlap_frames = int(framerate * max(0, overlap_sec))
        part_idx = 0
        current_start = 0
        
        base_dir = target_dir if target_dir else os.path.dirname(input_wav)
        os.makedirs(base_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(input_wav))[0]
        
        while current_start < total_frames:
            part_idx += 1
            current_end = min(total_frames, current_start + frames_per_part)
            read_start = max(0, current_start - (overlap_frames if part_idx > 1 else 0))
            read_end = current_end
            
            win.setpos(read_start)
            frames_data = win.readframes(read_end - read_start)
            
            out_path = os.path.join(base_dir, f"{base_name}_part_{part_idx:02d}.wav")
            with wave.open(out_path, "wb") as wout:
                wout.setparams(params)
                wout.writeframes(frames_data)
            part_files.append(out_path)
            
            if current_end >= total_frames:
                break
            current_start = current_end
            
    return part_files


class MergeWorker(QThread):
    """Merges WAV audio files using Native WAV merger and optionally splits final master file."""
    progress = pyqtSignal(str)
    finished_signal = pyqtSignal(str, list)
    error = pyqtSignal(str)

    def __init__(self, wav_files: List[str], output_path: str,
                 project_name: str, output_dir: str, silence_pad_ms: int = 300,
                 use_ffmpeg: bool = False, split_method: str = "none",
                 split_parts: int = 5, split_minutes: int = 60, split_size_mb: int = 50,
                 split_overlap_sec: int = 10, keep_temp_files: bool = True,
                 audio_bitrate: str = "copy"):
        super().__init__()
        self._wav_files = [f for f in wav_files if os.path.exists(f)]
        self._output_path = output_path
        self._project_name = project_name
        self._output_dir = output_dir
        self._silence_pad_ms = max(0, silence_pad_ms)
        self._use_ffmpeg = use_ffmpeg
        self._split_method = split_method
        self._split_parts = split_parts
        self._split_minutes = split_minutes
        self._split_size_mb = split_size_mb
        self._split_overlap_sec = split_overlap_sec
        self._keep_temp_files = keep_temp_files
        self._audio_bitrate = audio_bitrate

    def _run_native_wav_merge(self) -> None:
        self.progress.emit("جاري دمج المقاطع باستخدام محرك الاستوديو المحلي (Native WAV)...")
        os.makedirs(os.path.dirname(self._output_path), exist_ok=True)
        
        first_params = None
        total_files = len(self._wav_files)
        
        with wave.open(self._output_path, "wb") as wout:
            for idx, fname in enumerate(self._wav_files):
                self.progress.emit(f"🔀 دمج المقطع {idx + 1}/{total_files}...")
                with wave.open(fname, "rb") as win:
                    params = win.getparams()
                    if first_params is None:
                        first_params = params
                        wout.setparams(first_params)
                    else:
                        if self._silence_pad_ms > 0:
                            silence_frames = int(first_params.framerate * self._silence_pad_ms / 1000.0)
                            if silence_frames > 0:
                                silence_bytes = b'\x00' * (silence_frames * first_params.sampwidth * first_params.nchannels)
                                wout.writeframes(silence_bytes)
                    wout.writeframes(win.readframes(win.getnframes()))

    def run(self) -> None:
        if not self._wav_files:
            self.error.emit("لا توجد ملفات صوتية صالحة للدمج")
            return
        
        try:
            if not self._use_ffmpeg or self._output_path.lower().endswith(".wav"):
                self._run_native_wav_merge()
            else:
                self.progress.emit("جاري إنشاء قائمة الدمج الخارجية...")
                concat_list_path = os.path.join(self._output_dir, f"{self._project_name}_concat_list.txt")
                try:
                    with open(concat_list_path, "w", encoding="utf-8") as f:
                        for wav in self._wav_files:
                            abs_path = os.path.abspath(wav).replace("'", "'\\''")
                            f.write(f"file '{abs_path}'\n")
                    self.progress.emit("جاري الدمج النهائي عبر ffmpeg...")
                    cmd = ["ffmpeg", "-y", "-f", "concat", "-safe", "0",
                           "-i", concat_list_path, "-c", "copy", self._output_path]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
                    if result.returncode != 0:
                        self.error.emit(f"خطأ ffmpeg: {result.stderr[:200]}")
                        return
                finally:
                    try:
                        if os.path.exists(concat_list_path):
                            os.remove(concat_list_path)
                    except OSError:
                        pass

            split_files = [self._output_path]
            if self._split_method != "none":
                self.progress.emit("جاري تقسيم الملف الصوتي المدمج داخل مجلد المقاطع الفرعي...")
                split_files = split_wav_file(
                    self._output_path, self._split_method,
                    parts=self._split_parts, duration_min=self._split_minutes,
                    size_mb=self._split_size_mb, overlap_sec=self._split_overlap_sec,
                    target_dir=self._output_dir
                )

            if not self._keep_temp_files:
                self.progress.emit("جاري تنظيف الملفات المؤقتة للمقاطع...")
                for f in self._wav_files:
                    try:
                        if os.path.exists(f) and f != self._output_path and f not in split_files:
                            os.remove(f)
                    except OSError:
                        pass

            self.finished_signal.emit(self._output_path, split_files)
        except Exception as e:
            self.error.emit(f"خطأ أثناء الدمج والتقسيم: {str(e)}")


# ═══════════════════════════════════════════════════════════════
# ModelFetchWorker — جلب النماذج من API
# ═══════════════════════════════════════════════════════════════

class ModelFetchWorker(QObject):
    """Worker to fetch available TTS models from Gemini API."""
    finished = pyqtSignal(list, list)
    error = pyqtSignal(str)

    def __init__(self, api_key: str):
        super().__init__()
        self._api_key = api_key

    def run(self):
        if not GENAI_AVAILABLE:
            self.error.emit("google-genai غير مثبتة")
            return
        try:
            client = genai.Client(api_key=self._api_key)
            models = client.models.list()
            tts_models = []
            recommended = []
            for m in models:
                mid = getattr(m, "name", "") or ""
                if mid.startswith("models/"):
                    mid = mid[7:]
                if "tts" in mid.lower():
                    label = getattr(m, "display_name", "") or mid
                    family = "2.5" if "2.5" in mid else ("2.0" if "2.0" in mid else ("1.5" if "1.5" in mid else "other"))
                    tts_models.append({"id": mid, "label": label, "family": family})
                    if "preview" in mid.lower() or "flash" in mid.lower() or "pro" in mid.lower():
                        recommended.append(mid)
            if not tts_models:
                tts_models = list(DEFAULT_MODELS)
                recommended = [m["id"] for m in tts_models]
            self.finished.emit(tts_models, recommended)
        except Exception as e:
            self.error.emit(str(e))


# ═══════════════════════════════════════════════════════════════
# DIALOGS — حوارات التخصيص والمراقبة والتقسيم
# ═══════════════════════════════════════════════════════════════

class ModelsSelectionDialog(QDialog):
    """Dialog to select and filter which TTS models are enabled."""

    def __init__(self, parent, all_models: List[Dict], enabled_ids: List[str]):
        super().__init__(parent)
        apply_dialog_theme(self, parent._config.dark_theme if hasattr(parent, '_config') else True)
        self.setWindowTitle("تخصيص نماذج Gemini TTS المتاحة")
        self.setMinimumWidth(520)
        self.setMinimumHeight(420)
        self._all_models = all_models
        self._enabled_ids = set(enabled_ids) if enabled_ids else {m["id"] for m in all_models}
        self._checkboxes: List[Tuple[QCheckBox, Dict]] = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        info_label = QLabel("حدد النماذج الصوتية التي تود إظهارها في قائمة الاختيار الرئيسية للتطبيق:")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        btn_row = QHBoxLayout()
        select_all_btn = QPushButton("تحديد الكل")
        select_all_btn.clicked.connect(self._select_all)
        btn_row.addWidget(select_all_btn)
        select_none_btn = QPushButton("إلغاء تحديد الكل")
        select_none_btn.clicked.connect(self._select_none)
        btn_row.addWidget(select_none_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout(content)

        grouped: Dict[str, List[Dict]] = {}
        for m in self._all_models:
            fam = m.get("family", "other")
            grouped.setdefault(fam, []).append(m)

        family_names = {"2.5": "عائلة Gemini 2.5 (الأحدث والأسرع)", "2.0": "عائلة Gemini 2.0", "1.5": "عائلة Gemini 1.5", "other": "نماذج أخرى"}
        for fam_key in sorted(grouped.keys(), reverse=True):
            group_box = QGroupBox(family_names.get(fam_key, f"عائلة {fam_key}"))
            group_layout = QVBoxLayout(group_box)
            for m in grouped[fam_key]:
                cb = QCheckBox(f"{m.get('label', m['id'])} ({m['id']})")
                cb.setChecked(m["id"] in self._enabled_ids)
                group_layout.addWidget(cb)
                self._checkboxes.append((cb, m))
            content_layout.addWidget(group_box)

        content_layout.addStretch()
        scroll.setWidget(content)
        layout.addWidget(scroll)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _select_all(self):
        for cb, _ in self._checkboxes:
            cb.setChecked(True)

    def _select_none(self):
        for cb, _ in self._checkboxes:
            cb.setChecked(False)

    def get_selected(self) -> List[Dict]:
        return [m for cb, m in self._checkboxes if cb.isChecked()]


class AddKeysDialog(QDialog):
    """Dialog to add new API keys."""

    def __init__(self, parent=None):
        super().__init__(parent)
        apply_dialog_theme(self, parent._config.dark_theme if hasattr(parent, '_config') else True)
        self.setWindowTitle("إضافة مفاتيح Gemini API جديدة")
        self.setMinimumWidth(540)
        self.setMinimumHeight(380)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        info = QLabel("أدخل مفاتيح API الخاصة بك (مفتاح واحد في كل سطر). سيتم حفظها بأمان في ملف tokens.txt وتوزيع الأحمال عليها تلقائياً.")
        info.setWordWrap(True)
        layout.addWidget(info)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText("AIzaSy...\nAIzaSy...")
        self.text_edit.setFont(QFont("Consolas", 11))
        layout.addWidget(self.text_edit)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["إضافة للمفاتيح الحالية (دمج)", "استبدال جميع المفاتيح الحالية (مسح القديم)"])
        layout.addWidget(self.mode_combo)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_keys(self) -> List[str]:
        raw = self.text_edit.toPlainText().strip()
        if not raw:
            return []
        keys = []
        for line in raw.splitlines():
            k = line.strip()
            if k and not k.startswith("#"):
                keys.append(k)
        return keys

    def is_replace(self) -> bool:
        return self.mode_combo.currentIndex() == 1


class KeysStatusDialog(QDialog):
    """Dialog showing status of all API keys."""

    def __init__(self, parent, all_keys: List[str], global_key_mgr: GlobalKeyManager):
        super().__init__(parent)
        apply_dialog_theme(self, parent._config.dark_theme if hasattr(parent, '_config') else True)
        self.setWindowTitle("📊 مراقبة وحالة مفاتيح Gemini API")
        self.setMinimumWidth(720)
        self.setMinimumHeight(440)
        self._all_keys = all_keys
        self._global_key_mgr = global_key_mgr
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        self.table = QTableWidget(len(self._all_keys), 5)
        self.table.setAlternatingRowColors(False)
        self.table.setHorizontalHeaderLabels(["المفتاح", "الحالة العامة", "تفاصيل الكوتة والحظر", "ناجح", "فاشل"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        for i, key in enumerate(self._all_keys):
            status_info = self._global_key_mgr.get_key_status(key)
            self.table.setItem(i, 0, QTableWidgetItem(status_info["masked"]))

            if status_info["is_fully_blocked"]:
                item_status = QTableWidgetItem("⏸️ محظور مؤقتاً")
                item_status.setForeground(QBrush(QColor("#fab387")))
            else:
                item_status = QTableWidgetItem("✅ نشط ومتاح")
                item_status.setForeground(QBrush(QColor("#a6e3a1")))
            self.table.setItem(i, 1, item_status)

            blocked_dict = status_info.get("blocked_models", {})
            if blocked_dict:
                desc = " | ".join(f"{mid}: {reason}" for mid, reason in blocked_dict.items())
            else:
                desc = "لا يوجد حظر فعال"
            self.table.setItem(i, 2, QTableWidgetItem(desc))
            self.table.setItem(i, 3, QTableWidgetItem(str(status_info["success"])))
            self.table.setItem(i, 4, QTableWidgetItem(str(status_info["fail"])))

        layout.addWidget(self.table)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)


class SettingsDialog(QDialog):
    """Application advanced preferences and v10.8.0 integrated features dialog."""

    def __init__(self, config: ProjectConfig, parent=None):
        super().__init__(parent)
        self.setWindowTitle(UI["settings_title"])
        self.setMinimumWidth(560)
        self._config = config
        apply_dialog_theme(self, config.dark_theme if config else True)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        
        # Tab 1: التزامن والمحاولات
        concurrency_tab = QWidget()
        c_form = QFormLayout(concurrency_tab)
        
        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(1, 20)
        self.workers_spin.setValue(self._config.max_workers)
        c_form.addRow(UI["max_workers"], self.workers_spin)

        self.retry_spin = QSpinBox()
        self.retry_spin.setRange(1, 15)
        self.retry_spin.setValue(self._config.retry_count)
        c_form.addRow(UI["retry_count"], self.retry_spin)
        
        self.rounds_spin = QSpinBox()
        self.rounds_spin.setRange(1, 30)
        self.rounds_spin.setValue(self._config.max_retry_rounds)
        self.rounds_spin.setSuffix(" جولات تلقائية")
        c_form.addRow("جولات إعادة المحاولة التلقائية للمقاطع الفاشلة:", self.rounds_spin)
        
        self.adaptive_check = QCheckBox("⚡ التكيف التلقائي مع ضغط الشبكة والكوتة (Adaptive Parallel Workers)")
        self.adaptive_check.setChecked(self._config.adaptive_parallel)
        c_form.addRow("التوازي الذكي:", self.adaptive_check)
        
        self.forever_check = QCheckBox("🔁 إعادة محاولة الأجزاء الفاشلة في حلقة مستمرة حتى النجاح (Infinite Retry)")
        self.forever_check.setChecked(self._config.retry_failed_forever)
        c_form.addRow("إصرار التحويل:", self.forever_check)
        
        self.shutdown_check = QCheckBox("💻 إيقاف تشغيل الكمبيوتر تلقائياً بعد انتهاء جميع المقاطع والدمج")
        self.shutdown_check.setChecked(self._config.shutdown_after_queue)
        c_form.addRow("بعد الانتهاء:", self.shutdown_check)
        
        tabs.addTab(concurrency_tab, "⚙️ التزامن والمحاولات")

        # Tab 2: الدمج وتقسيم المخرجات النهائي
        merge_tab = QWidget()
        m_form = QFormLayout(merge_tab)
        
        self.silence_spin = QSpinBox()
        self.silence_spin.setRange(0, 5000)
        self.silence_spin.setSingleStep(50)
        self.silence_spin.setValue(self._config.silence_pad_ms)
        self.silence_spin.setSuffix(" مللي ثانية")
        m_form.addRow("فاصل صمت بين المقاطع المدمجة:", self.silence_spin)
        
        self.keep_temp_check = QCheckBox("📂 الاحتفاظ بملفات المقاطع المنفصلة (chunk_xxxx.wav) بعد إتمام الدمج النهائي")
        self.keep_temp_check.setChecked(self._config.keep_temp_files)
        m_form.addRow("الملفات المؤقتة:", self.keep_temp_check)
        
        self.split_method_combo = QComboBox()
        self.split_method_combo.addItems([
            "بدون تقسيم — ملف واحد كامل (None)",
            "تقسيم إلى عدد محدد من الأجزاء (Parts)",
            "تقسيم كل مدة زمنية بالدقائق (Duration)",
            "تقسيم حسب حجم الملف بالميجابايت (Size)"
        ])
        if self._config.output_split_method == "parts":
            self.split_method_combo.setCurrentIndex(1)
        elif self._config.output_split_method == "duration":
            self.split_method_combo.setCurrentIndex(2)
        elif self._config.output_split_method == "size":
            self.split_method_combo.setCurrentIndex(3)
        else:
            self.split_method_combo.setCurrentIndex(0)
        m_form.addRow("تقسيم الملف المدمج:", self.split_method_combo)
        
        self.parts_spin = QSpinBox()
        self.parts_spin.setRange(2, 100)
        self.parts_spin.setValue(self._config.output_split_parts)
        self.parts_spin.setSuffix(" أجزاء")
        m_form.addRow("عدد الأجزاء الهدف:", self.parts_spin)
        
        self.minutes_spin = QSpinBox()
        self.minutes_spin.setRange(5, 600)
        self.minutes_spin.setValue(self._config.output_split_minutes)
        self.minutes_spin.setSuffix(" دقيقة لكل جزء")
        m_form.addRow("مدة كل جزء:", self.minutes_spin)
        
        self.size_spin = QSpinBox()
        self.size_spin.setRange(5, 2000)
        self.size_spin.setValue(self._config.output_split_size_mb)
        self.size_spin.setSuffix(" ميجابايت")
        m_form.addRow("حجم كل جزء:", self.size_spin)
        
        self.overlap_spin = QSpinBox()
        self.overlap_spin.setRange(0, 60)
        self.overlap_spin.setValue(self._config.output_split_overlap_sec)
        self.overlap_spin.setSuffix(" ثوانٍ")
        m_form.addRow("تداخل زمني بين الأجزاء المقسمة:", self.overlap_spin)
        
        tabs.addTab(merge_tab, "🎙️ الدمج وتقسيم المخرجات")

        # Tab 3: معالجة وتشكيل النص
        text_tab = QWidget()
        t_form = QFormLayout(text_tab)
        
        self.clean_check = QCheckBox(UI["auto_clean_symbols"])
        self.clean_check.setChecked(self._config.auto_clean_symbols)
        t_form.addRow("تنظيف النص:", self.clean_check)
        
        self.diacritics_check = QCheckBox(UI["remove_diacritics"])
        self.diacritics_check.setChecked(self._config.remove_diacritics)
        t_form.addRow("التشكيل والحركات:", self.diacritics_check)
        
        tabs.addTab(text_tab, "✨ معالجة النص العربي")

        # Tab 4: المظهر البصري
        ui_tab = QWidget()
        u_form = QFormLayout(ui_tab)
        
        self.theme_check = QCheckBox("تفعيل الوضع الداكن الليلي عالي الوضوح (Dark Theme)")
        self.theme_check.setChecked(self._config.dark_theme)
        self.theme_check.toggled.connect(self._on_live_theme_toggle)
        u_form.addRow("المظهر العام:", self.theme_check)
        
        tabs.addTab(ui_tab, "🎨 المظهر البصري")

        layout.addWidget(tabs)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_live_theme_toggle(self, checked: bool) -> None:
        if self.parent() and hasattr(self.parent(), "_config"):
            self.parent()._config.dark_theme = checked
            self.parent()._apply_theme()
        apply_dialog_theme(self, checked)

    def get_values(self) -> Dict[str, Any]:
        s_idx = self.split_method_combo.currentIndex()
        split_m = "none"
        if s_idx == 1: split_m = "parts"
        elif s_idx == 2: split_m = "duration"
        elif s_idx == 3: split_m = "size"
        
        return {
            "max_workers": self.workers_spin.value(),
            "retry_count": self.retry_spin.value(),
            "max_retry_rounds": self.rounds_spin.value(),
            "adaptive_parallel": self.adaptive_check.isChecked(),
            "retry_failed_forever": self.forever_check.isChecked(),
            "shutdown_after_queue": self.shutdown_check.isChecked(),
            "silence_pad_ms": self.silence_spin.value(),
            "keep_temp_files": self.keep_temp_check.isChecked(),
            "output_split_method": split_m,
            "output_split_parts": self.parts_spin.value(),
            "output_split_minutes": self.minutes_spin.value(),
            "output_split_size_mb": self.size_spin.value(),
            "output_split_overlap_sec": self.overlap_spin.value(),
            "auto_clean_symbols": self.clean_check.isChecked(),
            "remove_diacritics": self.diacritics_check.isChecked(),
            "dark_theme": self.theme_check.isChecked(),
        }


class TextPreviewDialog(QDialog):
    """Displays exact source text preview."""

    def __init__(self, text: str, parent=None):
        super().__init__(parent)
        apply_dialog_theme(self, parent._config.dark_theme if hasattr(parent, '_config') else True)
        self.setWindowTitle(UI["preview_text"])
        self.setMinimumSize(780, 560)
        layout = QVBoxLayout(self)
        text_edit = QPlainTextEdit()
        text_edit.setPlainText(text)
        text_edit.setReadOnly(True)
        text_edit.setFont(QFont("Tahoma", 12))
        layout.addWidget(text_edit)
        btn = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btn.rejected.connect(self.reject)
        layout.addWidget(btn)


class ChunkPreviewDialog(QDialog):
    """Table view of all segmented chunks prior to processing."""

    def __init__(self, chunks: List[ChunkInfo], parent=None):
        super().__init__(parent)
        apply_dialog_theme(self, parent._config.dark_theme if hasattr(parent, '_config') else True)
        self.setWindowTitle(UI["preview_chunks"])
        self.setMinimumSize(920, 600)
        layout = QVBoxLayout(self)
        table = QTableWidget(len(chunks), 4)
        table.setHorizontalHeaderLabels(["#", "نص المقطع الصافي", "عدد الأحرف", "المدة التقديرية"])
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        for i, c in enumerate(chunks):
            table.setItem(i, 0, QTableWidgetItem(str(c.index + 1)))
            preview_str = c.text[:140] + ("..." if len(c.text) > 140 else "")
            table.setItem(i, 1, QTableWidgetItem(preview_str))
            table.setItem(i, 2, QTableWidgetItem(str(c.char_count)))
            table.setItem(i, 3, QTableWidgetItem(f"{c.estimated_duration:.1f} ث"))

        layout.addWidget(table)
        btn = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btn.rejected.connect(self.reject)
        layout.addWidget(btn)


# ═══════════════════════════════════════════════════════════════
# FILE LOADERS & HELPER UTILITIES
# ═══════════════════════════════════════════════════════════════

def check_ffmpeg() -> bool:
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, timeout=5)
        return True
    except (FileNotFoundError, subprocess.TimeoutExpired, Exception):
        return False


def load_text_file(filepath: str) -> str:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".docx":
        if not DOCX_AVAILABLE:
            raise ImportError("مكتبة python-docx غير مثبتة قراءة ملفات docx.")
        doc = DocxDocument(filepath)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif ext == ".pdf":
        if not PDF_AVAILABLE:
            raise ImportError("مكتبة pypdf غير مثبتة لقراءة ملفات pdf.")
        reader = PdfReader(filepath)
        text_pages = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text_pages.append(t)
        return "\n".join(text_pages)
    else:
        for enc in ["utf-8", "utf-8-sig", "utf-16", "cp1256", "iso-8859-6", "latin-1"]:
            try:
                with open(filepath, "r", encoding=enc) as f:
                    return f.read()
            except (UnicodeDecodeError, Exception):
                continue
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            return f.read()


def count_words(text: str) -> int:
    return len(text.split())


def estimate_duration_seconds(char_count: int) -> float:
    return char_count / CHARS_PER_SECOND


# ═══════════════════════════════════════════════════════════════
# MAIN WINDOW — النافذة الرئيسية والاستوديو
# ═══════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    """Main application window for Gemini TTS Pro v3.8.13 Studio."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(UI["app_title"])
        self.setWindowIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))

        self._config = ProjectConfig()
        self._load_config()
        self._source_text: str = ""
        self._chunks: List[ChunkInfo] = []
        self._book_queue: List[BookJob] = []
        self._current_book_index: int = -1
        self._current_retry_round: int = 0
        self._worker: Optional[TTSWorkerThread] = None
        self._merge_worker: Optional[MergeWorker] = None
        self._fetch_worker: Optional[ModelFetchWorker] = None
        self._fetch_thread: Optional[QThread] = None
        self._closing = False
        self._ffmpeg_available = check_ffmpeg()

        self._global_key_mgr = GlobalKeyManager(CONFIG_DIR / GLOBAL_KEYS_STATE_FILE, rpm_limit_per_key=14)

        self._setup_ui()
        self._setup_menu()
        self._setup_shortcuts()
        self._apply_theme()
        self._startup_checks()

        self._periodic_timer = QTimer(self)
        self._periodic_timer.setInterval(45000)
        self._periodic_timer.timeout.connect(self._on_periodic_check)
        self._periodic_timer.start()

    def _load_config(self) -> None:
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for k, v in data.items():
                    if hasattr(self._config, k):
                        setattr(self._config, k, v)
            except Exception:
                pass

    def _save_config(self) -> None:
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            self._config.window_width = self.width()
            self._config.window_height = self.height()
            self._config.model_id = self._get_current_model_id()
            self._config.voice = self._get_current_voice()
            self._config.speed = self.speed_slider.value() / 10.0
            self._config.lang_hint = self.lang_input.text().strip()
            self._config.chunk_duration = self.chunk_spin.value()
            self._config.chunk_words = self.chunk_words_spin.value()
            self._config.chunk_sentences = self.chunk_sentences_spin.value()
            self._config.context_carry = self.context_check.isChecked()
            self._config.output_folder = self.output_path_input.text().strip()
            self._config.project_name = self.project_name_input.text().strip()
            self._config.auto_merge = self.auto_merge_check.isChecked()
            self._config.reading_style = self.reading_style_combo.currentText()

            s_idx = self.split_method_combo.currentIndex()
            if s_idx == 1: self._config.split_method = "words"
            elif s_idx == 2: self._config.split_method = "sentences"
            else: self._config.split_method = "seconds"

            b_idx = self.bitrate_combo.currentIndex()
            if b_idx == 1: self._config.audio_bitrate = "48k"
            elif b_idx == 2: self._config.audio_bitrate = "64k"
            elif b_idx == 3: self._config.audio_bitrate = "96k"
            elif b_idx == 4: self._config.audio_bitrate = "128k"
            else: self._config.audio_bitrate = "copy"

            atomic_write_text(CONFIG_FILE, json.dumps(asdict(self._config), ensure_ascii=False, indent=2))
        except Exception:
            pass

    def _get_all_keys(self) -> List[str]:
        tokens_file = CONFIG_DIR / TOKENS_FILE_NAME
        if not tokens_file.exists():
            return []
        try:
            raw = tokens_file.read_text(encoding="utf-8").strip()
            keys = [line.strip() for line in raw.splitlines() if line.strip() and not line.strip().startswith("#")]
            return keys
        except Exception:
            return []

    def _save_keys(self, keys: List[str]) -> None:
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            tokens_file = CONFIG_DIR / TOKENS_FILE_NAME
            atomic_write_text(tokens_file, "\n".join(keys) + "\n")
        except Exception:
            pass

    def _setup_ui(self) -> None:
        self.setMinimumSize(1240, 780)
        self.resize(self._config.window_width, self._config.window_height)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # ── LEFT PANEL (مقسم إلى تبويبين لسهولة المتابعة والترتيب) ──
        self.left_tabs = QTabWidget()
        
        # ── TAB 1: التحكم العام والإلقاء والتقسيم ──
        tab1_scroll = QScrollArea()
        tab1_scroll.setWidgetResizable(True)
        tab1_scroll.setFrameShape(QFrame.Shape.NoFrame)
        tab1_widget = QWidget()
        tab1_layout = QVBoxLayout(tab1_widget)
        tab1_layout.setSpacing(10)

        # 1. Source Text Group
        src_group = QGroupBox(UI["source_text"])
        src_layout = QVBoxLayout(src_group)
        
        top_src_row = QHBoxLayout()
        self.new_project_btn = QPushButton("✨ مشروع / تحويل جديد")
        self.new_project_btn.setStyleSheet("background-color: #89b4fa; color: #11111b; font-weight: bold;")
        self.new_project_btn.clicked.connect(self._new_project_action)
        top_src_row.addWidget(self.new_project_btn)
        top_src_row.addStretch()
        src_layout.addLayout(top_src_row)
        
        file_row = QHBoxLayout()
        self.file_path_label = QLineEdit()
        self.file_path_label.setReadOnly(True)
        self.file_path_label.setPlaceholderText(UI["file_path"])
        if self._config.source_file:
            self.file_path_label.setText(self._config.source_file)
        file_row.addWidget(self.file_path_label)
        self.browse_btn = QPushButton(UI["browse"])
        self.browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self.browse_btn)
        src_layout.addLayout(file_row)
        
        self.text_stats_label = QLabel(UI["text_stats"])
        self.text_stats_label.setWordWrap(True)
        self.text_stats_label.setStyleSheet("font-size: 12px; font-weight: 500;")
        src_layout.addWidget(self.text_stats_label)
        
        self.preview_text_btn = QPushButton(UI["preview_text"])
        self.preview_text_btn.clicked.connect(self._preview_text)
        self.preview_text_btn.setEnabled(False)
        src_layout.addWidget(self.preview_text_btn)
        tab1_layout.addWidget(src_group)

        # 2. TTS Settings Group
        tts_group = QGroupBox(UI["tts_settings"])
        tts_layout = QFormLayout(tts_group)
        tts_layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)

        voice_row = QHBoxLayout()
        self.voice_combo = QComboBox()
        self._populate_voice_combo()
        voice_row.addWidget(self.voice_combo, 1)
        self.preview_voice_btn = QPushButton(UI["preview_voice"])
        self.preview_voice_btn.setMaximumWidth(110)
        self.preview_voice_btn.clicked.connect(self._preview_voice)
        voice_row.addWidget(self.preview_voice_btn)
        tts_layout.addRow(UI["voice"], voice_row)

        speed_row = QHBoxLayout()
        self.speed_slider = QSlider(Qt.Orientation.Horizontal)
        self.speed_slider.setRange(5, 20)
        self.speed_slider.setValue(int(self._config.speed * 10))
        self.speed_label = QLabel(f"{self._config.speed:.1f}x")
        self.speed_label.setMinimumWidth(35)
        self.speed_slider.valueChanged.connect(lambda v: self.speed_label.setText(f"{v / 10:.1f}x"))
        speed_row.addWidget(self.speed_slider)
        speed_row.addWidget(self.speed_label)
        tts_layout.addRow(UI["speed"], speed_row)

        self.lang_input = QLineEdit(self._config.lang_hint)
        self.lang_input.setMaximumWidth(90)
        tts_layout.addRow(UI["lang_hint"], self.lang_input)

        self.reading_style_combo = QComboBox()
        self.reading_style_combo.addItems(list(READING_STYLES.keys()))
        if self._config.reading_style in READING_STYLES:
            self.reading_style_combo.setCurrentText(self._config.reading_style)
        tts_layout.addRow("أسلوب القراءة:", self.reading_style_combo)
        tab1_layout.addWidget(tts_group)

        # 3. API Keys & Quota Group (أعلى تبويب التقسيم)
        key_group = QGroupBox(UI["api_keys"])
        key_layout = QVBoxLayout(key_group)
        self.keys_label = QLabel("🔑 إجمالي: 0 | نشطة: 0")
        self.keys_label.setStyleSheet("font-weight: bold; font-size: 13px;")
        key_layout.addWidget(self.keys_label)

        self.blocked_label = QLabel("⏸️ الحظر: لا يوجد")
        self.blocked_label.setStyleSheet("font-size: 11px;")
        self.blocked_label.setWordWrap(True)
        key_layout.addWidget(self.blocked_label)

        key_btn_row = QHBoxLayout()
        self.add_keys_btn = QPushButton(UI["add_keys"])
        self.add_keys_btn.clicked.connect(self._on_add_keys)
        key_btn_row.addWidget(self.add_keys_btn)

        self.validate_keys_btn = QPushButton(UI["validate_keys"])
        self.validate_keys_btn.clicked.connect(self._validate_keys)
        key_btn_row.addWidget(self.validate_keys_btn)

        self.keys_status_btn = QPushButton("📊 التفاصيل")
        self.keys_status_btn.clicked.connect(self._show_keys_status)
        key_btn_row.addWidget(self.keys_status_btn)
        key_layout.addLayout(key_btn_row)

        self.unblock_all_btn = QPushButton("🔓 فك الحظر اليدوي عن كل المفاتيح")
        self.unblock_all_btn.setStyleSheet("background-color: #3b3d54; color: #89b4fa; font-weight: bold;")
        self.unblock_all_btn.clicked.connect(self._unblock_all_keys_action)
        key_layout.addWidget(self.unblock_all_btn)
        tab1_layout.addWidget(key_group)

        # 4. Chunking & Text Engineering Group
        chunk_group = QGroupBox(UI["chunking"])
        chunk_layout = QVBoxLayout(chunk_group)
        
        split_m_row = QHBoxLayout()
        split_m_row.addWidget(QLabel(UI["split_method_label"]))
        self.split_method_combo = QComboBox()
        self.split_method_combo.addItems(["بالمدة الزمنية (Seconds)", "بعد الكلمات (Words)", "بعد الجمل (Sentences)"])
        if self._config.split_method == "words":
            self.split_method_combo.setCurrentIndex(1)
        elif self._config.split_method == "sentences":
            self.split_method_combo.setCurrentIndex(2)
        else:
            self.split_method_combo.setCurrentIndex(0)
        self.split_method_combo.currentIndexChanged.connect(self._on_split_method_changed)
        split_m_row.addWidget(self.split_method_combo, 1)
        chunk_layout.addLayout(split_m_row)

        self.stack_chunk_inputs = QStackedWidget()
        
        w_sec = QWidget()
        l_sec = QHBoxLayout(w_sec)
        l_sec.setContentsMargins(0, 0, 0, 0)
        l_sec.addWidget(QLabel(UI["chunk_duration"]))
        self.chunk_spin = QSpinBox()
        self.chunk_spin.setRange(20, 600)
        self.chunk_spin.setValue(self._config.chunk_duration)
        self.chunk_spin.setSuffix(" ث")
        self.chunk_spin.setSingleStep(10)
        self.chunk_spin.valueChanged.connect(self._update_chunk_estimate)
        l_sec.addWidget(self.chunk_spin)
        self.stack_chunk_inputs.addWidget(w_sec)

        w_words = QWidget()
        l_words = QHBoxLayout(w_words)
        l_words.setContentsMargins(0, 0, 0, 0)
        l_words.addWidget(QLabel(UI["chunk_words_label"]))
        self.chunk_words_spin = QSpinBox()
        self.chunk_words_spin.setRange(20, 500)
        self.chunk_words_spin.setValue(self._config.chunk_words)
        self.chunk_words_spin.setSuffix(" كلمة")
        self.chunk_words_spin.setSingleStep(20)
        self.chunk_words_spin.valueChanged.connect(self._update_chunk_estimate)
        l_words.addWidget(self.chunk_words_spin)
        self.stack_chunk_inputs.addWidget(w_words)

        w_sent = QWidget()
        l_sent = QHBoxLayout(w_sent)
        l_sent.setContentsMargins(0, 0, 0, 0)
        l_sent.addWidget(QLabel(UI["chunk_sentences_label"]))
        self.chunk_sentences_spin = QSpinBox()
        self.chunk_sentences_spin.setRange(2, 50)
        self.chunk_sentences_spin.setValue(self._config.chunk_sentences)
        self.chunk_sentences_spin.setSuffix(" جمل")
        self.chunk_sentences_spin.setSingleStep(2)
        self.chunk_sentences_spin.valueChanged.connect(self._update_chunk_estimate)
        l_sent.addWidget(self.chunk_sentences_spin)
        self.stack_chunk_inputs.addWidget(w_sent)

        self.stack_chunk_inputs.setCurrentIndex(self.split_method_combo.currentIndex())
        chunk_layout.addWidget(self.stack_chunk_inputs)

        self.chunk_est_label = QLabel()
        self._update_chunk_estimate()
        chunk_layout.addWidget(self.chunk_est_label)
        
        self.preview_chunks_btn = QPushButton(UI["preview_chunks"])
        self.preview_chunks_btn.clicked.connect(self._preview_chunks)
        self.preview_chunks_btn.setEnabled(False)
        chunk_layout.addWidget(self.preview_chunks_btn)
        
        self.context_check = QCheckBox(UI["context_carry"])
        self.context_check.setChecked(self._config.context_carry)
        chunk_layout.addWidget(self.context_check)
        tab1_layout.addWidget(chunk_group)
        tab1_layout.addStretch()
        tab1_scroll.setWidget(tab1_widget)
        self.left_tabs.addTab(tab1_scroll, "🎙️ التحكم والإلقاء والتقسيم")

        # ── TAB 2: المخرجات والنماذج وإعادة الترميز ──
        tab2_scroll = QScrollArea()
        tab2_scroll.setWidgetResizable(True)
        tab2_scroll.setFrameShape(QFrame.Shape.NoFrame)
        tab2_widget = QWidget()
        tab2_layout = QVBoxLayout(tab2_widget)
        tab2_layout.setSpacing(10)

        # 1. Models Management Group
        models_group = QGroupBox(UI["models_management"])
        models_layout = QVBoxLayout(models_group)
        model_row = QHBoxLayout()
        model_row.addWidget(QLabel(UI["model"]))
        self.model_combo = QComboBox()
        self._populate_model_combo()
        model_row.addWidget(self.model_combo, 1)
        models_layout.addLayout(model_row)

        model_btn_row = QHBoxLayout()
        self.fetch_models_btn = QPushButton(UI["fetch_models"])
        self.fetch_models_btn.clicked.connect(self._on_fetch_models)
        model_btn_row.addWidget(self.fetch_models_btn)
        self.select_models_btn = QPushButton(UI["select_models"])
        self.select_models_btn.clicked.connect(self._on_select_models)
        model_btn_row.addWidget(self.select_models_btn)
        models_layout.addLayout(model_btn_row)

        self.models_status_label = QLabel(UI["models_status"])
        self.models_status_label.setWordWrap(True)
        self.models_status_label.setStyleSheet("font-size: 11px;")
        models_layout.addWidget(self.models_status_label)
        tab2_layout.addWidget(models_group)

        # 2. Output & Project Group
        out_group = QGroupBox(UI["output"])
        out_layout = QVBoxLayout(out_group)
        out_row = QHBoxLayout()
        self.output_path_input = QLineEdit(self._config.output_folder)
        self.output_path_input.setPlaceholderText(UI["output_folder"])
        out_row.addWidget(self.output_path_input)
        self.browse_output_btn = QPushButton(UI["browse"])
        self.browse_output_btn.clicked.connect(self._browse_output)
        out_row.addWidget(self.browse_output_btn)
        out_layout.addLayout(out_row)
        
        proj_row = QHBoxLayout()
        proj_row.addWidget(QLabel(UI["project_name"]))
        self.project_name_input = QLineEdit(self._config.project_name)
        proj_row.addWidget(self.project_name_input)
        out_layout.addLayout(proj_row)
        
        folders_row = QHBoxLayout()
        btn_open_root_left = QPushButton(UI["open_root_folder"])
        btn_open_root_left.clicked.connect(self._open_root_output_folder)
        folders_row.addWidget(btn_open_root_left)
        btn_open_book_left = QPushButton(UI["open_folder"])
        btn_open_book_left.clicked.connect(self._open_output_folder)
        folders_row.addWidget(btn_open_book_left)
        out_layout.addLayout(folders_row)

        self.auto_merge_check = QCheckBox(UI["auto_merge"])
        self.auto_merge_check.setChecked(self._config.auto_merge)
        out_layout.addWidget(self.auto_merge_check)
        tab2_layout.addWidget(out_group)

        # 3. Audio Re-encoding & Bitrate Group
        reencode_group = QGroupBox("🎛️ ضغط وإعادة ترميز الصوت المدمج (Audio Re-encoding)")
        reencode_layout = QVBoxLayout(reencode_group)
        reencode_info = QLabel("اختر دقة ومعدل ضغط الملف النهائي المدمج لتقليل حجم الكتب الطويلة (يتطلب ffmpeg للتحويل إلى MP3):")
        reencode_info.setWordWrap(True)
        reencode_info.setStyleSheet("font-size: 11px;")
        reencode_layout.addWidget(reencode_info)

        bitrate_row = QHBoxLayout()
        bitrate_row.addWidget(QLabel(UI["audio_bitrate_label"]))
        self.bitrate_combo = QComboBox()
        self.bitrate_combo.addItems([
            "بدون إعادة ترميز — الدقة الأصلية المباشرة (WAV - الافتراضي)",
            "48 kbps (حجم صغير جداً - مناسب للكلام)",
            "64 kbps (حجم اقتصادي ممتاز للكتب - متوازن)",
            "96 kbps (جودة عالية وحجم مناسب)",
            "128 kbps (جودة صوت استوديو فائقة)"
        ])
        if getattr(self._config, 'audio_bitrate', 'copy') == "48k": self.bitrate_combo.setCurrentIndex(1)
        elif getattr(self._config, 'audio_bitrate', 'copy') == "64k": self.bitrate_combo.setCurrentIndex(2)
        elif getattr(self._config, 'audio_bitrate', 'copy') == "96k": self.bitrate_combo.setCurrentIndex(3)
        elif getattr(self._config, 'audio_bitrate', 'copy') == "128k": self.bitrate_combo.setCurrentIndex(4)
        else: self.bitrate_combo.setCurrentIndex(0)
        bitrate_row.addWidget(self.bitrate_combo, 1)
        reencode_layout.addLayout(bitrate_row)
        tab2_layout.addWidget(reencode_group)

        tab2_layout.addStretch()
        tab2_scroll.setWidget(tab2_widget)
        self.left_tabs.addTab(tab2_scroll, "📁 المخرجات والنماذج والترميز")

        # ── RIGHT PANEL (لوحة التشغيل والمتابعة الحية) ──
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(8)

        # Control bar
        ctrl_bar = QHBoxLayout()
        self.start_btn = QPushButton(UI["start"])
        self.start_btn.setObjectName("btn_start")
        self.start_btn.clicked.connect(self._start_processing)
        ctrl_bar.addWidget(self.start_btn, 2)
        
        self.pause_btn = QPushButton(UI["pause"])
        self.pause_btn.setObjectName("btn_pause")
        self.pause_btn.setEnabled(False)
        self.pause_btn.clicked.connect(self._toggle_pause)
        ctrl_bar.addWidget(self.pause_btn, 1)
        
        self.stop_btn = QPushButton(UI["stop"])
        self.stop_btn.setObjectName("btn_stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop_processing)
        ctrl_bar.addWidget(self.stop_btn, 1)
        
        self.merge_btn = QPushButton(UI["merge_available"])
        self.merge_btn.clicked.connect(self._merge_available)
        self.merge_btn.setEnabled(False)
        ctrl_bar.addWidget(self.merge_btn, 1)
        
        self.retry_all_btn = QPushButton(UI["retry_failed_all"])
        self.retry_all_btn.clicked.connect(self._retry_all_failed_chunks)
        self.retry_all_btn.setEnabled(False)
        ctrl_bar.addWidget(self.retry_all_btn, 1)

        self.open_folder_btn = QPushButton(UI["open_folder"])
        self.open_folder_btn.clicked.connect(self._open_output_folder)
        ctrl_bar.addWidget(self.open_folder_btn, 1)

        self.open_root_btn = QPushButton(UI["open_root_folder"])
        self.open_root_btn.clicked.connect(self._open_root_output_folder)
        ctrl_bar.addWidget(self.open_root_btn, 1)
        
        right_layout.addLayout(ctrl_bar)

        # Stats bar
        stats_row = QHBoxLayout()
        self.stat_keys = QLabel("🔑 0/0")
        stats_row.addWidget(self.stat_keys)
        self.stat_eta = QLabel("⏳ ETA: —")
        stats_row.addWidget(self.stat_eta)
        self.audio_dur_label = QLabel("🔊 الصوت المولد: —")
        stats_row.addWidget(self.audio_dur_label)
        stats_row.addStretch()
        right_layout.addLayout(stats_row)

        # Progress
        progress_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        progress_row.addWidget(self.progress_bar)
        right_layout.addLayout(progress_row)

        self.chunk_counter_label = QLabel("📊 المنجز من المقاطع: 0 / 0 مقطع (0.0%)")
        self.chunk_counter_label.setStyleSheet("font-size: 14px; font-weight: 800; color: #a6e3a1; background-color: #313244; padding: 6px 12px; border: 1px solid #89b4fa; border-radius: 6px;")
        right_layout.addWidget(self.chunk_counter_label)

        self.current_chunk_label = QLabel(UI["current_chunk"])
        self.current_chunk_label.setStyleSheet("font-weight: 600;")
        right_layout.addWidget(self.current_chunk_label)

        # Chunks table (7 columns: +retry button)
        self.chunks_table = QTableWidget(0, 7)
        self.chunks_table.setHorizontalHeaderLabels(["#", "الحالة", "الأحرف", "المدة", "الملف الصوتي", "الزمن", "إجراء"])
        self.chunks_table.horizontalHeader().setStretchLastSection(False)
        self.chunks_table.setAlternatingRowColors(False)
        self.chunks_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.chunks_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.chunks_table.verticalHeader().setVisible(False)
        h = self.chunks_table.horizontalHeader()
        for col in [0, 1, 2, 3, 5, 6]:
            h.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.chunks_table.cellDoubleClicked.connect(self._on_table_double_click)
        self.chunks_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.chunks_table.customContextMenuRequested.connect(self._show_table_context_menu)
        right_layout.addWidget(self.chunks_table)

        # Log
        log_header = QHBoxLayout()
        log_header.addWidget(QLabel(UI["log_panel"]))
        log_header.addStretch()
        self.export_log_btn = QPushButton(UI["export_log"])
        self.export_log_btn.setMaximumWidth(130)
        self.export_log_btn.clicked.connect(self._export_log)
        log_header.addWidget(self.export_log_btn)
        right_layout.addLayout(log_header)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMaximumHeight(170)
        self.log_edit.setFont(QFont("Consolas", 10))
        right_layout.addWidget(self.log_edit)

        # Splitter sizes
        splitter.addWidget(self.left_tabs)
        splitter.addWidget(right_widget)
        splitter.setSizes([440, 800])

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar_label = QLabel(UI["status_ready"])
        self.status_bar.addPermanentWidget(self.status_bar_label)

        # ETA timer
        self._eta_timer = QTimer(self)
        self._eta_timer.setInterval(2000)
        self._eta_timer.timeout.connect(self._update_eta)

    def _on_split_method_changed(self, idx: int) -> None:
        self.stack_chunk_inputs.setCurrentIndex(idx)
        self._update_chunk_estimate()

    def _setup_menu(self) -> None:
        menubar = self.menuBar()
        file_menu = menubar.addMenu("ملف")
        new_proj_action = QAction("✨ تحويل / مشروع جديد (New Project)", self)
        new_proj_action.setShortcut(QKeySequence("Ctrl+N"))
        new_proj_action.triggered.connect(self._new_project_action)
        file_menu.addAction(new_proj_action)
        
        open_action = QAction(UI["open_file"], self)
        open_action.setShortcut(QKeySequence("Ctrl+O"))
        open_action.triggered.connect(self._browse_file)
        file_menu.addAction(open_action)
        self.recent_menu = file_menu.addMenu(UI["recent_files"])
        self._load_recent_files_menu()
        file_menu.addSeparator()
        exit_action = QAction("خروج", self)
        exit_action.setShortcut(QKeySequence("Ctrl+Q"))
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        view_menu = menubar.addMenu("عرض")
        self.theme_action = QAction("☀️ إضاءة نهارية" if self._config.dark_theme else "🌙 إضاءة ليلية داكنة", self)
        self.theme_action.triggered.connect(self._toggle_theme)
        view_menu.addAction(self.theme_action)

        tools_menu = menubar.addMenu("أدوات")
        settings_action = QAction(UI["settings"], self)
        settings_action.triggered.connect(self._open_settings)
        tools_menu.addAction(settings_action)

        help_menu = menubar.addMenu("مساعدة")
        about_action = QAction(UI["about"], self)
        about_action.triggered.connect(lambda: QMessageBox.about(self, "حول التطبيق", UI["about_text"]))
        help_menu.addAction(about_action)

    def _setup_shortcuts(self) -> None:
        QShortcut(QKeySequence("Ctrl+Return"), self).activated.connect(self._start_processing)
        QShortcut(QKeySequence("Ctrl+Space"), self).activated.connect(self._toggle_pause)

    def _apply_theme(self) -> None:
        ss = get_app_stylesheet(self._config.dark_theme)
        self.setStyleSheet(ss)
        if QApplication.instance():
            QApplication.instance().setStyleSheet(ss)

    def _toggle_theme(self) -> None:
        self._config.dark_theme = not self._config.dark_theme
        self.theme_action.setText("☀️ إضاءة نهارية" if self._config.dark_theme else "🌙 إضاءة ليلية داكنة")
        self._apply_theme()
        self._save_config()

    def _startup_checks(self) -> None:
        if not self._ffmpeg_available:
            self._log("ℹ️ محرك ffmpeg غير مثبت في النظام — سيتم استخدام محرك الدمج المدمج فائق السرعة (Pure Python WAV Merger).", "#89b4fa")
        if not GENAI_AVAILABLE:
            self._log("❌ تنبيه: مكتبة google-genai غير مثبتة. الرجاء تشغيل: pip install google-genai", "#f38ba8")
        self._update_keys_label()
        self._update_models_status_label()
        if self._config.source_file and os.path.exists(self._config.source_file):
            self._load_file(self._config.source_file)

    def _on_periodic_check(self):
        self._global_key_mgr.cleanup_and_save()
        self._update_keys_label()
        self._update_models_status_label()

    def _populate_model_combo(self) -> None:
        self.model_combo.clear()
        models = self._config.all_fetched_models if self._config.all_fetched_models else DEFAULT_MODELS
        enabled_set = set(self._config.enabled_model_ids) if self._config.enabled_model_ids else {m["id"] for m in models}
        active_models = [m for m in models if m["id"] in enabled_set]
        if not active_models:
            active_models = list(DEFAULT_MODELS)
        for m in active_models:
            self.model_combo.addItem(f"{m.get('label', m['id'])} ({m['id']})", m["id"])
        if self._config.model_id:
            idx = self.model_combo.findData(self._config.model_id)
            if idx >= 0:
                self.model_combo.setCurrentIndex(idx)

    def _populate_voice_combo(self) -> None:
        self.voice_combo.clear()
        for idx, (vid, desc) in enumerate(VOICE_CATALOG):
            self.voice_combo.addItem(f"{idx+1}. {vid} ({desc})", vid)

        saved_v = self._config.voice
        found = False
        for i in range(self.voice_combo.count()):
            if self.voice_combo.itemData(i) == saved_v:
                self.voice_combo.setCurrentIndex(i)
                found = True
                break
        if not found:
            for i in range(self.voice_combo.count()):
                if self.voice_combo.itemData(i) == "Kore":
                    self.voice_combo.setCurrentIndex(i)
                    break

    def _get_current_voice(self) -> str:
        data = self.voice_combo.currentData()
        if data:
            return str(data)
        raw = self.voice_combo.currentText()
        match = re.search(r"^[0-9]+\.\s*([A-Za-z]+)", raw)
        if match:
            return match.group(1)
        return "Kore"

    def _get_current_model_id(self) -> str:
        data = self.model_combo.currentData()
        return str(data) if data else "gemini-2.5-flash-preview-tts"

    def _update_models_status_label(self) -> None:
        count = len(self._config.all_fetched_models) if self._config.all_fetched_models else len(DEFAULT_MODELS)
        active = self.model_combo.count()
        self.models_status_label.setText(f"النماذج: {active} مفعل من أصل {count}")

    def _on_fetch_models(self) -> None:
        keys = self._get_all_keys()
        if not keys:
            QMessageBox.warning(self, "تنبيه", "الرجاء إضافة مفتاح API واحد على الأقل لجلب النماذج.")
            return
        self.fetch_models_btn.setEnabled(False)
        self.models_status_label.setText("جاري الاتصال بـ Google API لجلب النماذج...")
        self._fetch_thread = QThread()
        self._fetch_worker = ModelFetchWorker(keys[0])
        self._fetch_worker.moveToThread(self._fetch_thread)
        self._fetch_thread.started.connect(self._fetch_worker.run)
        self._fetch_worker.finished.connect(self._on_models_fetched)
        self._fetch_worker.error.connect(self._on_models_fetch_failed)
        self._fetch_worker.finished.connect(self._fetch_thread.quit)
        self._fetch_worker.error.connect(self._fetch_thread.quit)
        self._fetch_thread.start()

    def _on_models_fetched(self, models: list, recommended: list) -> None:
        self.fetch_models_btn.setEnabled(True)
        self._config.all_fetched_models = models
        self._config.enabled_model_ids = recommended if not self._config.enabled_model_ids else self._config.enabled_model_ids
        self._populate_model_combo()
        self._update_models_status_label()
        self._save_config()
        self._log(f"✓ تم تحديث النماذج الصوتية ({len(models)} نموذج متاح)", "#a6e3a1")

    def _on_models_fetch_failed(self, error: str) -> None:
        self.fetch_models_btn.setEnabled(True)
        self._update_models_status_label()
        self._log(f"✗ تعذر جلب النماذج: {error}", "#fab387")

    def _on_select_models(self) -> None:
        models = self._config.all_fetched_models if self._config.all_fetched_models else DEFAULT_MODELS
        dlg = ModelsSelectionDialog(self, models, self._config.enabled_model_ids)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            selected = dlg.get_selected()
            self._config.enabled_model_ids = [m["id"] for m in selected]
            self._populate_model_combo()
            self._update_models_status_label()
            self._save_config()

    def _on_add_keys(self) -> None:
        dlg = AddKeysDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_keys = dlg.get_keys()
            if not new_keys:
                return
            current_keys = self._get_all_keys()
            if dlg.is_replace():
                final_keys = new_keys
            else:
                combined = list(current_keys)
                for k in new_keys:
                    if k not in combined:
                        combined.append(k)
                final_keys = combined
            self._save_keys(final_keys)
            self._update_keys_label()
            self._log(f"🔑 تم تحديث قائمة المفاتيح: {len(final_keys)} مفتاح صالع", "#a6e3a1")

    def _update_keys_label(self) -> None:
        keys = self._get_all_keys()
        model_id = self._get_current_model_id()
        avail = self._global_key_mgr.get_available_keys(keys, model_id)
        self.keys_label.setText(f"🔑 إجمالي: {len(keys)} | نشطة ومتاحة: {len(avail)}")
        self.stat_keys.setText(f"🔑 {len(avail)}/{len(keys)}")
        self.blocked_label.setText(self._global_key_mgr.get_blocked_summary())

    def _validate_keys(self) -> None:
        keys = self._get_all_keys()
        if not keys:
            QMessageBox.warning(self, "تنبيه", UI["no_api_keys"])
            return
        self._log("🔍 جاري فحص والتحقق من صلاحية المفاتيح...", "#89b4fa")
        valid_count = 0
        for k in keys:
            status = self._global_key_mgr.get_key_status(k)
            if not status["is_fully_blocked"]:
                valid_count += 1
        self._update_keys_label()
        self._log(f"✓ تم التحقق: {valid_count} مفتاح متاح من أصل {len(keys)}", "#a6e3a1")

    def _show_keys_status(self) -> None:
        keys = self._get_all_keys()
        if not keys:
            QMessageBox.information(self, "المفاتيح", "لا يوجد مفاتيح مضافة حتى الآن.")
            return
        dlg = KeysStatusDialog(self, keys, self._global_key_mgr)
        dlg.exec()

    def _unblock_all_keys_action(self) -> None:
        count = self._global_key_mgr.unblock_all_keys()
        self._update_keys_label()
        self._log(f"🔓 تم فك الحظر اليدوي عن جميع المفاتيح وتصفير العدادات ({count} قيد أزيل)", "#a6e3a1")
        msg = f"تم فك الحظر وإزالة جميع قيود الدقائق عن جميع المفاتيح بنجاح ({count} قيد أزيل).\nأصبحت جميع المفاتيح مفتوحة ومتاحة للعمل فوراً."
        QMessageBox.information(self, "فك الحظر اليدوي", msg)

    def _browse_file(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, UI["open_file"], "", UI["file_filter"])
        if not paths:
            return
        if len(paths) == 1:
            self._load_file(paths[0])
        else:
            self._load_multiple_books(paths)

    def _load_file(self, path: str) -> None:
        try:
            text = load_text_file(path)
            self._source_text = text.strip()
            self.file_path_label.setText(path)
            self.progress_bar.setValue(0)
            self.stat_eta.setText("⏳ ETA: —")
            self.audio_dur_label.setText("🔊 الصوت المولد: —")
            self._config.source_file = path
            self._add_recent_file(path)
            self._save_config()

            chars = len(self._source_text)
            words = count_words(self._source_text)
            tokens = max(1, chars // 4)
            dur = estimate_duration_seconds(chars)
            mins, secs = int(dur // 60), int(dur % 60)
            self.text_stats_label.setText(
                f"{chars:,} {UI['chars']} | {words:,} {UI['words']} (~{tokens:,} {UI['tokens_est']})\n"
                f"{UI['est_duration']} {mins} {UI['minutes']} و {secs} {UI['seconds']}"
            )
            self.preview_text_btn.setEnabled(True)
            self.preview_chunks_btn.setEnabled(True)
            base = os.path.splitext(os.path.basename(path))[0]
            clean_name = re.sub(r'[^\w\s-]', '', base).strip().replace(' ', '_') or "project"
            self.project_name_input.setText(clean_name)
            self._build_chunks()
            
            # تسجيل الكتاب كمشروع مستقل في طابور المشاريع
            job = BookJob(file_path=path, project_name=clean_name, status="pending",
                          source_text=self._source_text, chunks=self._chunks)
            self._book_queue = [job]
            self._current_book_index = 0
            
            self._log(f"📄 تم تحميل الملف وجاهز للتحويل: {os.path.basename(path)} ({chars:,} حرف)", "#a6e3a1")
        except Exception as e:
            QMessageBox.critical(self, "خطأ في التحميل", f"تعذر قراءة الملف:\n{str(e)}")

    def _load_multiple_books(self, paths: List[str]) -> None:
        """Loads multiple books into self-contained project subfolders inside root output_folder."""
        self._book_queue = []
        for p in paths:
            if not os.path.exists(p):
                continue
            base = os.path.splitext(os.path.basename(p))[0]
            clean_name = re.sub(r'[^\w\s-]', '', base).strip().replace(' ', '_') or "book"
            try:
                txt = load_text_file(p).strip()
                if not txt:
                    continue
                mgr = ChunkManager(
                    txt, split_method=self._config.split_method,
                    chunk_duration=self.chunk_spin.value(),
                    chunk_words=self.chunk_words_spin.value(),
                    chunk_sentences=self.chunk_sentences_spin.value(),
                    context_carry=self.context_check.isChecked(),
                    auto_clean=self._config.auto_clean_symbols,
                    remove_diacritics=self._config.remove_diacritics
                )
                job = BookJob(file_path=p, project_name=clean_name, status="pending",
                              source_text=txt, chunks=mgr.chunks)
                self._book_queue.append(job)
            except Exception as e:
                self._log(f"✗ خطأ في قراءة الكتاب {os.path.basename(p)}: {str(e)}", "#f38ba8")

        if not self._book_queue:
            return
            
        self._current_book_index = 0
        first_book = self._book_queue[0]
        self._source_text = first_book.source_text
        self._chunks = first_book.chunks
        self.file_path_label.setText(first_book.file_path)
        self.project_name_input.setText(first_book.project_name)
        self._config.source_file = first_book.file_path
        self._save_config()
        self._populate_table()
        self._update_chunk_counter_display()
        
        chars = len(self._source_text)
        words = count_words(self._source_text)
        tokens = max(1, chars // 4)
        dur = estimate_duration_seconds(chars)
        mins, secs = int(dur // 60), int(dur % 60)
        self.text_stats_label.setText(
            f"{chars:,} {UI['chars']} | {words:,} {UI['words']} (~{tokens:,} {UI['tokens_est']}) | كتاب 1/{len(self._book_queue)}\n{UI['est_duration']} {mins} {UI['minutes']} و {secs} {UI['seconds']}"
        )
        self.preview_text_btn.setEnabled(True)
        self.preview_chunks_btn.setEnabled(True)
        self._log(f"📚 تم إضافة {len(self._book_queue)} كتب لطابور المعالجة المستقلة (كل كتاب بمجلد مستقل باسمه)", "#89b4fa")

    def _switch_to_next_book_in_queue(self) -> bool:
        """Seamlessly switches and auto-runs next book in queue (MultiBookOrchestrator)."""
        if not self._book_queue or self._current_book_index + 1 >= len(self._book_queue):
            return False
            
        self._current_book_index += 1
        next_book = self._book_queue[self._current_book_index]
        self._source_text = next_book.source_text
        self._chunks = next_book.chunks
        self.file_path_label.setText(next_book.file_path)
        self.project_name_input.setText(next_book.project_name)
        self._config.source_file = next_book.file_path
        self._save_config()
        self._populate_table()
        self._update_chunk_counter_display()
        
        chars = len(self._source_text)
        words = count_words(self._source_text)
        tokens = max(1, chars // 4)
        dur = estimate_duration_seconds(chars)
        mins, secs = int(dur // 60), int(dur % 60)
        self.text_stats_label.setText(
            f"{chars:,} {UI['chars']} | {words:,} {UI['words']} (~{tokens:,} {UI['tokens_est']}) | كتاب {self._current_book_index + 1}/{len(self._book_queue)}\n{UI['est_duration']} {mins} {UI['minutes']} و {secs} {UI['seconds']}"
        )
        self.progress_bar.setValue(0)
        self._log(f"📚 جاري الانتقال للكتاب التالي في الطابور ({self._current_book_index + 1}/{len(self._book_queue)}): {next_book.project_name}", "#89b4fa")
        self._start_processing()
        return True

    def _new_project_action(self) -> None:
        if self._worker and self._worker.isRunning():
            reply = QMessageBox.question(
                self, "تأكيد البدء من جديد",
                "هناك تحويل صوتي قيد التشغيل حالياً. هل تريد إيقافه والبدء بمشروع جديد؟ (تم حفظ نقطة الاستئناف للمشروع الحالي تلقائياً)",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._worker.stop()
                self._worker.wait(1000)
            else:
                return
        elif self._chunks and any(c.status != "pending" for c in self._chunks):
            reply = QMessageBox.question(
                self, "مشروع جديد",
                "هل تود تنظيف الواجهة والبدء بتحويل جديد بالكامل؟\n(لن يتم حذف أي ملفات صوتية تم حفظها بالفعل على القرص)",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        self._source_text = ""
        self._chunks = []
        self._book_queue = []
        self._current_book_index = -1
        self.file_path_label.clear()
        self.text_stats_label.setText("إحصائيات النص: 0 حرف | 0 كلمة")
        self.preview_text_btn.setEnabled(False)
        self.preview_chunks_btn.setEnabled(False)
        self.project_name_input.setText("project_new")
        self.progress_bar.setValue(0)
        self.chunks_table.setRowCount(0)
        self.current_chunk_label.setText("حالة العمليات: جاهز للعمل")
        self.status_bar_label.setText("جاهز للعمل")
        self.stat_eta.setText("⏳ ETA: —")
        self.audio_dur_label.setText("🔊 الصوت المولد: —")
        self.start_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.merge_btn.setEnabled(False)
        self.retry_all_btn.setEnabled(False)
        self._config.source_file = ""
        self._save_config()
        self._log("✨ تم تنظيف واجهة الاستوديو وبدء تحويل جديد جاهز للعمل", "#89b4fa")

    def _preview_text(self) -> None:
        if not self._source_text:
            return
        TextPreviewDialog(self._source_text, self).exec()

    def _browse_output(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "اختيار مجلد الإخراج الرئيسي", self.output_path_input.text())
        if path:
            self.output_path_input.setText(path)
            self._config.output_folder = path
            self._save_config()

    def _open_output_folder(self) -> None:
        out_dir = os.path.join(self.output_path_input.text().strip(), self.project_name_input.text().strip())
        if not os.path.exists(out_dir):
            out_dir = self.output_path_input.text().strip()
        if os.path.exists(out_dir):
            play_or_open_file(out_dir)
            self._log(f"📂 تم فتح مجلد الكتاب الحالي: {out_dir}", "#89b4fa")
        else:
            QMessageBox.information(self, "مجلد المخرجات", "مجلد مخرجات الكتاب غير موجود بعد. ابدأ المعالجة لإنشائه.")

    def _open_root_output_folder(self) -> None:
        root_dir = self.output_path_input.text().strip()
        if os.path.exists(root_dir):
            play_or_open_file(root_dir)
            self._log(f"📂 تم فتح المسار الرئيسي للمخرجات: {root_dir}", "#89b4fa")
        else:
            QMessageBox.information(self, "المسار الرئيسي", "المسار الرئيسي للمخرجات غير موجود بعد.")

    def _update_chunk_counter_display(self) -> None:
        if not hasattr(self, "chunk_counter_label") or not self._chunks:
            if hasattr(self, "chunk_counter_label"):
                self.chunk_counter_label.setText("📊 المنجز من المقاطع: 0 / 0 مقطع (0.0%)")
            return
        total = len(self._chunks)
        done = sum(1 for c in self._chunks if c.status == "done")
        failed = sum(1 for c in self._chunks if c.status == "failed")
        processing = sum(1 for c in self._chunks if c.status == "processing")
        pct = (done / total * 100.0) if total > 0 else 0.0
        
        status_txt = f"📊 المنجز من المقاطع: {done} / {total} مقطع ({pct:.1f}%)"
        if processing > 0:
            status_txt += f" | ⚡ جاري: {processing}"
        if failed > 0:
            status_txt += f" | ❌ فاشل: {failed}"
        self.chunk_counter_label.setText(status_txt)

    def _load_recent_files_menu(self) -> None:
        self.recent_menu.clear()
        for fpath in self._config.recent_files:
            if os.path.exists(fpath):
                action = QAction(os.path.basename(fpath), self)
                action.triggered.connect(lambda ch, p=fpath: self._load_file(p))
                self.recent_menu.addAction(action)

    def _add_recent_file(self, path: str) -> None:
        if path in self._config.recent_files:
            self._config.recent_files.remove(path)
        self._config.recent_files.insert(0, path)
        self._config.recent_files = self._config.recent_files[:10]
        self._load_recent_files_menu()

    def _update_chunk_estimate(self, val: int = 0) -> None:
        s_idx = self.split_method_combo.currentIndex()
        if not self._source_text:
            self.chunk_est_label.setText("(~1 مقطع لكل 900 حرف)")
            return
            
        if s_idx == 1:  # words
            w_target = self.chunk_words_spin.value()
            w_total = max(1, count_words(self._source_text))
            est = max(1, (w_total + w_target - 1) // w_target)
            self.chunk_est_label.setText(f"(~{est} مقطع حسب الكلمات)")
        elif s_idx == 2:  # sentences
            s_target = self.chunk_sentences_spin.value()
            s_total = max(1, len(re.split(r'[.!?؟…\n]+', self._source_text)))
            est = max(1, (s_total + s_target - 1) // s_target)
            self.chunk_est_label.setText(f"(~{est} مقطع حسب الجمل)")
        else:  # seconds
            dur = self.chunk_spin.value()
            chars_per_chunk = dur * CHARS_PER_SECOND
            est = max(1, (len(self._source_text) + chars_per_chunk - 1) // chars_per_chunk)
            self.chunk_est_label.setText(f"(~{est} مقطع زمنياً)")
            
        self._build_chunks()

    def _build_chunks(self) -> List[ChunkInfo]:
        if not self._source_text:
            return []
            
        s_idx = self.split_method_combo.currentIndex()
        method = "seconds"
        if s_idx == 1: method = "words"
        elif s_idx == 2: method = "sentences"
        
        mgr = ChunkManager(
            self._source_text, split_method=method,
            chunk_duration=self.chunk_spin.value(),
            chunk_words=self.chunk_words_spin.value(),
            chunk_sentences=self.chunk_sentences_spin.value(),
            context_carry=self.context_check.isChecked(),
            auto_clean=self._config.auto_clean_symbols,
            remove_diacritics=self._config.remove_diacritics
        )
        self._chunks = mgr.chunks
        self._populate_table()
        self._update_chunk_counter_display()
        self.retry_all_btn.setEnabled(any(c.status == "failed" for c in self._chunks))
        return self._chunks

    def _preview_chunks(self) -> None:
        if not self._chunks:
            self._build_chunks()
        if self._chunks:
            ChunkPreviewDialog(self._chunks, self).exec()

    def _preview_voice(self) -> None:
        voice = self._get_current_voice()
        style_name = self.reading_style_combo.currentText()
        clean_style = re.sub(r'[^\w\s-]', '', style_name).strip().replace(' ', '_')
        
        try:
            VOICE_PREVIEWS_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
            
        cached_preview_path = VOICE_PREVIEWS_DIR / f"preview_{voice}_{clean_style}.wav"
        
        # 1. التحقق من الذاكرة المحلية أولاً: إذا كانت العينة مخزنة وصالحة، يتم تشغيلها فوراً بدون طلب API!
        if cached_preview_path.exists() and cached_preview_path.stat().st_size > 500:
            play_or_open_file(str(cached_preview_path))
            self._log(f"🔊 تشغيل عينة الصوت المحفوظة محلياً للنبرة: {voice} ({style_name}) [بدون استهلاك كوتة]", "#a6e3a1")
            return

        keys = self._get_all_keys()
        if not keys:
            QMessageBox.warning(self, "تنبيه", UI["no_api_keys"])
            return
        self.preview_voice_btn.setEnabled(False)
        self.preview_voice_btn.setText("جاري التوليد...")
        model_id = self._get_current_model_id()
        style_instruction = READING_STYLES.get(style_name, READING_STYLES["رواية (افتراضي)"])

        def _do():
            try:
                client = genai.Client(api_key=keys[0])
                sample_text = f"مرحباً بك في استوديو جيميناي تي تي إس برو الإصدار {APP_VERSION}. هذا إلقاء تجريبي بصوت {voice}."
                speech_cfg = types.SpeechConfig(
                    voice_config=types.VoiceConfig(
                        prebuilt_voice_config=types.PrebuiltVoiceConfig(
                            voice_name=voice)))
                use_sys_inst = not ("preview-tts" in model_id.lower() or "tts" in model_id.lower())
                response = None
                if use_sys_inst:
                    try:
                        response = client.models.generate_content(
                            model=model_id,
                            contents=sample_text,
                            config=types.GenerateContentConfig(
                                system_instruction=style_instruction,
                                response_modalities=["AUDIO"],
                                speech_config=speech_cfg,
                            ),
                        )
                    except Exception as e:
                        if is_developer_instruction_error(e):
                            use_sys_inst = False
                        else:
                            raise
                if not use_sys_inst or response is None:
                    response = client.models.generate_content(
                        model=model_id,
                        contents=sample_text,
                        config=types.GenerateContentConfig(
                            response_modalities=["AUDIO"],
                            speech_config=speech_cfg,
                        ),
                    )
                audio_data = response.candidates[0].content.parts[0].inline_data.data
                out_path = str(cached_preview_path)
                with wave.open(out_path, "wb") as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(DEFAULT_SAMPLE_RATE)
                    wf.writeframes(audio_data)
                return True, out_path
            except Exception as e:
                return False, str(e)

        import concurrent.futures
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        future = executor.submit(_do)

        def _check():
            if future.done():
                self.preview_voice_btn.setEnabled(True)
                self.preview_voice_btn.setText(UI["preview_voice"])
                success, res = future.result()
                if success:
                    play_or_open_file(res)
                    self._log(f"🔊 تم توليد وحفظ عينة الصوت محلياً لتشغيلها لاحقاً فوراً: {voice}", "#a6e3a1")
                else:
                    self._log(f"✗ فشل توليد العينة الصوتية: {res}", "#f38ba8")
            else:
                QTimer.singleShot(150, _check)
        QTimer.singleShot(150, _check)

    def _validate_inputs(self) -> bool:
        if not self._get_all_keys():
            QMessageBox.warning(self, "تنبيه", UI["no_api_keys"])
            return False
            
        # فحص ذاتي ذكي: إذا كان هناك مسار ملف صالح في مربع المسار ولكن النص فارغ لأي سبب، نعيد قراءة الملف فوراً بدل إظهار رسالة تنبيه
        if not self._source_text and self.file_path_label.text().strip():
            fpath = self.file_path_label.text().strip()
            if os.path.exists(fpath):
                self._load_file(fpath)
                
        if not self._source_text:
            QMessageBox.warning(self, "تنبيه", UI["no_text"])
            return False
        if not self.output_path_input.text().strip():
            QMessageBox.warning(self, "تنبيه", UI["no_output"])
            return False
        if not self.project_name_input.text().strip():
            QMessageBox.warning(self, "تنبيه", UI["no_project"])
            return False
        return True

    def _start_processing(self, is_auto_retry_round: bool = False) -> None:
        if not is_auto_retry_round:
            self._current_retry_round = 0
        if not self._validate_inputs():
            return
        if not self._chunks:
            self._build_chunks()
        if not self._chunks:
            return

        self._save_config()
        self._update_keys_label()

        out_dir = os.path.join(self.output_path_input.text().strip(), self.project_name_input.text().strip())
        os.makedirs(out_dir, exist_ok=True)

        self.start_btn.setEnabled(False)
        self.pause_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)
        self.merge_btn.setEnabled(False)
        self.retry_all_btn.setEnabled(False)
        self.progress_bar.setValue(0)

        keys = self._get_all_keys()
        self._worker = TTSWorkerThread(
            chunks=self._chunks,
            model_id=self._get_current_model_id(),
            voice=self._get_current_voice(),
            speed=self.speed_slider.value() / 10.0,
            lang_hint=self.lang_input.text().strip(),
            output_dir=out_dir,
            project_name=self.project_name_input.text().strip(),
            api_keys=keys,
            max_workers=self._config.max_workers,
            retry_count=self._config.retry_count,
            reading_style=self.reading_style_combo.currentText(),
            global_key_mgr=self._global_key_mgr,
            adaptive_parallel=self._config.adaptive_parallel,
            retry_failed_forever=self._config.retry_failed_forever
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.status.connect(self._on_status)
        self._worker.chunk_started.connect(self._on_chunk_started)
        self._worker.chunk_done.connect(self._on_chunk_done)
        self._worker.chunk_failed.connect(self._on_chunk_failed)
        self._worker.log_message.connect(self._on_log_message)
        self._worker.finished_signal.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.key_used.connect(self._on_key_used)

        self._eta_timer.start()
        self._worker.start()

    def _toggle_pause(self) -> None:
        if not self._worker:
            return
        if self._worker.is_paused():
            self._worker.resume()
            self.pause_btn.setText(UI["pause"])
            self.status_bar_label.setText("جاري المعالجة...")
            self._log("▶ تم استئناف المعالجة", "#a6e3a1")
        else:
            self._worker.pause()
            self.pause_btn.setText("▶ استئناف")
            self.status_bar_label.setText("إيقاف مؤقت")
            self._log("⏸ تم الإيقاف المؤقت", "#fab387")

    def _stop_processing(self) -> None:
        if not self._worker:
            return
        reply = QMessageBox.question(self, "تأكيد الإيقاف", UI["confirm_stop"],
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self._worker.stop()
            self.stop_btn.setEnabled(False)
            self.pause_btn.setEnabled(False)

    def _populate_table(self) -> None:
        self.chunks_table.setRowCount(len(self._chunks))
        for i, c in enumerate(self._chunks):
            self.chunks_table.setItem(i, 0, QTableWidgetItem(str(c.index + 1)))
            
            st_item = QTableWidgetItem(c.status)
            if c.status == "done":
                st_item.setForeground(QBrush(QColor("#a6e3a1")))
            elif c.status == "processing":
                st_item.setForeground(QBrush(QColor("#89b4fa")))
            elif c.status == "failed":
                st_item.setForeground(QBrush(QColor("#f38ba8")))
            self.chunks_table.setItem(i, 1, st_item)
            
            self.chunks_table.setItem(i, 2, QTableWidgetItem(str(c.char_count)))
            self.chunks_table.setItem(i, 3, QTableWidgetItem(f"{c.estimated_duration:.1f} ث"))
            self.chunks_table.setItem(i, 4, QTableWidgetItem(os.path.basename(c.output_file) if c.output_file else "—"))
            self.chunks_table.setItem(i, 5, QTableWidgetItem(f"{c.time_taken:.1f}ث" if c.time_taken > 0 else "—"))
            
            btn = QPushButton("🔁 إعادة محاولة")
            btn.setMinimumHeight(24)
            btn.clicked.connect(lambda ch, idx=i: self._retry_single_chunk(idx))
            btn.setEnabled(c.status == "failed")
            self.chunks_table.setCellWidget(i, 6, btn)

    def _on_table_double_click(self, row: int, col: int) -> None:
        if row < 0 or row >= len(self._chunks):
            return
        c = self._chunks[row]
        if c.status == "done" and c.output_file and os.path.exists(c.output_file):
            play_or_open_file(c.output_file)
            self._log(f"🔊 تشغيل المقطع {row + 1}: {os.path.basename(c.output_file)}", "#a6e3a1")

    def _show_table_context_menu(self, pos) -> None:
        row = self.chunks_table.rowAt(pos.y())
        if row < 0 or row >= len(self._chunks):
            return
        c = self._chunks[row]
        menu = QMenu(self)
        
        play_action = menu.addAction("🔊 تشغيل الملف الصوتي")
        play_action.setEnabled(c.status == "done" and bool(c.output_file) and os.path.exists(c.output_file))
        play_action.triggered.connect(lambda: play_or_open_file(c.output_file))
        
        retry_action = menu.addAction("🔁 إعادة محاولة هذا المقطع")
        retry_action.triggered.connect(lambda: self._retry_single_chunk(row))
        
        copy_text_action = menu.addAction("📋 نسخ نص المقطع إلى الحافظة")
        copy_text_action.triggered.connect(lambda: QApplication.clipboard().setText(c.text))
        
        folder_action = menu.addAction("📂 فتح موقع الملف في المتصفح")
        folder_action.setEnabled(bool(c.output_file) and os.path.exists(c.output_file))
        folder_action.triggered.connect(lambda: play_or_open_file(os.path.dirname(c.output_file)))
        
        menu.exec(self.chunks_table.viewport().mapToGlobal(pos))

    def _on_progress(self, pct: int) -> None:
        self.progress_bar.setValue(pct)

    def _on_status(self, msg: str) -> None:
        self.status_bar_label.setText(msg)

    def _on_chunk_started(self, idx: int) -> None:
        if idx < len(self._chunks):
            self._chunks[idx].status = "processing"
            self._update_chunk_counter_display()
            st_item = QTableWidgetItem("⚡ جاري المعالجة")
            st_item.setForeground(QBrush(QColor("#89b4fa")))
            self.chunks_table.setItem(idx, 1, st_item)
            self.current_chunk_label.setText(f"{UI['current_chunk']} المقطع {idx + 1}/{len(self._chunks)}")

    def _on_chunk_done(self, idx: int, filepath: str) -> None:
        if idx < len(self._chunks):
            c = self._chunks[idx]
            c.status = "done"
            c.output_file = filepath
            self._update_chunk_counter_display()
            st_item = QTableWidgetItem("✅ مكتمل")
            st_item.setForeground(QBrush(QColor("#a6e3a1")))
            self.chunks_table.setItem(idx, 1, st_item)
            self.chunks_table.setItem(idx, 4, QTableWidgetItem(os.path.basename(filepath)))
            
            btn = self.chunks_table.cellWidget(idx, 6)
            if isinstance(btn, QPushButton):
                btn.setEnabled(False)
            self.retry_all_btn.setEnabled(any(ch.status == "failed" for ch in self._chunks))

    def _on_chunk_failed(self, idx: int, error: str) -> None:
        if idx < len(self._chunks):
            c = self._chunks[idx]
            c.status = "failed"
            c.error_msg = error
            self._update_chunk_counter_display()
            st_item = QTableWidgetItem("❌ فشل")
            st_item.setForeground(QBrush(QColor("#f38ba8")))
            self.chunks_table.setItem(idx, 1, st_item)
            
            btn = self.chunks_table.cellWidget(idx, 6)
            if isinstance(btn, QPushButton):
                btn.setEnabled(True)
            self.retry_all_btn.setEnabled(True)

    def _retry_single_chunk(self, chunk_idx: int) -> None:
        if not self._validate_inputs():
            return
        if chunk_idx < 0 or chunk_idx >= len(self._chunks):
            return
        chunk = self._chunks[chunk_idx]
        self._log(f"🔁 جاري إعادة محاولة المقطع {chunk_idx + 1} منفردًا...", "#89b4fa")
        
        btn = self.chunks_table.cellWidget(chunk_idx, 6)
        if isinstance(btn, QPushButton):
            btn.setEnabled(False)
            
        out_dir = os.path.join(self.output_path_input.text().strip(), self.project_name_input.text().strip())
        os.makedirs(out_dir, exist_ok=True)
        keys = self._get_all_keys()

        self._on_chunk_started(chunk_idx)

        def _do_retry():
            worker = TTSWorkerThread([chunk], self._get_current_model_id(), self.voice_combo.currentText(),
                                     self.speed_slider.value() / 10.0, self.lang_input.text().strip(),
                                     out_dir, self.project_name_input.text().strip(), keys,
                                     max_workers=1, retry_count=self._config.retry_count,
                                     reading_style=self.reading_style_combo.currentText(),
                                     global_key_mgr=self._global_key_mgr,
                                     adaptive_parallel=self._config.adaptive_parallel,
                                     retry_failed_forever=self._config.retry_failed_forever)
            idx, success, res = worker._process_single_chunk(chunk)
            return idx, success, res

        import concurrent.futures
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        future = executor.submit(_do_retry)

        def _check_retry():
            if future.done():
                idx, success, res = future.result()
                if success:
                    self._on_chunk_done(idx, res)
                    self._log(f"✓ نجحت إعادة محاولة المقطع {idx + 1}", "#a6e3a1")
                else:
                    self._on_chunk_failed(idx, res)
                    self._log(f"✗ فشلت محاولة المقطع {idx + 1}: {res}", "#f38ba8")
            else:
                QTimer.singleShot(200, _check_retry)
        QTimer.singleShot(200, _check_retry)

    def _retry_all_failed_chunks(self) -> None:
        failed_indices = [c.index for c in self._chunks if c.status == "failed"]
        if not failed_indices:
            self._log("ℹ️ لا توجد مقاطع فاشلة لإعادة محاولتها", "#a6e3a1")
            return
        self._log(f"🔁 جاري إعادة محاولة جميع المقاطع الفاشلة ({len(failed_indices)} مقطع)...", "#89b4fa")
        self._start_processing()

    def _on_finished(self, msg: str) -> None:
        self._eta_timer.stop()
        self.start_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.status_bar_label.setText(msg)
        self.stat_eta.setText("⏳ ETA: مكتمل")

        completed = self._get_completed_wav_files()
        total_dur = 0.0
        for f in completed:
            try:
                with wave.open(f, "rb") as wf:
                    total_dur += wf.getnframes() / wf.getframerate()
            except Exception:
                pass
        mins, secs = int(total_dur // 60), int(total_dur % 60)
        self.audio_dur_label.setText(f"🔊 الصوت المولد: {mins}د و {secs}ث ({len(completed)} ملف)")
        
        self.merge_btn.setEnabled(len(completed) > 1)
        failed_chunks = [c for c in self._chunks if c.status == "failed"]
        self.retry_all_btn.setEnabled(len(failed_chunks) > 0)

        # التحقق من وجود مقاطع فاشلة وتشغيل جولة الإعادة التلقائية إن لم نصل للحد الأقصى
        if failed_chunks and not self._closing and self._current_retry_round < self._config.max_retry_rounds:
            self._current_retry_round += 1
            self._log(f"⚡ [جولة الإعادة التلقائية {self._current_retry_round}/{self._config.max_retry_rounds}] تم اكتشاف {len(failed_chunks)} مقطع فاشل. جاري إعادة محاولتها تلقائياً الآن...", "#fab387")
            QTimer.singleShot(1500, lambda: self._start_processing(is_auto_retry_round=True))
            return

        if len(completed) == len(self._chunks) and len(completed) > 0:
            if self._current_book_index >= 0 and self._current_book_index < len(self._book_queue):
                self._book_queue[self._current_book_index].status = "done"
            self._log(f"🎉 {msg}", "#a6e3a1")
            if self.auto_merge_check.isChecked():
                if len(completed) > 1:
                    self._log("⚡ جاري تنفيذ الدمج التلقائي المباشر للمقاطع المكتملة...", "#89b4fa")
                    self._merge_all_finalize()
                elif len(completed) == 1:
                    self._log("⚡ جاري نسخ المقطع الوحيد إلى ملف المشروع النهائي المدمج داخل الوجهة مباشرة...", "#89b4fa")
                    project_name = self.project_name_input.text().strip()
                    output_final = os.path.join(self.output_path_input.text().strip(), f"{project_name}_FINAL.wav")
                    try:
                        shutil.copy2(completed[0], output_final)
                        self._on_merge_done(output_final, [output_final])
                    except Exception as e:
                        self._log(f"❌ تعذر حفظ الملف النهائي المدمج: {str(e)}", "#f38ba8")
            else:
                if not self._switch_to_next_book_in_queue():
                    if self._config.shutdown_after_queue:
                        self._trigger_auto_shutdown()
        else:
            self._log(f"⚠️ {msg} (تبقت {len(failed_chunks)} مقاطع فاشلة بعد اكتمال {self._current_retry_round} جولات إعادة)", "#fab387")

    def _trigger_auto_shutdown(self) -> None:
        self._log("⚠️ تم تفعيل الإيقاف التلقائي للنظام بعد اكتمال المشروع — بدء العد التنازلي (60 ثانية)...", "#fab387")
        reply = QMessageBox.warning(
            self, "إيقاف تشغيل الكمبيوتر",
            "🎉 تم إنجاز التحويل الصوتي بالكامل بنجاح!\n\n⚠️ خيار إيقاف تشغيل الكمبيوتر تلقائياً مفعل.\nسيتم إيقاف تشغيل النظام خلال 60 ثانية...\n\nانقر على «إلغاء» لإيقاف العد التنازلي والبقاء في الاستوديو.",
            QMessageBox.StandardButton.Cancel, QMessageBox.StandardButton.Cancel
        )
        if reply == QMessageBox.StandardButton.Cancel:
            self._log("ℹ️ تم إلغاء إيقاف تشغيل الكمبيوتر بواسطة المستخدم.", "#89b4fa")
        else:
            try:
                if sys.platform == "win32":
                    os.system("shutdown /s /t 60")
                else:
                    os.system("sudo shutdown -h +1")
            except Exception as e:
                self._log(f"❌ تعذر إيقاف التشغيل التلقائي: {str(e)}", "#f38ba8")

    def _on_error(self, msg: str) -> None:
        self._log(f"❌ {msg}", "#f38ba8")

    def _on_log_message(self, msg: str, color: str) -> None:
        self._log(msg, color)

    def _on_key_used(self, masked_key: str) -> None:
        self.status_bar_label.setText(f"⚡ جاري التوليد على المفتاح: {masked_key}")

    def _update_eta(self) -> None:
        if not self._worker or not self._chunks:
            return
        completed = sum(1 for c in self._chunks if c.status == "done")
        total = len(self._chunks)
        if completed == 0 or completed == total:
            return
        elapsed = time.time() - getattr(self._worker, "_start_timestamp", time.time())
        if elapsed < 2.0:
            return
        rate = completed / elapsed
        rem = (total - completed) / rate if rate > 0 else 0
        m, s = int(rem // 60), int(rem % 60)
        self.stat_eta.setText(f"⏳ متبقي: ~{m}د و {s}ث ({completed}/{total})")

    def _get_completed_wav_files(self) -> List[str]:
        out_dir = os.path.join(self.output_path_input.text().strip(), self.project_name_input.text().strip())
        files = []
        for c in self._chunks:
            if c.status == "done" and c.output_file and os.path.exists(c.output_file):
                files.append(c.output_file)
            else:
                p = os.path.join(out_dir, f"{self.project_name_input.text().strip()}_chunk_{c.index:04d}.wav")
                if os.path.exists(p):
                    c.status = "done"
                    c.output_file = p
                    files.append(p)
        return files

    def _merge_available(self) -> None:
        files = self._get_completed_wav_files()
        if len(files) < 2:
            QMessageBox.information(self, "دمج المقاطع", "تحتاج إلى مقطع مكتمل واحد على الأقل للدمج.")
            return
        project_name = self.project_name_input.text().strip()
        output_dir_chunks = os.path.join(self.output_path_input.text().strip(), project_name)
        output_final = os.path.join(self.output_path_input.text().strip(), f"{project_name}_PARTIAL.wav")
        self._run_merge(files, output_final, output_dir_chunks)

    def _merge_all_finalize(self) -> None:
        files = self._get_completed_wav_files()
        if not files:
            self._log("ℹ️ لا توجد ملفات مكتملة للدمج النهائي", "#fab387")
            return
        project_name = self.project_name_input.text().strip()
        output_dir_chunks = os.path.join(self.output_path_input.text().strip(), project_name)
        output_final = os.path.join(self.output_path_input.text().strip(), f"{project_name}_FINAL.wav")
        self._run_merge(files, output_final, output_dir_chunks)

    def _run_merge(self, files: List[str], output_path: str, chunks_dir: str = "") -> None:
        project_name = self.project_name_input.text().strip()
        list_dir = chunks_dir if chunks_dir else self.output_path_input.text().strip()
        self._merge_worker = MergeWorker(
            wav_files=files,
            output_path=output_path,
            project_name=project_name,
            output_dir=list_dir,
            silence_pad_ms=self._config.silence_pad_ms,
            use_ffmpeg=self._ffmpeg_available,
            split_method=self._config.output_split_method,
            split_parts=self._config.output_split_parts,
            split_minutes=self._config.output_split_minutes,
            split_size_mb=self._config.output_split_size_mb,
            split_overlap_sec=self._config.output_split_overlap_sec,
            keep_temp_files=self._config.keep_temp_files,
            audio_bitrate=self._config.audio_bitrate
        )
        self._merge_worker.progress.connect(lambda m: self._log(f"🔀 {m}", "#89b4fa"))
        self._merge_worker.finished_signal.connect(self._on_merge_done)
        self._merge_worker.error.connect(lambda e: self._log(f"✗ خطأ في الدمج: {e}", "#f38ba8"))
        self._merge_worker.start()

    def _on_merge_done(self, path: str, split_files: list) -> None:
        self._log(f"✅ تم دمج الصوت النهائي وتجهيز المخرجات بنجاح: {path}", "#a6e3a1")
        if len(split_files) > 1:
            self._log(f"📦 تم تقسيم المخرج المدمج إلى {len(split_files)} أجزاء صوتية متسلسلة", "#89b4fa")
            
        if self._book_queue and self._current_book_index + 1 < len(self._book_queue):
            self._log(f"✅ تم دمج الكتاب الحالي ({os.path.basename(path)}) بنجاح. جاري الانتقال للكتاب التالي تلقائياً...", "#a6e3a1")
            self._switch_to_next_book_in_queue()
        else:
            reply = QMessageBox.question(
                self, "تم الدمج بنجاح",
                f"{UI['merge_success']}\n{path}\n" + (f"\n📦 تم التقسيم إلى: {len(split_files)} ملفات\n" if len(split_files) > 1 else "") + "\nهل تود تشغيل الملف الصوتي الآن؟",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                play_or_open_file(path)

            if self._config.shutdown_after_queue:
                self._trigger_auto_shutdown()

    def _log(self, message: str, color: str = "#6c7086") -> None:
        if self._closing:
            return
        cursor = self.log_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_edit.setTextCursor(cursor)
        ts = datetime.now().strftime("%H:%M:%S")
        html = f'<div style="color: {color}; margin-bottom: 2px;">[{ts}] {message}</div>'
        cursor.insertHtml(html)
        self.log_edit.setTextCursor(cursor)
        self.log_edit.ensureCursorVisible()

    def _export_log(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "حفظ ملف السجل", f"tts_studio_log_{int(time.time())}.html", "HTML Log (*.html);;Text Log (*.txt)")
        if not path:
            return
        try:
            if path.endswith(".html"):
                atomic_write_text(Path(path), self.log_edit.toHtml())
            else:
                atomic_write_text(Path(path), self.log_edit.toPlainText())
            self._log(f"💾 تم تصدير السجل بنجاح إلى: {path}", "#a6e3a1")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"تعذر حفظ السجل: {str(e)}")

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self._config, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            vals = dlg.get_values()
            self._config.max_workers = vals["max_workers"]
            self._config.retry_count = vals["retry_count"]
            self._config.max_retry_rounds = vals["max_retry_rounds"]
            self._config.adaptive_parallel = vals["adaptive_parallel"]
            self._config.retry_failed_forever = vals["retry_failed_forever"]
            self._config.shutdown_after_queue = vals["shutdown_after_queue"]
            self._config.silence_pad_ms = vals["silence_pad_ms"]
            self._config.keep_temp_files = vals["keep_temp_files"]
            self._config.output_split_method = vals["output_split_method"]
            self._config.output_split_parts = vals["output_split_parts"]
            self._config.output_split_minutes = vals["output_split_minutes"]
            self._config.output_split_size_mb = vals["output_split_size_mb"]
            self._config.output_split_overlap_sec = vals["output_split_overlap_sec"]
            self._config.auto_clean_symbols = vals["auto_clean_symbols"]
            self._config.remove_diacritics = vals["remove_diacritics"]
            if self._config.dark_theme != vals["dark_theme"]:
                self._config.dark_theme = vals["dark_theme"]
                self._apply_theme()
            self._save_config()
            self._build_chunks()
            self._log("⚙️ تم حفظ الإعدادات المتقدمة وتطبيقها بنجاح", "#a6e3a1")

    def closeEvent(self, event) -> None:
        if self._worker and self._worker.isRunning():
            reply = QMessageBox.question(
                self, "تأكيد الخروج",
                "المعالجة الصوتية تعمل حالياً. هل تريد إيقافها وإغلاق الاستوديو؟ (تم حفظ نقطة استئناف تلقائياً)",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._closing = True
                self._worker.stop()
                self._worker.wait(1500)
                event.accept()
            else:
                event.ignore()
                return
        self._closing = True
        self._save_config()
        event.accept()


# ═══════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════

def _global_exception_handler(exctype, value, tb):
    """Global crash protection hook integrated from PDF Master Pro v29.4."""
    err_msg = "".join(traceback.format_exception(exctype, value, tb))
    try:
        log_file = CONFIG_DIR / "crash_report.log"
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"\n--- CRASH AT {datetime.now().isoformat()} ---\n{err_msg}\n")
    except Exception:
        pass
    print(f"CRASH INTERCEPTED: {err_msg}", file=sys.stderr)

def main() -> None:
    sys.excepthook = _global_exception_handler
    if hasattr(threading, "excepthook"):
        threading.excepthook = lambda args: _global_exception_handler(args.exc_type, args.exc_value, args.exc_traceback)
        
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
