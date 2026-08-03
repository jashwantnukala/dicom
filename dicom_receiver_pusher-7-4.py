# =========================================================
# R-Apps DICOM Receiver + AutoRouter
# =========================================================
# FEATURES
# ---------------------------------------------------------
# - Two professional tabs: RECEIVER and PUSHER, plus LOGS tab
# - Full DICOM Storage SCP (all modalities + extra SOPs via sopclass.ini)
# - Encrypted Receiver/Push configuration (rec.enc / push.enc)
# - Multiple push destinations with profile management
# - Routing rules engine (route by modality / institution / AE title)
# - Optional de-identification (anonymization) before push
# - DICOM Query/Retrieve (C-FIND / C-MOVE) from a remote PACS
# - TLS scaffold for DICOM associations (off by default; needs certs)
# - Manual local config confirmation (no external OTP/email dependency)
# - Folder IMPORT facility (recursively scans for .dcm / .dic / .dicom,
#   including extensionless DICOM files) directly into the worklist
# - Automatic worklist refresh (event-driven + background timer fallback)
# - Multi-threaded Pusher (configurable worker threads) with:
#       * per-patient progress bar (sent images / total images)
#       * overall progress bar (sent / total across the whole push job)
#       * live colour-coded status in the worklist
#       * automatic retry with exponential backoff on transient failures
#       * accurate "sent" vs "attempted" counters
#       * throughput / ETA display
# - Receiver and Pusher share one worklist (PatientID indexed) with a
#   "Source" column (Received / Imported) and a "Status" column, with
#   optional study-level grouping
# - Stale-pending highlighting (studies sitting un-pushed too long)
# - In-GUI editor for sopclass.ini (SOP Classes + Transfer Syntaxes)
# - In-GUI log viewer (receiver_errors.log / push_errors.log / audit.log)
# - Desktop (toast) notifications on receive / push completion or failure
# - Connection audit trail (who connected, when, success/failure)
# - Disk space monitoring with low-space warnings
# - Editable saved configs (no more permanently-locked fields)
# - System tray minimize support with graceful shutdown
# =========================================================

import os
import sys
import re
import csv
import json
import shutil
import random
import string
import platform
import datetime
import threading
import subprocess
import configparser
import argparse
import queue
import time
import traceback
import logging
import hashlib
import secrets
import socket
import contextlib
from collections import deque, Counter, defaultdict
from logging.handlers import RotatingFileHandler
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    psutil = None
    PSUTIL_AVAILABLE = False

try:
    import pyzipper
    PYZIPPER_AVAILABLE = True
except ImportError:
    pyzipper = None
    PYZIPPER_AVAILABLE = False

try:
    import ldap3
    from ldap3 import Server, Connection, ALL, SUBTREE, SIMPLE
    LDAP3_AVAILABLE = True
except ImportError:
    ldap3 = None
    LDAP3_AVAILABLE = False

try:
    import matplotlib
    matplotlib.use("Agg")  # headless rendering -- this app never shows a matplotlib window
    import matplotlib.pyplot as plt
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak,
    )
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

import customtkinter as ctk
import pydicom

# Optional pixel-data codec plugins. pydicom auto-detects these once
# installed and our transcode fallback in _transcode_dataset_for_target()
# will use them automatically -- no code change needed to use them, they
# just need to actually be importable. Without them, ds.decompress() on a
# JPEG2000/JPEG-LS/RLE file raises, and any push to a PACS that doesn't
# natively accept that transfer syntax fails outright instead of falling
# back to a re-encode. We only check importability here (for a startup
# warning); we never reference these modules directly elsewhere.
#
# §fix (transfer-syntax overhaul): DECODE and ENCODE are checked
# separately, because they are NOT the same capability:
#   - python-gdcm can DECODE (decompress) most compressed transfer
#     syntaxes for pydicom, but it is decode-only -- it cannot be used
#     by Dataset.compress() to re-COMPRESS pixel data.
#   - the pylibjpeg family (pylibjpeg + pylibjpeg-libjpeg/-openjpeg/-rle)
#     can both decode AND encode.
# Previously this app only ever called ds.decompress() and never
# ds.compress(), so this distinction didn't matter -- but that also meant
# any push to a peer that only accepted a COMPRESSED transfer syntax
# silently sent raw pixel data mislabeled as compressed (see
# _transcode_dataset_for_target()), which is the root cause of most
# "transfer syntax" failures/corrupted-image reports. That code path now
# actually compresses, so the encode/decode distinction matters here.
try:
    import gdcm  # noqa: F401  (python-gdcm) -- alternate decode-only backend
    GDCM_AVAILABLE = True
except Exception:
    GDCM_AVAILABLE = False

_CODEC_PLUGINS = ("pylibjpeg", "pylibjpeg_libjpeg", "pylibjpeg_openjpeg", "pylibjpeg_rle")
MISSING_CODEC_PLUGINS = []
for _plugin in _CODEC_PLUGINS:
    try:
        __import__(_plugin)
    except Exception:
        MISSING_CODEC_PLUGINS.append(_plugin)

DECODE_CODECS_AVAILABLE = GDCM_AVAILABLE or not MISSING_CODEC_PLUGINS
ENCODE_CODECS_AVAILABLE = not MISSING_CODEC_PLUGINS
CODECS_FULLY_AVAILABLE = DECODE_CODECS_AVAILABLE  # kept: existing startup-warning check

try:
    import io as _icon_io
    import base64 as _icon_b64
    from lucide_icons_data import LUCIDE_ICONS_B64
    LUCIDE_ICONS_AVAILABLE = True
except Exception:
    LUCIDE_ICONS_AVAILABLE = False
    LUCIDE_ICONS_B64 = {}

# Optional native drag-and-drop for the Attachments feature (Inspector
# panel). Genuinely optional: this app must work identically with or
# without it, since it isn't a standard part of every Python/Tk
# install and needs a compiled native library (tkdnd) that may not be
# present in every deployment's packaging. Everywhere this is used, the
# multi-select "Add Attachment(s)..." file-picker button remains the
# primary, always-available path -- drag-and-drop is additive, never
# the only way to attach a file.
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKINTERDND2_AVAILABLE = True
except Exception:
    TKINTERDND2_AVAILABLE = False
    DND_FILES = None
import pystray
import zipfile
import tempfile
import io
from docx import Document as DocxDocument
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from tkinter import ttk, messagebox, filedialog
import tkinter as tk
import tkinter.font as tkfont
from cryptography.fernet import Fernet

from pynetdicom import AE, evt, build_role, debug_logger
from pynetdicom.sop_class import (
    PatientRootQueryRetrieveInformationModelFind,
    StudyRootQueryRetrieveInformationModelFind,
    PatientRootQueryRetrieveInformationModelMove,
    StudyRootQueryRetrieveInformationModelMove,
)

# §fix: debug_logger() is opt-in now (see APP_SETTINGS["verbose_dicom_protocol_logging"]
# and its check in startup()), NOT called unconditionally here. Calling it
# at import time meant pynetdicom's raw, unmanaged stdout StreamHandler
# (bypassing this app's own log level / rotation / file handlers entirely)
# was permanently active for every association for the app's whole
# lifetime -- full PDU/DIMSE traces, all the time, in production, whether
# anyone was diagnosing anything or not. That's a real log-volume and
# performance cost, and it dumps low-level protocol detail (AE titles,
# UIDs, raw PDU bytes) to an unmanaged stream instead of this app's own
# audited logging path. Turn it on in Settings only when actively
# diagnosing a transfer-syntax/negotiation issue, then turn it back off.

from pydicom.uid import UID
import pydicom.uid as _pydicom_uid
from PIL import Image, ImageDraw, ImageTk

# =========================================================
# CONSTANTS
# =========================================================

APP_DIR = os.path.dirname(os.path.abspath(__file__))

LOGO_FILE = os.path.join(APP_DIR, "logo.png")
_logo_pil_cache = {"value": "unloaded"}


def _load_logo_pil():
    """Loads logo.png (transparent background, white text/mark) once and
    caches it. Returns None -- silently -- if the file isn't there, so
    every place that uses the logo just falls back to text-only."""
    if _logo_pil_cache["value"] != "unloaded":
        return _logo_pil_cache["value"]
    img = None
    try:
        if os.path.isfile(LOGO_FILE):
            img = Image.open(LOGO_FILE).convert("RGBA")
    except Exception:
        img = None
    _logo_pil_cache["value"] = img
    return img


def make_logo_ctk_image(size=28):
    """Returns a ctk.CTkImage for logo.png at the given pixel height
    (width scaled to preserve aspect ratio), or None if logo.png isn't
    present -- callers should just skip the image label in that case."""
    src = _load_logo_pil()
    if src is None:
        return None
    try:
        w, h = src.size
        target_h = size
        target_w = max(1, int(w * (target_h / h)))
        return ctk.CTkImage(light_image=src, dark_image=src, size=(target_w, target_h))
    except Exception:
        return None


KEY_FILE = "key.key"

RECEIVER_CONFIG = "rec.enc"
PUSH_CONFIG = "push.enc"          # legacy single-destination config (auto-migrated)
DESTINATIONS_CONFIG = "destinations.enc"   # new multi-destination config
NOTIFICATIONS_CONFIG = "notifications.enc"  # encrypted: may contain SMTP credentials
BANDWIDTH_CONFIG_FILE = "bandwidth_config.json"  # not secret -- just a throttle setting
VIEW_OPTIONS_FILE = "view_options.json"  # not secret -- column visibility + table density
BACKUP_DIR = "backups"  # not secret -- holds full application backup ZIPs
SCHEDULED_REPORTS_DIR = "scheduled_reports"  # not secret -- holds PDF report copies generated by the scheduled report email (8.2)
BACKUP_HISTORY_FILE = "backup_history.json"  # not secret -- backup run metadata
BACKUP_SCHEDULE_CONFIG_FILE = "backup_schedule_config.json"  # not secret -- scheduled-backup settings
LDAP_CONFIG_FILE = "ldap_config.enc"  # encrypted: holds the service-account bind password
LDAP_USERS_FILE = "ldap_users.json"  # not secret -- imported directory roster (no passwords)
ROUTING_RULES_FILE = "routing_rules.json"  # not secret -- just routing logic
TRANSFER_CHECKPOINTS_FILE = "transfer_checkpoints.json"  # resume-from-failure state (not secret)
OFFLINE_QUEUE_FILE = "offline_queue.json"  # persistent offline push queue (not secret)
DAILY_STATS_HISTORY_FILE = "daily_stats_history.json"  # not secret -- rolling day->counts summary for 7/30-day trend sparklines
TLS_CONFIG_FILE = "tls_config.json"        # TLS scaffold settings (off by default)

CSV_FILE = "dicom_worklist.csv"

RECEIVER_LOG = "receiver_errors.log"
PUSH_LOG = "push_errors.log"
AUDIT_LOG = "audit.log"
APP_LOG = "app.log"

# ---- Structured (JSON-lines) counterparts of the plain-text logs above.
# These carry the full enterprise diagnostic field set (see write_receiver_log
# / write_push_log / write_audit_log). The plain-text logs are left 100%
# unchanged for backward compatibility with anything that already parses
# them; the .jsonl files are strictly additive.
RECEIVER_LOG_JSONL = "receiver_events.jsonl"
PUSH_LOG_JSONL = "push_events.jsonl"
AUDIT_LOG_JSONL = "audit_events.jsonl"
APP_LOG_JSONL = "app_events.jsonl"

ALL_LOG_FILES = [RECEIVER_LOG, PUSH_LOG, AUDIT_LOG, APP_LOG]
ALL_LOG_FILES_JSONL = [RECEIVER_LOG_JSONL, PUSH_LOG_JSONL, AUDIT_LOG_JSONL, APP_LOG_JSONL]

# ---- Log rotation / archival (immutable audit trail -- see
# LOG_RETENTION_CONFIG_FILE / rotate_logs_if_needed()). Logs are NEVER
# permanently deleted through the UI; once a log exceeds its retention
# window or size threshold it is compressed into LOG_ARCHIVE_DIR and a
# fresh log file is started. Archives are kept forever and are browsable
# and exportable from the Logs tab.
LOG_ARCHIVE_DIR = "log_archives"
LOG_RETENTION_CONFIG_FILE = "log_retention_config.json"
DEFAULT_LOG_RETENTION_DAYS = 90
VALID_LOG_RETENTION_DAYS = [30, 60, 90, 180, 365]
LOG_ROTATION_MAX_BYTES = 5_000_000  # also rotate early if a log gets this big

SOP_INI = "sopclass.ini"

ADMIN_AUTH_FILE = "admin_auth.json"   # salt+hash for the Admin PIN. This is a
                              # *verification* secret someone types in, not
                              # config-at-rest, so it deliberately does NOT
                              # use the Fernet key mechanism above -- it's
                              # hashed (PBKDF2-HMAC) with a random salt
                              # instead, the same way a login password would be.
ADMIN_PIN_MIN_DIGITS = 6
ADMIN_PIN_PBKDF2_ITERATIONS = 200_000

OUTPUT_DIR = "received_dicoms"

DICOM_EXTENSIONS = (".dcm", ".dic", ".dicom")

DEFAULT_PUSH_WORKER_THREADS = 4
MAX_PUSH_WORKER_THREADS = 16

REFRESH_INTERVAL_MS = 2000  # background fallback worklist refresh interval
LOW_DISK_WARNING_GB = 2.0   # warn when free space drops below this

MAX_RETRY_ATTEMPTS = 3
RETRY_BASE_DELAY_SEC = 2  # exponential backoff: base * 2^attempt
PUSH_ASSOC_TIMEOUT_SEC = 30  # acse/dimse/network timeout for push associations
                              # (previously unset, so slow-but-legitimate
                              # servers had no defined grace period)

# ---------------------------------------------------------
# Document Transfer (Radiology_Report.docx / Patient_History.txt)
# ---------------------------------------------------------
# One TCP connection per patient, length-prefixed/versioned framing. Kept
# entirely separate from the DICOM association machinery above -- see
# push_patient_documents() / _handle_doc_transfer_connection().
DOC_TRANSFER_MAGIC = b"RDOC"
DOC_TRANSFER_PROTO_VERSION = 0x01


class DocTransferCancelled(Exception):
    """Raised inside push_patient_documents()'s send loop when the user
    hits Cancel on the live-progress panel -- caught specifically so the
    final status reads 'Cancelled' rather than being lumped in with a
    genuine connection failure."""
    pass


DOC_TRANSFER_CONNECT_TIMEOUT_SEC = 10  # short + fixed, separate from
                                        # get_network_timeout_sec() (tuned
                                        # for DICOM associations) so a
                                        # stalled doc transfer can never
                                        # make a whole patient push hang.
DOC_TRANSFER_CHUNK_SIZE = 65536
DOC_TRANSFER_ACCEPT_BACKLOG = 5
DOC_TRANSFER_MAX_FILE_SIZE_BYTES = 500 * 1024 * 1024  # 500 MB default cap;
                                        # overridden by the configurable
                                        # "doc_transfer_max_size_mb" app
                                        # setting wherever that's loaded --
                                        # this constant is only the fallback
                                        # for a settings file that predates
                                        # the setting (or is missing it).

DEFAULT_PUSH_CALLING_AE = "RAPPS_PUSH"  # single source of truth for the
                              # Calling AE Title used when a destination's
                              # "calling_ae" field is left blank. Both the
                              # real push AND the C-ECHO "test connection"
                              # button must use this same value -- if they
                              # ever differ, C-ECHO can succeed against a
                              # remote PACS whose Calling-AE-Title whitelist
                              # only knows the echo identity, while the real
                              # push (using a different identity) then gets
                              # its association rejected. Permissive test
                              # servers that don't check Calling AE Title at
                              # all won't expose this, which is why it can
                              # look fine in a simple sandbox but fail
                              # against a real, whitelist-enforcing PACS.

STALE_PENDING_HOURS = 24  # studies pending longer than this get highlighted

# Status constants
STATUS_PENDING = "Pending"
STATUS_SENDING = "Sending"
STATUS_SENT = "Sent"
STATUS_FAILED = "Failed"
STATUS_IMPORTED = "Imported"
STATUS_RECEIVED = "Received"
STATUS_RETRYING = "Retrying"
STATUS_QUEUED = "Queued"  # sitting in the persistent offline queue (destination unreachable)

STATUS_COLORS = {
    STATUS_PENDING: "#8b93a7",
    STATUS_SENDING: "#2f8eff",
    STATUS_SENT: "#2ecc71",
    STATUS_FAILED: "#f04747",
    STATUS_IMPORTED: "#9b59b6",
    STATUS_RECEIVED: "#f1c40f",
    STATUS_RETRYING: "#e67e22",
    STATUS_QUEUED: "#3498db",
}

# Display-only pseudo-status (2.4): never stored as d["status"] itself (that
# stays STATUS_SENDING so existing state-machine logic elsewhere is
# untouched) -- populate_tree() swaps in this tag/label purely for rows
# where push_single_patient found prior checkpoint progress to resume.
STATUS_RESUMING_DISPLAY = "Resuming"
STATUS_COLORS[STATUS_RESUMING_DISPLAY] = "#17a2b8"

STALE_HIGHLIGHT_TAG = "stale"
STALE_HIGHLIGHT_COLOR = "#ff5555"

SEARCH_MATCH_HIGHLIGHT_TAG = "search_match"
SEARCH_MATCH_HIGHLIGHT_BG = "#3a3510"  # subtle warm highlight, distinct from stale red

# =========================================================
# APP LOGGER (internal diagnostics, separate from receiver/push logs)
# =========================================================

def _build_logger():
    logger = logging.getLogger("rapps")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = RotatingFileHandler(APP_LOG, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(handler)
    return logger


app_logger = _build_logger()


def log_exception(context):
    """Call inside an except block to record a full traceback to app.log."""
    app_logger.error("%s\n%s", context, traceback.format_exc())

# =========================================================
# ATOMIC PERSISTENCE (§3.2 fix)
# =========================================================
# Every function in this file that persists application state (worklist
# CSV, encrypted .enc configs, routing rules, checkpoints, offline queue,
# app settings, admin PIN, TLS config, backup history, LDAP config, ...)
# used to open its target file directly in "w" mode and write straight
# into it. A crash, kill -9, power loss, or an antivirus file lock
# mid-write leaves that file truncated or corrupt -- and several of these
# (the worklist CSV, destinations, routing rules, offline queue, transfer
# checkpoints) are exactly the files this app cannot safely reload from
# anywhere else.
#
# atomic_open_for_write()/atomic_write() below are the one shared
# primitive every persistence function should route through instead:
# write to a temp file in the SAME directory as the target (so the final
# rename is on the same filesystem/volume and therefore atomic), then
# os.replace() it into place only once the write has fully succeeded.
# If anything raises before that point, the temp file is discarded and
# the original file at `path` is left completely untouched -- there is
# no window where `path` itself is partially written.
#
# This generalizes the one place in the codebase that already did this
# correctly (_doc_transfer_receive_one_file(), which writes to
# "<path>.part" then os.replace()s it in) into a shared helper used
# everywhere persisted state is saved.


@contextlib.contextmanager
def atomic_open_for_write(path, mode="w", encoding="utf-8", newline=None):
    """Context manager handing back a writable file object backed by a
    temp file in the same directory as `path`. If the `with` block runs
    to completion without raising, the temp file is flushed, fsync()'d,
    closed, and atomically os.replace()d into place as `path`. If the
    block raises (or the flush/replace itself fails), the temp file is
    discarded and `path` is left exactly as it was before the call.

    This exists as a context manager -- not just a whole-string writer --
    because some callers (csv.writer, json.dump) need a real file object
    to write incrementally into rather than a single pre-built string.
    See atomic_write() below for the common "I already have the full
    string/bytes" case.

    `mode` must be "w" or "wb". `encoding`/`newline` are only meaningful
    for text mode ("w") and are passed straight through to os.fdopen(),
    same as they'd be passed to a normal open() call.
    """
    if mode not in ("w", "wb"):
        raise ValueError(f"atomic_open_for_write only supports mode='w' or 'wb', got {mode!r}")
    directory = os.path.dirname(os.path.abspath(path)) or "."
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".atomicpart", dir=directory)
    is_binary = mode == "wb"
    f = os.fdopen(fd, mode) if is_binary else os.fdopen(fd, mode, encoding=encoding, newline=newline)
    try:
        yield f
        f.flush()
        os.fsync(f.fileno())
        f.close()
        os.replace(tmp_path, path)
    except Exception:
        try:
            f.close()
        except Exception:
            pass
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def atomic_write(path, data, binary=False, encoding="utf-8"):
    """Convenience wrapper around atomic_open_for_write() for the common
    case of already having the full content to write as one string (or
    bytes, if binary=True). Same all-or-nothing guarantee: either `path`
    ends up containing exactly `data`, or -- on any failure -- `path` is
    left completely unchanged. Raises on failure; callers are responsible
    for catching and logging via log_exception(), same as every other
    persistence function in this file already does."""
    mode = "wb" if binary else "w"
    with atomic_open_for_write(path, mode=mode, encoding=encoding) as f:
        f.write(data)

# =========================================================
# GLOBAL STATE
# =========================================================

patient_data = {}          # pid -> dict(...)
data_lock = threading.RLock()

# Bumped by every function that mutates patient_data. Lets populate_tree()
# skip a full Treeview rebuild when nothing has actually changed, instead of
# unconditionally rebuilding on every 2s periodic_refresh() tick.
_patient_data_version = [0]


def _bump_data_version():
    _patient_data_version[0] += 1

ui_event_queue = queue.Queue()   # thread-safe events consumed by the GUI loop

push_job = {
    "running": False,
    "total_images": 0,
    "sent_images": 0,      # successfully sent (accurate, not "attempted")
    "attempted_images": 0,
    "stop_flag": False,
    "started_at": None,
    "current_pid": "",       # patient currently being sent (for the status line)
    "current_dest": "",      # "AE@ip:port" currently being sent to
}

# 2.4 -- pid -> "sent/total" string, populated only while push_single_patient
# is actively resuming a patient that had prior checkpoint progress for this
# destination. Cleared once that attempt finishes (success or failure).
# Read by populate_tree() (row tag/label) and refresh_pusher_monitoring()
# (global status badge) -- display-only, never written into patient_data.
_push_resume_progress = {}

receiver_state = {
    "server_ae": None,
    "thread": None,
    "running": False,
    "autoroute": True,
}

# =========================================================
# ENTERPRISE DASHBOARD STATE (landing-page tab)
# =========================================================

APP_START_TIME = time.time()
reports_last_path = {"value": None}  # last-generated PDF report path, for Print/Email buttons
reports_last_meta = {"label": None, "start_dt": None, "end_dt": None, "generated_at": None}

# SOP Class UIDs treated as "documents" (encapsulated PDFs/CDA/mesh docs)
# for the "Documents Received Today" dashboard stat.
DOCUMENT_SOP_CLASS_UIDS = {
    "1.2.840.10008.5.1.4.1.1.104.1",  # EncapsulatedPDFStorage
    "1.2.840.10008.5.1.4.1.1.104.2",  # EncapsulatedCDAStorage
    "1.2.840.10008.5.1.4.1.1.104.3",  # EncapsulatedSTLStorage
    "1.2.840.10008.5.1.4.1.1.104.4",  # EncapsulatedOBJStorage
    "1.2.840.10008.5.1.4.1.1.104.5",  # EncapsulatedMTLStorage
}

_dashboard_stats_lock = threading.RLock()

# In-memory "today" counters. Intentionally in-memory (like push_job's
# throughput counters elsewhere in this file) rather than re-derived from
# the log files on every tick, so the dashboard stays cheap to refresh
# every couple of seconds. They roll over automatically at local midnight.
daily_stats = {
    "date": datetime.date.today().isoformat(),
    "studies_received_uids": set(),
    "images_received": 0,
    "documents_received": 0,
    "studies_sent_uids": set(),
    "images_sent": 0,
    "failed_transfers": 0,
}


DAILY_STATS_HISTORY_MAX_DAYS = 90  # keep enough for a 30-day view with room to spare
_daily_stats_history_cache = {"value": None}


def load_daily_stats_history():
    """List of {"date": "YYYY-MM-DD", "studies_received", "images_received",
    "documents_received", "studies_sent", "images_sent", "failed_transfers"}
    dicts, oldest first. Same load/save-with-cache pattern as
    load_backup_schedule() etc."""
    if _daily_stats_history_cache["value"] is not None:
        return _daily_stats_history_cache["value"]
    history = []
    try:
        if os.path.exists(DAILY_STATS_HISTORY_FILE):
            with open(DAILY_STATS_HISTORY_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, list):
                history = loaded
    except Exception:
        log_exception("Failed to load daily_stats_history.json")
    _daily_stats_history_cache["value"] = history
    return history


def save_daily_stats_history(history):
    try:
        atomic_write(DAILY_STATS_HISTORY_FILE, json.dumps(history, indent=2))
    except Exception:
        log_exception("Failed to save daily_stats_history.json")
    _daily_stats_history_cache["value"] = history


def _append_daily_stats_to_history(prior_date, prior_snapshot):
    """Called from _roll_daily_stats_if_needed() right before it resets
    daily_stats for the new day -- appends the day that just ended."""
    history = load_daily_stats_history()
    history = [h for h in history if h.get("date") != prior_date]  # avoid dupes on restart mid-day
    history.append({
        "date": prior_date,
        "studies_received": len(prior_snapshot["studies_received_uids"]),
        "images_received": prior_snapshot["images_received"],
        "documents_received": prior_snapshot["documents_received"],
        "studies_sent": len(prior_snapshot["studies_sent_uids"]),
        "images_sent": prior_snapshot["images_sent"],
        "failed_transfers": prior_snapshot["failed_transfers"],
    })
    history.sort(key=lambda h: h.get("date", ""))
    if len(history) > DAILY_STATS_HISTORY_MAX_DAYS:
        history = history[-DAILY_STATS_HISTORY_MAX_DAYS:]
    save_daily_stats_history(history)


def get_doc_transfer_daily_trend(days=7):
    """B.5: docs-delivered-per-day and doc-transfer failure-rate over the
    last N days. Computed directly from the existing push-log JSONL
    (event_type == "DOC-TRANSFER") rather than the persisted
    daily_stats_history.json -- so it works retroactively over
    already-logged days with no rollover/migration needed, consistent
    with how B.1/B.2/B.4/B.6 all derive from the same log stream."""
    trend = []
    today = datetime.date.today()
    for i in range(days - 1, -1, -1):
        day = today - datetime.timedelta(days=i)
        start = datetime.datetime.combine(day, datetime.time.min)
        end = datetime.datetime.combine(day, datetime.time.max)
        stats = compute_doc_transfer_stats(start, end)
        trend.append({
            "date": day.isoformat(),
            "docs_delivered": stats["docs_sent_total"],
            "failure_rate_pct": (stats["failures"] / stats["attempts"] * 100.0) if stats["attempts"] else 0.0,
        })
    return trend


def get_daily_stats_trend(days=7):
    """Returns the last `days` entries from the persisted history PLUS
    today's in-progress daily_stats as the final point, for
    _draw_sparkline() to render a 7/30-day view instead of only the
    ~2-minute in-session dash_history window."""
    history = load_daily_stats_history()
    today = datetime.date.today().isoformat()
    trend = [h for h in history if h.get("date") != today][-max(days - 1, 0):]
    with _dashboard_stats_lock:
        trend = trend + [{
            "date": today,
            "studies_received": len(daily_stats["studies_received_uids"]),
            "images_received": daily_stats["images_received"],
            "documents_received": daily_stats["documents_received"],
            "studies_sent": len(daily_stats["studies_sent_uids"]),
            "images_sent": daily_stats["images_sent"],
            "failed_transfers": daily_stats["failed_transfers"],
        }]
    return trend[-days:]


def _roll_daily_stats_if_needed():
    today = datetime.date.today().isoformat()
    prior_date = None
    prior_snapshot = None
    with _dashboard_stats_lock:
        if daily_stats["date"] != today:
            prior_date = daily_stats["date"]
            prior_snapshot = {
                "studies_received_uids": daily_stats["studies_received_uids"],
                "images_received": daily_stats["images_received"],
                "documents_received": daily_stats["documents_received"],
                "studies_sent_uids": daily_stats["studies_sent_uids"],
                "images_sent": daily_stats["images_sent"],
                "failed_transfers": daily_stats["failed_transfers"],
            }
            daily_stats["date"] = today
            daily_stats["studies_received_uids"] = set()
            daily_stats["images_received"] = 0
            daily_stats["documents_received"] = 0
            daily_stats["studies_sent_uids"] = set()
            daily_stats["images_sent"] = 0
            daily_stats["failed_transfers"] = 0
    if prior_snapshot is not None:
        try:
            _append_daily_stats_to_history(prior_date, prior_snapshot)
        except Exception:
            log_exception("Failed to append prior day to daily_stats_history.json")


def record_receive_stat(study_uid, sop_class_uid):
    _roll_daily_stats_if_needed()
    with _dashboard_stats_lock:
        daily_stats["images_received"] += 1
        if study_uid:
            daily_stats["studies_received_uids"].add(study_uid)
        if sop_class_uid in DOCUMENT_SOP_CLASS_UIDS:
            daily_stats["documents_received"] += 1


def record_push_stat(study_uid, images_sent, images_failed):
    _roll_daily_stats_if_needed()
    with _dashboard_stats_lock:
        daily_stats["images_sent"] += images_sent
        daily_stats["failed_transfers"] += images_failed
        if images_sent and study_uid:
            daily_stats["studies_sent_uids"].add(study_uid)


# Rolling history buffers for the Dashboard's live graphs. Each tuple is
# (unix_timestamp, value). ~2 minutes of history at the existing 2s
# REFRESH_INTERVAL_MS tick rate.
DASHBOARD_HISTORY_LEN = 60
dash_history = {
    "studies_received": deque(maxlen=DASHBOARD_HISTORY_LEN),
    "studies_sent": deque(maxlen=DASHBOARD_HISTORY_LEN),
    "failed_transfers": deque(maxlen=DASHBOARD_HISTORY_LEN),
    "queue_size": deque(maxlen=DASHBOARD_HISTORY_LEN),
    "network_kbps": deque(maxlen=DASHBOARD_HISTORY_LEN),
}

# Cache of last-known per-destination health, refreshed on demand by the
# Dashboard (and by the future PACS Health Monitor tab, which will reuse
# this same cache). Keyed by destination name.
destination_health_cache = {}

# Network throughput sampling baseline (psutil counters + timestamp of the
# last sample) so we can compute a KB/s rate between dashboard ticks.
_net_io_baseline = {"time": None, "bytes_sent": None, "bytes_recv": None}
_last_cpu_percent = [0.0]


# ---- Admin/User role-split state --------------------------------------
# admin_session["unlocked"]: True once the correct PIN has been entered this
#   run; persists until Lock is clicked or the app closes (no idle-timeout
#   auto-lock in this pass -- TODO: add idle-timeout auto-lock in a future
#   version).
# current_view["role"]: which tab set is ACTUALLY on screen right now --
#   "admin" or "user". An unlocked Admin can preview the User tab set via
#   "Switch to User View" without losing admin_session["unlocked"], so these
#   two are tracked separately.
admin_session = {"unlocked": False}
current_view = {"role": "user"}
current_identity = {"username": None, "display_name": None, "source": "local", "ldap_role": None}
admin_tabs_active = {"value": False}  # whether the admin-only tabs currently exist in the tabview

# D.3: last-activity timestamp for the admin idle-timeout auto-lock.
# Updated by _record_ui_activity(), bound to Motion/Key/Button events once
# the main window exists (see app.bind_all calls near the end of the file).
# Checked once per periodic_refresh() tick in check_admin_idle_timeout().
_last_ui_activity_time = [time.time()]


def _record_ui_activity(event=None):
    _last_ui_activity_time[0] = time.time()


def check_admin_idle_timeout():
    """D.3: flips admin_session["unlocked"] back to False and returns the
    view to the Lock screen after admin_idle_timeout_min minutes of
    inactivity (0 = disabled). This only gates the UI -- an in-flight
    push/receive job is never interrupted, since those run independently
    on their own worker threads regardless of admin_session state."""
    if not admin_session["unlocked"]:
        return
    try:
        timeout_min = int(APP_SETTINGS.get("admin_idle_timeout_min", 15) or 0)
    except Exception:
        timeout_min = 0
    if timeout_min <= 0:
        return
    idle_sec = time.time() - _last_ui_activity_time[0]
    if idle_sec >= timeout_min * 60:
        do_admin_lock()

app_shutdown_event = threading.Event()
DEFAULT_SOP_CLASSES = {
    # Originally curated to exactly 128 entries to match DICOM's hard
    # per-association presentation-context limit -- but that limit only
    # actually applies to add_requested_context() (a single outgoing push
    # association, see cap_sop_list()), not to the receiver's
    # add_supported_context() catalog, which has no such limit (verified
    # directly against pynetdicom). The receiver is therefore NOT capped
    # to this list's length; cap_sop_list() is still applied on the
    # pusher's push_single_patient() path, where the true 128-per-
    # association limit does apply. Everything commonly seen in general
    # radiology/PACS traffic (CT/MR/US/XA/NM/PET/RT/SR/waveforms/
    # presentation states/secondary capture/etc.) is kept; only the
    # least-used entry (ImplantTemplateGroupStorage, a niche implant-
    # planning SOP) was ever dropped, to make room back when this WAS
    # capped at 128. Add it back via the in-GUI SOP editor if your site
    # actually needs it.
    "ComputedRadiographyImageStorage": "1.2.840.10008.5.1.4.1.1.1",
    "DigitalXRayImageStorageForPresentation": "1.2.840.10008.5.1.4.1.1.1.1",
    "DigitalXRayImageStorageForProcessing": "1.2.840.10008.5.1.4.1.1.1.1.1",
    "DigitalMammographyXRayImageStorageForPresentation": "1.2.840.10008.5.1.4.1.1.1.2",
    "DigitalMammographyXRayImageStorageForProcessing": "1.2.840.10008.5.1.4.1.1.1.2.1",
    "DigitalIntraOralXRayImageStorageForPresentation": "1.2.840.10008.5.1.4.1.1.1.3",
    "DigitalIntraOralXRayImageStorageForProcessing": "1.2.840.10008.5.1.4.1.1.1.3.1",
    "CTImageStorage": "1.2.840.10008.5.1.4.1.1.2",
    "EnhancedCTImageStorage": "1.2.840.10008.5.1.4.1.1.2.1",
    "LegacyConvertedEnhancedCTImageStorage": "1.2.840.10008.5.1.4.1.1.2.2",
    "UltrasoundMultiFrameImageStorage": "1.2.840.10008.5.1.4.1.1.3.1",
    "MRImageStorage": "1.2.840.10008.5.1.4.1.1.4",
    "EnhancedMRImageStorage": "1.2.840.10008.5.1.4.1.1.4.1",
    "MRSpectroscopyStorage": "1.2.840.10008.5.1.4.1.1.4.2",
    "EnhancedMRColorImageStorage": "1.2.840.10008.5.1.4.1.1.4.3",
    "LegacyConvertedEnhancedMRImageStorage": "1.2.840.10008.5.1.4.1.1.4.4",
    "UltrasoundImageStorage": "1.2.840.10008.5.1.4.1.1.6.1",
    "EnhancedUSVolumeStorage": "1.2.840.10008.5.1.4.1.1.6.2",
    "SecondaryCaptureImageStorage": "1.2.840.10008.5.1.4.1.1.7",
    "MultiFrameSingleBitSecondaryCaptureImageStorage": "1.2.840.10008.5.1.4.1.1.7.1",
    "MultiFrameGrayscaleByteSecondaryCaptureImageStorage": "1.2.840.10008.5.1.4.1.1.7.2",
    "MultiFrameGrayscaleWordSecondaryCaptureImageStorage": "1.2.840.10008.5.1.4.1.1.7.3",
    "MultiFrameTrueColorSecondaryCaptureImageStorage": "1.2.840.10008.5.1.4.1.1.7.4",
    "TwelveLeadECGWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.1.1",
    "GeneralECGWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.1.2",
    "AmbulatoryECGWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.1.3",
    "HemodynamicWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.2.1",
    "CardiacElectrophysiologyWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.3.1",
    "BasicVoiceAudioWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.4.1",
    "GeneralAudioWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.4.2",
    "ArterialPulseWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.5.1",
    "RespiratoryWaveformStorage": "1.2.840.10008.5.1.4.1.1.9.6.1",
    "GrayscaleSoftcopyPresentationStateStorage": "1.2.840.10008.5.1.4.1.1.11.1",
    "ColorSoftcopyPresentationStateStorage": "1.2.840.10008.5.1.4.1.1.11.2",
    "PseudoColorSoftcopyPresentationStateStorage": "1.2.840.10008.5.1.4.1.1.11.3",
    "BlendingSoftcopyPresentationStateStorage": "1.2.840.10008.5.1.4.1.1.11.4",
    "XAXRFGrayscaleSoftcopyPresentationStateStorage": "1.2.840.10008.5.1.4.1.1.11.5",
    "XRayAngiographicImageStorage": "1.2.840.10008.5.1.4.1.1.12.1",
    "EnhancedXAImageStorage": "1.2.840.10008.5.1.4.1.1.12.1.1",
    "XRayRadiofluoroscopicImageStorage": "1.2.840.10008.5.1.4.1.1.12.2",
    "EnhancedXRFImageStorage": "1.2.840.10008.5.1.4.1.1.12.2.1",
    "XRay3DAngiographicImageStorage": "1.2.840.10008.5.1.4.1.1.13.1.1",
    "XRay3DCraniofacialImageStorage": "1.2.840.10008.5.1.4.1.1.13.1.2",
    "BreastTomosynthesisImageStorage": "1.2.840.10008.5.1.4.1.1.13.1.3",
    "BreastProjectionXRayImageStorageForPresentation": "1.2.840.10008.5.1.4.1.1.13.1.4",
    "BreastProjectionXRayImageStorageForProcessing": "1.2.840.10008.5.1.4.1.1.13.1.5",
    "IntravascularOpticalCoherenceTomographyImageStorageForPresentation": "1.2.840.10008.5.1.4.1.1.14.1",
    "IntravascularOpticalCoherenceTomographyImageStorageForProcessing": "1.2.840.10008.5.1.4.1.1.14.2",
    "NuclearMedicineImageStorage": "1.2.840.10008.5.1.4.1.1.20",
    "RawDataStorage": "1.2.840.10008.5.1.4.1.1.66",
    "SpatialRegistrationStorage": "1.2.840.10008.5.1.4.1.1.66.1",
    "SpatialFiducialsStorage": "1.2.840.10008.5.1.4.1.1.66.2",
    "DeformableSpatialRegistrationStorage": "1.2.840.10008.5.1.4.1.1.66.3",
    "SegmentationStorage": "1.2.840.10008.5.1.4.1.1.66.4",
    "SurfaceSegmentationStorage": "1.2.840.10008.5.1.4.1.1.66.5",
    "RealWorldValueMappingStorage": "1.2.840.10008.5.1.4.1.1.67",
    "SurfaceScanMeshStorage": "1.2.840.10008.5.1.4.1.1.68.1",
    "SurfaceScanPointCloudStorage": "1.2.840.10008.5.1.4.1.1.68.2",
    "VLEndoscopicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.1",
    "VideoEndoscopicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.1.1",
    "VLMicroscopicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.2",
    "VideoMicroscopicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.2.1",
    "VLSlideCoordinatesMicroscopicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.3",
    "VLPhotographicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.4",
    "VideoPhotographicImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.4.1",
    "OphthalmicPhotography8BitImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.1",
    "OphthalmicPhotography16BitImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.2",
    "StereometricRelationshipStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.3",
    "OphthalmicTomographyImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.4",
    "WideFieldOphthalmicPhotographyStereographicProjectionImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.5",
    "WideFieldOphthalmicPhotography3DCoordinatesImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.6",
    "OphthalmicOpticalCoherenceTomographyEnFaceImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.7",
    "OphthalmicOpticalCoherenceTomographyBscanVolumeAnalysisStorage": "1.2.840.10008.5.1.4.1.1.77.1.5.8",
    "VLWholeSlideMicroscopyImageStorage": "1.2.840.10008.5.1.4.1.1.77.1.6",
    "LensometryMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.78.1",
    "AutorefractionMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.78.2",
    "KeratometryMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.78.3",
    "SubjectiveRefractionMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.78.4",
    "VisualAcuityMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.78.5",
    "SpectaclePrescriptionReportStorage": "1.2.840.10008.5.1.4.1.1.78.6",
    "OphthalmicAxialMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.78.7",
    "IntraocularLensCalculationsStorage": "1.2.840.10008.5.1.4.1.1.78.8",
    "MacularGridThicknessAndVolumeReportStorage": "1.2.840.10008.5.1.4.1.1.79.1",
    "OphthalmicVisualFieldStaticPerimetryMeasurementsStorage": "1.2.840.10008.5.1.4.1.1.80.1",
    "OphthalmicThicknessMapStorage": "1.2.840.10008.5.1.4.1.1.81.1",
    "CornealTopographyMapStorage": "1.2.840.10008.5.1.4.1.1.82.1",
    "BasicTextSRStorage": "1.2.840.10008.5.1.4.1.1.88.11",
    "EnhancedSRStorage": "1.2.840.10008.5.1.4.1.1.88.22",
    "ComprehensiveSRStorage": "1.2.840.10008.5.1.4.1.1.88.33",
    "Comprehensive3DSRStorage": "1.2.840.10008.5.1.4.1.1.88.34",
    "ExtensibleSRStorage": "1.2.840.10008.5.1.4.1.1.88.35",
    "ProcedureLogStorage": "1.2.840.10008.5.1.4.1.1.88.40",
    "MammographyCADSRStorage": "1.2.840.10008.5.1.4.1.1.88.50",
    "KeyObjectSelectionDocumentStorage": "1.2.840.10008.5.1.4.1.1.88.59",
    "ChestCADSRStorage": "1.2.840.10008.5.1.4.1.1.88.65",
    "XRayRadiationDoseSRStorage": "1.2.840.10008.5.1.4.1.1.88.67",
    "RadiopharmaceuticalRadiationDoseSRStorage": "1.2.840.10008.5.1.4.1.1.88.68",
    "ColonCADSRStorage": "1.2.840.10008.5.1.4.1.1.88.69",
    "ImplantationPlanSRDocumentStorage": "1.2.840.10008.5.1.4.1.1.88.70",
    "AcquisitionContextSRStorage": "1.2.840.10008.5.1.4.1.1.88.71",
    "SimplifiedAdultEchoSRStorage": "1.2.840.10008.5.1.4.1.1.88.72",
    "PatientRadiationDoseSRStorage": "1.2.840.10008.5.1.4.1.1.88.73",
    "PlannedImagingAgentAdministrationSRStorage": "1.2.840.10008.5.1.4.1.1.88.74",
    "PerformedImagingAgentAdministrationSRStorage": "1.2.840.10008.5.1.4.1.1.88.75",
    "ContentAssessmentResultsStorage": "1.2.840.10008.5.1.4.1.1.90.1",
    "EncapsulatedPDFStorage": "1.2.840.10008.5.1.4.1.1.104.1",
    "EncapsulatedCDAStorage": "1.2.840.10008.5.1.4.1.1.104.2",
    "EncapsulatedSTLStorage": "1.2.840.10008.5.1.4.1.1.104.3",
    "EncapsulatedOBJStorage": "1.2.840.10008.5.1.4.1.1.104.4",
    "EncapsulatedMTLStorage": "1.2.840.10008.5.1.4.1.1.104.5",
    "PositronEmissionTomographyImageStorage": "1.2.840.10008.5.1.4.1.1.128",
    "EnhancedPETImageStorage": "1.2.840.10008.5.1.4.1.1.130",
    "LegacyConvertedEnhancedPETImageStorage": "1.2.840.10008.5.1.4.1.1.128.1",
    "BasicStructuredDisplayStorage": "1.2.840.10008.5.1.4.1.1.131",
    "CTPerformedProcedureProtocolStorage": "1.2.840.10008.5.1.4.1.1.200.2",
    "RTImageStorage": "1.2.840.10008.5.1.4.1.1.481.1",
    "RTDoseStorage": "1.2.840.10008.5.1.4.1.1.481.2",
    "RTStructureSetStorage": "1.2.840.10008.5.1.4.1.1.481.3",
    "RTBeamsTreatmentRecordStorage": "1.2.840.10008.5.1.4.1.1.481.4",
    "RTPlanStorage": "1.2.840.10008.5.1.4.1.1.481.5",
    "RTBrachyTreatmentRecordStorage": "1.2.840.10008.5.1.4.1.1.481.6",
    "RTTreatmentSummaryRecordStorage": "1.2.840.10008.5.1.4.1.1.481.7",
    "RTIonPlanStorage": "1.2.840.10008.5.1.4.1.1.481.8",
    "RTIonBeamsTreatmentRecordStorage": "1.2.840.10008.5.1.4.1.1.481.9",
    "RTBeamsDeliveryInstructionStorage": "1.2.840.10008.5.1.4.34.7",
    "RTBrachyApplicationSetupDeliveryInstructionStorage": "1.2.840.10008.5.1.4.34.10",
    "GenericImplantTemplateStorage": "1.2.840.10008.5.1.4.43.1",
    "ImplantAssemblyTemplateStorage": "1.2.840.10008.5.1.4.44.1",

}

# SOP_CLASS_AUDIT_EXCLUDED: entries present in the uploaded sopclass_copy.ini
# that were deliberately NOT merged into DEFAULT_SOP_CLASSES above, because
# they either aren't Storage SOP Classes (this app is Storage-only -- it has
# no C-FIND/C-MOVE/C-GET SCP handlers, so accepting these contexts would
# negotiate successfully but never actually work) or couldn't be verified
# against the current, finalized DICOM standard in the time available for
# this audit:
#   - Not Storage SOP Classes (Query/Retrieve, Workflow, or Print/Media
#     Management -- would be redundant/non-functional here): verification /
#     verification_ctx (already handled separately, see VERIFICATION_SOP_CLASS
#     below -- adding it here would be a duplicate), presentationlut,
#     mediacreationmanagement, displaysystem / displaysysteminstance,
#     compositeinstancerootretrievemove/get, compositeinstanceretrievewithout
#     bulkdataget, colorpaletteinformationmodelfind/get/move, generalrelevant
#     patientinformationquery, cardiacrelevantpatientinformationquery,
#     substanceapprovalquery, unifiedprocedurestepevent.
#   - Confirmed DRAFT status, not yet finalized (verified against
#     dicomstandard.org's own supplement PDF, which explicitly labels it
#     "- Draft -" with a placeholder UID in one revision): labelmap
#     segmentationstorage (Sup 243). heightmapsegmentationstorage is the
#     same supplement family and was excluded for the same reason.
#   - Could not be independently verified as finalized (current text vs.
#     draft/proposed supplement) in the time available -- recommend
#     checking PS3.6 Annex A directly before adding: the RT Radiation Set
#     family (rtradiationsetstorage/rtradiationrecordsetstorage/
#     rtradiationsalvagerecordstorage/rtradiationsetdeliveryinstructionstorage),
#     enhancedrtimagestorage/enhancedcontinuousrtimagestorage,
#     photoacousticimagestorage, tractographyresultsstorage,
#     dermoscopicphotographyimagestorage, enhancedxrayradiationdosesrstorage,
#     waveformannotationsrstorage, general32bitecgwaveformstorage,
#     waveformpresentationstatestorage/waveformacquisitionpresentationstatestorage,
#     multichannelrespiratorywaveformstorage, electromyogramwaveformstorage,
#     electrooculogramwaveformstorage, microscopybulksimpleannotationsstorage,
#     multiplevolumerenderingvolumetricpresentationstatestorage,
#     grayscaleplanarmprvolumetricpresentationstatestorage,
#     compositingplanarmprvolumetricpresentationstatestorage,
#     variablemodalitylutsoftcopypresentationstagestorage,
#     substanceadministrationlogginginstance.
# None of these were REMOVED from a working list -- they were simply never
# added, so this is purely additive risk management, not a regression.

DEFAULT_TRANSFER_SYNTAXES = {
    "ImplicitVRLittleEndian": "1.2.840.10008.1.2",
    "ExplicitVRLittleEndian": "1.2.840.10008.1.2.1",
    "DeflatedExplicitVRLittleEndian": "1.2.840.10008.1.2.1.99",
    "ExplicitVRBigEndian": "1.2.840.10008.1.2.2",
    "JPEGBaseline8Bit": "1.2.840.10008.1.2.4.50",
    "JPEGExtended12Bit": "1.2.840.10008.1.2.4.51",
    "JPEGLossless": "1.2.840.10008.1.2.4.57",
    "JPEGLosslessSV1": "1.2.840.10008.1.2.4.70",
    "JPEGLSLossless": "1.2.840.10008.1.2.4.80",
    "JPEGLSNearLossless": "1.2.840.10008.1.2.4.81",
    "JPEG2000Lossless": "1.2.840.10008.1.2.4.90",
    "JPEG2000": "1.2.840.10008.1.2.4.91",
    "JPEG2000MCLossless": "1.2.840.10008.1.2.4.92",
    "JPEG2000MC": "1.2.840.10008.1.2.4.93",
    "RLELossless": "1.2.840.10008.1.2.5",
    "HighThroughputJPEG2000Lossless": "1.2.840.10008.1.2.4.201",
    "HighThroughputJPEG2000LosslessRPCL": "1.2.840.10008.1.2.4.202",
    "HighThroughputJPEG2000": "1.2.840.10008.1.2.4.203",
    # Added during the sopclass_copy.ini audit -- verified directly
    # against pydicom.uid's own constants (HEVCMP51, HEVCM10P51) rather
    # than relying on the ini alone. HEVC was completely absent before
    # despite being explicitly required.
    "HEVCMainProfileLevel51": "1.2.840.10008.1.2.4.107",
    "HEVCMain10ProfileLevel51": "1.2.840.10008.1.2.4.108",
}


# Comprehensive, always-current list of every Transfer Syntax UID pydicom
# itself knows about (Implicit/Explicit VR, JPEG family, JPEG-LS, JPEG2000
# family incl. HTJ2K, RLE, MPEG2/MPEG-4/HEVC video, SMPTE ST 2110, etc).
# This is pulled straight from pydicom.uid.AllTransferSyntaxes rather than
# hand-maintained, so it stays correct as pydicom adds new transfer syntax
# UIDs -- unlike DEFAULT_TRANSFER_SYNTAXES/sopclass.ini (which an admin can
# freely edit/trim), this list is used as the unconditional "propose/accept
# everything" fallback so an association is never rejected, or a peer's
# native transfer syntax never offered, purely because sopclass.ini happens
# to be missing an entry.
ALL_TRANSFER_SYNTAXES = list(dict.fromkeys(UID(str(ts)) for ts in _pydicom_uid.AllTransferSyntaxes))


_sop_ini_migration_done = {"value": False}


def _migrate_sop_ini():
    """Adds any SOP Class / Transfer Syntax UID present in
    DEFAULT_SOP_CLASSES / DEFAULT_TRANSFER_SYNTAXES but missing (by UID,
    not by key name -- an admin may have renamed a key) from an existing
    sopclass.ini on disk. This is what makes source-level improvements to
    those two dicts (e.g. adding HEVC)
    actually reach a machine that already has a sopclass.ini file --
    without this, create_sop_ini() below would see the file already
    exists and never touch it again, silently stranding every such
    improvement on any machine that isn't a fresh install. Purely
    additive: never removes or overwrites an existing entry, so an
    admin's hand-edits (including deliberate removals) always survive.
    Runs at most once per process -- create_sop_ini() is called on
    every load_extra_sops()/load_transfer_syntaxes(), which would
    otherwise mean re-reading and rewriting this file constantly."""
    if _sop_ini_migration_done["value"]:
        return
    _sop_ini_migration_done["value"] = True
    try:
        config = configparser.ConfigParser()
        config.read(SOP_INI)
        changed = False

        if "SOP_CLASSES" not in config:
            config["SOP_CLASSES"] = {}
        existing_uids = {v.strip() for v in config["SOP_CLASSES"].values()}
        for name, uid in DEFAULT_SOP_CLASSES.items():
            if uid not in existing_uids:
                key = name.lower()
                suffix = 2
                while key in config["SOP_CLASSES"]:
                    key = f"{name.lower()}_{suffix}"
                    suffix += 1
                config["SOP_CLASSES"][key] = uid
                changed = True

        if "TRANSFER_SYNTAXES" not in config:
            config["TRANSFER_SYNTAXES"] = {}
        existing_ts_uids = {v.strip() for v in config["TRANSFER_SYNTAXES"].values()}
        for name, uid in DEFAULT_TRANSFER_SYNTAXES.items():
            if uid not in existing_ts_uids:
                key = name.lower()
                suffix = 2
                while key in config["TRANSFER_SYNTAXES"]:
                    key = f"{name.lower()}_{suffix}"
                    suffix += 1
                config["TRANSFER_SYNTAXES"][key] = uid
                changed = True

        if changed:
            with atomic_open_for_write(SOP_INI, mode="w") as f:
                config.write(f)
            write_audit_log("SOP-INI-MIGRATED",
                            "Added new default SOP Classes/Transfer Syntaxes to existing sopclass.ini")
    except Exception:
        log_exception("Failed to migrate sopclass.ini with new defaults")


def create_sop_ini():
    """Generate a full sopclass.ini (SOP Classes + Transfer Syntaxes) the
    first time the app runs. Both the Receiver and the Pusher read this
    SAME file, so negotiated contexts always match on both sides. Users
    may hand-edit this file later to add/remove entries.

    On every run AFTER the first, _migrate_sop_ini() adds anything new
    to DEFAULT_SOP_CLASSES/DEFAULT_TRANSFER_SYNTAXES since the file was
    created -- see its docstring for why that matters."""
    if os.path.exists(SOP_INI):
        _migrate_sop_ini()
        return
    config = configparser.ConfigParser()
    config["SOP_CLASSES"] = DEFAULT_SOP_CLASSES
    config["TRANSFER_SYNTAXES"] = DEFAULT_TRANSFER_SYNTAXES
    with atomic_open_for_write(SOP_INI, mode="w") as f:
        config.write(f)


def load_extra_sops():
    """Load every SOP Class UID listed under [SOP_CLASSES] in sopclass.ini."""
    create_sop_ini()
    config = configparser.ConfigParser()
    config.read(SOP_INI)
    sops = []
    if "SOP_CLASSES" in config:
        for name, uid in config["SOP_CLASSES"].items():
            try:
                sops.append(UID(uid.strip()))
            except Exception:
                # §3.3: this used to fail silently, so a typo'd UID in
                # sopclass.ini meant that SOP class was quietly never
                # accepted with zero indication why.
                app_logger.warning("Skipping invalid SOP Class UID in sopclass.ini: %s=%r", name, uid)
    return sops


def load_transfer_syntaxes():
    """Load every Transfer Syntax UID listed under [TRANSFER_SYNTAXES] in
    sopclass.ini. Falls back to Implicit/Explicit VR LE if the section is
    missing or empty."""
    create_sop_ini()
    config = configparser.ConfigParser()
    config.read(SOP_INI)
    ts = []
    if "TRANSFER_SYNTAXES" in config:
        for name, uid in config["TRANSFER_SYNTAXES"].items():
            try:
                ts.append(UID(uid.strip()))
            except Exception:
                app_logger.warning("Skipping invalid Transfer Syntax UID in sopclass.ini: %s=%r", name, uid)
    if not ts:
        ts = [UID("1.2.840.10008.1.2"), UID("1.2.840.10008.1.2.1")]
    return ts


def save_sop_ini(sop_classes: dict, transfer_syntaxes: dict):
    """Persist an edited SOP Class / Transfer Syntax set back to sopclass.ini.
    Used by the in-GUI SOP editor."""
    config = configparser.ConfigParser()
    config["SOP_CLASSES"] = sop_classes
    config["TRANSFER_SYNTAXES"] = transfer_syntaxes
    with atomic_open_for_write(SOP_INI, mode="w") as f:
        config.write(f)
    write_audit_log("SOP-CONFIG-CHANGED", f"sop_classes={len(sop_classes)} transfer_syntaxes={len(transfer_syntaxes)}")


def read_sop_ini_raw():
    """Return (sop_classes_dict, transfer_syntaxes_dict) for editing in the GUI."""
    create_sop_ini()
    config = configparser.ConfigParser()
    config.read(SOP_INI)
    sop_classes = dict(config["SOP_CLASSES"]) if "SOP_CLASSES" in config else {}
    transfer_syntaxes = dict(config["TRANSFER_SYNTAXES"]) if "TRANSFER_SYNTAXES" in config else {}
    return sop_classes, transfer_syntaxes

# =========================================================
# ENCRYPTION
# =========================================================

def generate_key():
    if not os.path.exists(KEY_FILE):
        atomic_write(KEY_FILE, Fernet.generate_key(), binary=True)


def load_key():
    with open(KEY_FILE, "rb") as f:
        return f.read()


def encrypt_and_save(file_name, data):
    fernet = Fernet(load_key())
    atomic_write(file_name, fernet.encrypt(data.encode()), binary=True)


def decrypt_and_load(file_name):
    if not os.path.exists(file_name):
        return None
    fernet = Fernet(load_key())
    with open(file_name, "rb") as f:
        try:
            return fernet.decrypt(f.read()).decode()
        except Exception:
            log_exception(f"Failed to decrypt {file_name} (key mismatch or corrupt file)")
            return None

# =========================================================
# ADMIN PIN AUTHENTICATION
# =========================================================
# The Admin PIN is a secret someone *types in* to prove identity, not
# config-at-rest -- so it is hashed with a random salt (PBKDF2-HMAC-SHA256)
# rather than encrypted with the shared Fernet key used elsewhere in this
# file. Only the salt + hash are ever persisted; the PIN itself is never
# stored or logged.


def _hash_pin(pin: str, salt: bytes) -> str:
    dk = hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), salt, ADMIN_PIN_PBKDF2_ITERATIONS)
    return dk.hex()


def admin_pin_is_configured():
    return os.path.exists(ADMIN_AUTH_FILE)


def save_admin_pin(pin: str):
    salt = secrets.token_bytes(16)
    record = {
        "salt": salt.hex(),
        "hash": _hash_pin(pin, salt),
        "iterations": ADMIN_PIN_PBKDF2_ITERATIONS,
        "updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    atomic_write(ADMIN_AUTH_FILE, json.dumps(record, indent=2))


def verify_admin_pin(pin: str) -> bool:
    if not admin_pin_is_configured():
        return False
    try:
        with open(ADMIN_AUTH_FILE, "r", encoding="utf-8") as f:
            record = json.load(f)
        salt = bytes.fromhex(record["salt"])
        expected = record["hash"]
    except Exception:
        log_exception("Failed to read admin_auth.json")
        return False
    return secrets.compare_digest(_hash_pin(pin, salt), expected)


def validate_pin_format(pin: str) -> (bool, str):
    """Returns (ok, error_message)."""
    if not pin.isdigit():
        return False, "PIN must contain only digits."
    if len(pin) < ADMIN_PIN_MIN_DIGITS:
        return False, f"PIN must be at least {ADMIN_PIN_MIN_DIGITS} digits."
    return True, ""

# =========================================================
# TLS SCAFFOLD (off by default; auto-detected for known TLS ports)
# =========================================================
# This section wires up TLS support for DICOM associations. tls.enabled
# is the manual override -- when True, every outbound association uses
# TLS regardless of port. It defaults to False because forcing TLS onto
# every destination would break plaintext-only servers (most test/dev
# SCPs, and any node on a standard non-TLS port). Instead,
# build_ssl_context_for_client() auto-enables TLS per-connection whenever
# the destination port is DICOM_TLS_PORT (2762, the IANA-registered
# 'dicom-tls' port) -- so destinations that live on the standard TLS port
# get TLS automatically, with no config file edits needed, while every
# other destination keeps working in plaintext exactly as before. When a
# ca_cert isn't supplied, certificate verification is skipped
# (CERT_NONE / check_hostname=False) rather than failing outright, and no
# client certificate is presented unless require_mutual_tls is set with
# cert/key supplied -- so this works against most DICOM-TLS servers with
# zero certs on hand. This file is auto-created with these defaults the
# first time the app runs in a given working folder.

DEFAULT_TLS_CONFIG = {
    "enabled": False,
    "ca_cert": "",
    "cert": "",
    "key": "",
    "require_mutual_tls": False,
}


def load_tls_config():
    if not os.path.exists(TLS_CONFIG_FILE):
        save_tls_config(DEFAULT_TLS_CONFIG)
        return dict(DEFAULT_TLS_CONFIG)
    try:
        with open(TLS_CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        merged = dict(DEFAULT_TLS_CONFIG)
        merged.update(cfg)
        return merged
    except Exception:
        log_exception("Failed to load tls_config.json, using defaults")
        return dict(DEFAULT_TLS_CONFIG)


def save_tls_config(cfg):
    atomic_write(TLS_CONFIG_FILE, json.dumps(cfg, indent=2))


def build_ssl_context_for_server(tls_cfg):
    """Build an ssl.SSLContext for the receiver if TLS is enabled and certs
    are present. Returns None if TLS is disabled (plaintext server, default
    behaviour) or if certs are missing/invalid (falls back to plaintext with
    a logged warning rather than crashing the receiver)."""
    if not tls_cfg.get("enabled"):
        return None
    cert, key = tls_cfg.get("cert"), tls_cfg.get("key")
    if not cert or not key or not os.path.exists(cert) or not os.path.exists(key):
        app_logger.warning("TLS enabled but cert/key missing or invalid; starting in plaintext mode.")
        return None
    try:
        import ssl
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(certfile=cert, keyfile=key)
        if tls_cfg.get("require_mutual_tls") and tls_cfg.get("ca_cert"):
            ctx.load_verify_locations(cafile=tls_cfg["ca_cert"])
            ctx.verify_mode = ssl.CERT_REQUIRED
        return ctx
    except Exception:
        log_exception("Failed to build server TLS context; starting in plaintext mode.")
        return None


def build_ssl_context_for_client(tls_cfg, port=None, dest_label=None):
    """Build an ssl.SSLContext for outbound (pusher) associations. Returns
    None if TLS isn't needed (plaintext, default behaviour for everything
    except known TLS ports) or if context creation fails.
    TLS is used when EITHER tls_cfg['enabled'] is True (manual override,
    applies to every destination) OR `port` equals DICOM_TLS_PORT (2762 --
    auto-detected per connection so destinations living on the standard
    DICOM-TLS port work without touching config, while everything else on
    a normal port is untouched and stays plaintext).

    §3.5 fix: when no ca_cert is configured, this still completes the
    handshake with certificate verification OFF (CERT_NONE /
    check_hostname=False) -- kept as the default so zero-config TLS keeps
    working out of the box against servers nobody has a CA cert for --
    but that used to be indistinguishable from a properly verified
    channel anywhere in the UI or logs. Now it's logged as an explicit
    warning every time, and the verification state is recorded (keyed by
    `dest_label`, falling back to the port number) so callers like the
    Destination Health panel can show it. See get_tls_verification_state().
    """
    use_tls = bool(tls_cfg.get("enabled")) or (port is not None and int(port) == DICOM_TLS_PORT)
    if not use_tls:
        return None
    label = dest_label or (f"port {port}" if port is not None else "unknown destination")
    try:
        import ssl
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        if tls_cfg.get("ca_cert") and os.path.exists(tls_cfg["ca_cert"]):
            ctx.load_verify_locations(cafile=tls_cfg["ca_cert"])
            _record_tls_verification_state(label, True)
        else:
            ctx.check_hostname = False
            ctx.verify_mode = __import__("ssl").CERT_NONE
            app_logger.warning(
                "TLS association with %s will complete WITHOUT certificate verification "
                "(no ca_cert configured) -- this protects against passive eavesdropping "
                "but NOT against a man-in-the-middle. Configure tls_config.ca_cert for a "
                "verified channel.", label)
            _record_tls_verification_state(label, False)
        if tls_cfg.get("require_mutual_tls") and tls_cfg.get("cert") and tls_cfg.get("key"):
            ctx.load_cert_chain(certfile=tls_cfg["cert"], keyfile=tls_cfg["key"])
        return ctx
    except Exception:
        log_exception("Failed to build client TLS context; starting in plaintext mode.")
        return None


# Per-destination TLS verification state (§3.5), written by
# build_ssl_context_for_client() and read by anything that wants to show
# "TLS enabled but unverified" rather than a plain "TLS enabled" that
# looks identical to a properly verified channel.
_tls_verification_state = {}
_tls_verification_state_lock = threading.Lock()


def _record_tls_verification_state(label, verified):
    with _tls_verification_state_lock:
        _tls_verification_state[label] = {"verified": verified, "checked_at": time.time()}


def get_tls_verification_state(label):
    """Returns True (TLS verified via a configured ca_cert), False (TLS
    in use but certificate verification is disabled -- §3.5), or None
    (TLS not in use, or this label has never been checked)."""
    with _tls_verification_state_lock:
        entry = _tls_verification_state.get(label)
    return entry["verified"] if entry else None

# =========================================================
# MULTI-DESTINATION PUSH CONFIG (with legacy auto-migration)
# =========================================================
# destinations.enc stores an encrypted JSON list of destination profiles:
#   [{"name": "...", "ae": "...", "ip": "...", "port": "...", "default": bool}, ...]
# The old single-destination push.enc ("AE|IP|PORT") is auto-migrated into
# this format the first time the app runs, as a destination named "Default".

def _migrate_legacy_push_config():
    """If destinations.enc doesn't exist yet but the legacy push.enc does,
    convert it into the new multi-destination format automatically."""
    if os.path.exists(DESTINATIONS_CONFIG):
        return
    legacy = decrypt_and_load(PUSH_CONFIG)
    if not legacy:
        return
    parts = legacy.split("|")
    if len(parts) != 3:
        return
    rae, rip, rport = parts
    destinations = [{
        "name": "Default",
        "ae": rae,
        "calling_ae": "",   # blank = use "RAPPS_PUSH" fallback
        "ip": rip,
        "port": rport,
        "default": True,
    }]
    encrypt_and_save(DESTINATIONS_CONFIG, json.dumps(destinations))
    app_logger.info("Migrated legacy push.enc into destinations.enc as 'Default'.")


def load_destinations():
    """Return the list of destination profile dicts. Triggers legacy
    migration on first call if needed."""
    _migrate_legacy_push_config()
    data = decrypt_and_load(DESTINATIONS_CONFIG)
    if not data:
        return []
    try:
        return json.loads(data)
    except Exception:
        log_exception("Failed to parse destinations.enc")
        return []


def save_destinations(destinations):
    encrypt_and_save(DESTINATIONS_CONFIG, json.dumps(destinations))
    write_audit_log("DESTINATIONS-CHANGED", f"count={len(destinations)}")


def get_default_destination():
    dests = load_destinations()
    for d in dests:
        if d.get("default"):
            return d
    return dests[0] if dests else None


def get_destination_by_name(name):
    for d in load_destinations():
        if d["name"] == name:
            return d
    return None

# =========================================================
# NOTIFICATIONS (desktop / email / webhook)
# =========================================================
# One central dispatcher (notify_event) is called from every place in the
# app that already knows something notification-worthy happened (receiver
# start/stop, push complete/failed, queue growing, destination on/offline,
# disk space low, backup events, TLS errors, auth failures, config
# changes). Each event type can be toggled independently, and each of the
# three channels (desktop/email/webhook) can be toggled independently.
# Sending is always backgrounded -- notify_event() itself never blocks the
# calling thread on SMTP or HTTP I/O.

NOTIFICATION_EVENT_TYPES = [
    ("receiver_started", "Receiver Started"),
    ("receiver_stopped", "Receiver Stopped"),
    ("push_complete", "Push Complete"),
    ("push_failed", "Push Failed"),
    ("queue_growing", "Queue Growing"),
    ("destination_offline", "Destination Offline"),
    ("destination_online", "Destination Online"),
    ("disk_space_low", "Disk Space Low"),
    ("backup_completed", "Backup Completed"),
    ("backup_failed", "Backup Failed"),
    ("tls_errors", "TLS Errors"),
    ("authentication_failures", "Authentication Failures"),
    ("configuration_changes", "Configuration Changes"),
    ("doc_transfer_failure_rate", "Document Transfer Failure Rate"),
]

CONFIG_CHANGE_AUDIT_EVENTS = {
    "DESTINATIONS-CHANGED", "ROUTING-RULES-CHANGED", "SOP-CONFIG-CHANGED",
    "LOG-RETENTION-CHANGED", "NOTIFICATIONS-CONFIG-CHANGED",
    "ADMIN-PIN-CHANGED", "ADMIN-PIN-SET", "TLS-CONFIG-CHANGED",
    "LDAP-CONFIG-CHANGED", "BANDWIDTH-CONFIG-CHANGED", "BACKUP-SCHEDULE-CHANGED",
}

DEFAULT_NOTIFICATIONS_CONFIG = {
    "channels": {"desktop": True, "email": False, "webhook": False},
    "events": {key: (key != "configuration_changes") for key, _ in NOTIFICATION_EVENT_TYPES},
    "email": {
        "smtp_host": "", "smtp_port": 587, "smtp_user": "", "smtp_password": "",
        "from_addr": "", "to_addrs": "", "use_tls": True,
    },
    "webhook": {"url": ""},
    # 3.2 -- one-shot (re-arms once the queue drops back below threshold)
    # alert when the offline queue grows past this size.
    "thresholds": {"offline_queue_size": 20},
}

_notifications_config_cache = {"value": None}


def load_notifications_config():
    """Encrypted at rest (destinations.enc-style) because it can hold an
    SMTP password. Cached in memory and only re-read after
    save_notifications_config() to avoid decrypting on every event."""
    if _notifications_config_cache["value"] is not None:
        return _notifications_config_cache["value"]
    data = decrypt_and_load(NOTIFICATIONS_CONFIG)
    cfg = json.loads(json.dumps(DEFAULT_NOTIFICATIONS_CONFIG))  # deep copy
    if data:
        try:
            loaded = json.loads(data)
            for section in ("channels", "events", "email", "webhook", "thresholds"):
                if section in loaded and isinstance(loaded[section], dict):
                    cfg[section].update(loaded[section])
        except Exception:
            log_exception("Failed to parse notifications.enc, using defaults")
    _notifications_config_cache["value"] = cfg
    return cfg


def save_notifications_config(cfg):
    encrypt_and_save(NOTIFICATIONS_CONFIG, json.dumps(cfg))
    _notifications_config_cache["value"] = cfg
    write_audit_log("NOTIFICATIONS-CONFIG-CHANGED", "Notification settings updated")


def _send_notification_email(cfg, subject, body):
    email_cfg = cfg["email"]
    if not email_cfg.get("smtp_host") or not email_cfg.get("to_addrs"):
        app_logger.warning("Email notification skipped: SMTP host or recipients not configured")
        return
    try:
        import smtplib
        from email.mime.text import MIMEText

        msg = MIMEText(body)
        msg["Subject"] = f"[R-Apps DICOM] {subject}"
        msg["From"] = email_cfg.get("from_addr") or email_cfg.get("smtp_user") or "noreply@rapps.local"
        to_list = [a.strip() for a in email_cfg.get("to_addrs", "").split(",") if a.strip()]
        msg["To"] = ", ".join(to_list)

        with smtplib.SMTP(email_cfg["smtp_host"], int(email_cfg.get("smtp_port", 587)), timeout=10) as server:
            if email_cfg.get("use_tls", True):
                server.starttls()
            if email_cfg.get("smtp_user"):
                server.login(email_cfg["smtp_user"], email_cfg.get("smtp_password", ""))
            server.sendmail(msg["From"], to_list, msg.as_string())
        app_logger.info("Notification email sent: %s", subject)
    except Exception:
        log_exception(f"Failed to send notification email: {subject}")


def _send_notification_webhook(cfg, event_key, title, message):
    url = cfg["webhook"].get("url")
    if not url:
        app_logger.warning("Webhook notification skipped: no URL configured")
        return
    try:
        import urllib.request

        payload = json.dumps({
            "event": event_key, "title": title, "message": message,
            "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
            "host": _computer_name(),
        }).encode("utf-8")
        req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"}, method="POST")
        urllib.request.urlopen(req, timeout=10)
        app_logger.info("Notification webhook delivered: %s", event_key)
    except Exception:
        log_exception(f"Failed to deliver notification webhook: {event_key}")


NOTIFICATION_HISTORY_MAX = 200
notification_history = []  # newest last; list of {"time","title","message","kind","event_key"}


def _infer_notification_kind(event_key, title):
    text = f"{event_key} {title}".lower()
    if any(w in text for w in ("fail", "error", "offline", "denied", "critical")):
        return "error"
    if any(w in text for w in ("warn", "retry", "degraded", "low disk", "stale")):
        return "warning"
    if any(w in text for w in ("success", "complete", "generated", "online", "connected")):
        return "success"
    return "info"


def _refresh_notification_center_badge_threadsafe():
    """Queues a UI refresh for the notification bell badge -- same
    ui_event_queue + pump_events pattern show_toast_threadsafe uses, so
    this is safe to call from notify_event() on any thread."""
    ui_event_queue.put(("notification_added", None))


def notify_event(event_key, title, message):
    """Central notification entry point. Safe to call from any thread --
    all actual I/O (SMTP/HTTP) is backgrounded; desktop toasts already
    route through the thread-safe UI event queue."""
    try:
        cfg = load_notifications_config()
        if not cfg["events"].get(event_key, True):
            return  # this specific event type is disabled

        notification_history.append({
            "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "title": title, "message": message, "event_key": event_key,
            "kind": _infer_notification_kind(event_key, title),
        })
        del notification_history[:-NOTIFICATION_HISTORY_MAX]
        _refresh_notification_center_badge_threadsafe()

        if cfg["channels"].get("desktop", True):
            show_toast_threadsafe(title, message)

        if cfg["channels"].get("email", False):
            threading.Thread(target=_send_notification_email, args=(cfg, title, message), daemon=True).start()

        if cfg["channels"].get("webhook", False):
            threading.Thread(target=_send_notification_webhook, args=(cfg, event_key, title, message), daemon=True).start()
    except Exception:
        log_exception(f"notify_event failed for {event_key}")



# =========================================================
# ROUTING RULES ENGINE
# =========================================================
# Rules are evaluated top-down; the first matching rule's destination wins.
# A rule matches on modality / institution / source AE (any blank field is
# a wildcard). If no rule matches, the default destination is used.
# Format (routing_rules.json, list, evaluated in order):
#   [{"modality": "CT", "institution": "", "source_ae": "", "destination": "PACS_A"}, ...]

def load_routing_rules():
    if not os.path.exists(ROUTING_RULES_FILE):
        return []
    try:
        with open(ROUTING_RULES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        log_exception("Failed to load routing_rules.json")
        return []


def save_routing_rules(rules):
    atomic_write(ROUTING_RULES_FILE, json.dumps(rules, indent=2))
    write_audit_log("ROUTING-RULES-CHANGED", f"count={len(rules)}")


def _rule_matches_patient_fields(rule, modality, institution, source_ae):
    """The exact matching logic resolve_destination_for_patient() applies
    per-rule, factored out so the Routing Rules tab's dry-run (5.2) tests
    a draft rule against live patient_data without duplicating this."""
    if rule.get("modality") and rule["modality"].upper() != modality.upper():
        return False
    if rule.get("institution") and rule["institution"].lower() != institution.lower():
        return False
    if rule.get("source_ae") and rule["source_ae"].upper() != source_ae.upper():
        return False
    return True


def resolve_destination_for_patient(pid):
    """Apply routing rules to a patient's data and return the matching
    destination profile dict, or the default destination if no rule
    matches. Returns None if there is no destination configured at all."""
    with data_lock:
        d = patient_data.get(pid, {})
        modality = d.get("modality", "")
        institution = d.get("institution", "")
        source_ae = d.get("source", "")

    for rule in load_routing_rules():
        if not _rule_matches_patient_fields(rule, modality, institution, source_ae):
            continue
        dest = get_destination_by_name(rule.get("destination", ""))
        if dest:
            return dest

    return get_default_destination()

# =========================================================
# RESUME TRANSFERS (checkpointing) + OFFLINE QUEUE
# =========================================================
# Two related but distinct pieces of enterprise reliability:
#
#   * Resume Transfers: if a push dies partway through (1000 images,
#     fails at 843), the NEXT attempt -- whether that's the existing
#     in-process retry-with-backoff below, a later offline-queue retry,
#     or a manual re-push after an app restart -- picks up at 844
#     instead of re-sending everything. State lives in
#     TRANSFER_CHECKPOINTS_FILE, keyed by patient ID, and survives app
#     restarts (plain JSON, not secret).
#
#   * Offline Queue: once a destination has been retried the in-process
#     MAX_RETRY_ATTEMPTS times and is still unreachable, the patient is
#     handed off to a persistent queue (OFFLINE_QUEUE_FILE) instead of
#     being abandoned as a plain, forgotten "Failed". A background
#     worker thread (offline_queue_worker_loop) keeps retrying queued
#     items on an exponential backoff (capped) and automatically resumes
#     them -- using the same checkpoint state -- the moment the
#     destination comes back online.

_checkpoint_lock = threading.RLock()
_checkpoint_cache = {"value": None}


def load_transfer_checkpoints():
    if _checkpoint_cache["value"] is not None:
        return _checkpoint_cache["value"]
    data = {}
    try:
        if os.path.exists(TRANSFER_CHECKPOINTS_FILE):
            with open(TRANSFER_CHECKPOINTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
    except Exception:
        log_exception("Failed to load transfer_checkpoints.json")
        data = {}
    _checkpoint_cache["value"] = data
    return data


def _save_transfer_checkpoints():
    try:
        with _checkpoint_lock:
            atomic_write(TRANSFER_CHECKPOINTS_FILE, json.dumps(_checkpoint_cache["value"] or {}, indent=2))
    except Exception:
        log_exception("Failed to save transfer_checkpoints.json")


def get_transfer_checkpoint(pid, destination_name):
    """Returns the set of SOP Instance UIDs already confirmed sent for
    this patient's CURRENT transfer attempt to this destination, or an
    empty set if there's no matching in-progress checkpoint (e.g. first
    attempt, or the destination changed since the last attempt)."""
    with _checkpoint_lock:
        checkpoints = load_transfer_checkpoints()
        entry = checkpoints.get(pid)
        if not entry or entry.get("destination") != destination_name:
            return set()
        return set(entry.get("sent_sop_uids", []))


def record_checkpoint_progress(pid, destination_name, sop_uid):
    """Called immediately after each individual file is confirmed sent,
    so a crash/interruption mid-transfer loses at most the one in-flight
    file, never the whole batch."""
    with _checkpoint_lock:
        checkpoints = load_transfer_checkpoints()
        entry = checkpoints.setdefault(pid, {
            "destination": destination_name, "sent_sop_uids": [],
            "retry_history": [], "interrupted_at": None, "failure_reason": None,
        })
        if entry.get("destination") != destination_name:
            # Resuming toward a different destination than last time --
            # the old progress doesn't apply, start this checkpoint over.
            entry = {"destination": destination_name, "sent_sop_uids": [],
                      "retry_history": [], "interrupted_at": None, "failure_reason": None}
            checkpoints[pid] = entry
        if sop_uid not in entry["sent_sop_uids"]:
            entry["sent_sop_uids"].append(sop_uid)
        _save_transfer_checkpoints()


def record_checkpoint_interruption(pid, destination_name, reason):
    """Records WHY and WHEN a transfer stopped short, and appends to the
    retry history, without discarding the sent_sop_uids progress already
    recorded -- that's what makes the next attempt a resume, not a
    restart."""
    with _checkpoint_lock:
        checkpoints = load_transfer_checkpoints()
        entry = checkpoints.setdefault(pid, {
            "destination": destination_name, "sent_sop_uids": [],
            "retry_history": [], "interrupted_at": None, "failure_reason": None,
        })
        now_iso = datetime.datetime.now().isoformat(timespec="seconds")
        entry["destination"] = destination_name
        entry["interrupted_at"] = now_iso
        entry["failure_reason"] = reason
        entry.setdefault("retry_history", []).append({
            "time": now_iso, "attempt": len(entry.get("retry_history", [])) + 1, "reason": reason,
        })
        _save_transfer_checkpoints()


def clear_transfer_checkpoint(pid):
    """Called once a patient's transfer completes successfully -- there's
    nothing left to resume."""
    with _checkpoint_lock:
        checkpoints = load_transfer_checkpoints()
        if pid in checkpoints:
            del checkpoints[pid]
            _save_transfer_checkpoints()


def get_checkpoint_info(pid):
    """Read-only accessor for the UI (resume progress, retry history,
    failure reason, interruption time)."""
    with _checkpoint_lock:
        return load_transfer_checkpoints().get(pid)


# ---------------------------------------------------------
# Offline queue
# ---------------------------------------------------------

RETRY_QUEUE_BASE_INTERVAL_SEC = 300  # retry failed pushes every 5 minutes
RETRY_QUEUE_MAX_INTERVAL_SEC = 300   # fixed cadence -- no exponential backoff growth
OFFLINE_QUEUE_WORKER_TICK_SEC = 10   # how often the worker wakes up to check due items

_offline_queue_lock = threading.RLock()
# Rolling samples feeding the Performance Metrics page's "Average Queue
# Time" and "Average Retry Time" stats -- in-memory only (same tradeoff
# as the Dashboard's history buffers), reset on app restart.
queue_wait_time_samples = deque(maxlen=200)
retry_interval_samples = deque(maxlen=200)
_offline_queue_cache = {"value": None}


def load_offline_queue():
    if _offline_queue_cache["value"] is not None:
        return _offline_queue_cache["value"]
    items = []
    try:
        if os.path.exists(OFFLINE_QUEUE_FILE):
            with open(OFFLINE_QUEUE_FILE, "r", encoding="utf-8") as f:
                items = json.load(f)
    except Exception:
        log_exception("Failed to load offline_queue.json")
        items = []
    _offline_queue_cache["value"] = items
    return items


def _save_offline_queue():
    try:
        with _offline_queue_lock:
            atomic_write(OFFLINE_QUEUE_FILE, json.dumps(_offline_queue_cache["value"] or [], indent=2))
    except Exception:
        log_exception("Failed to save offline_queue.json")


def enqueue_offline(pid, destination_name, reason, kind="dicom"):
    """Adds (or updates) a patient in the persistent offline queue.
    Preserves original queue position/queued_at if the patient was
    already queued -- re-failing doesn't push it to the back of the line.

    D.4: `kind` distinguishes a full DICOM-send failure ("dicom", the
    original/default behavior -- also what any pre-existing queued item
    without this field is treated as) from a document-transfer-only
    failure ("documents"), so offline_queue_worker_loop knows whether to
    retry via push_single_patient or via push_patient_documents alone."""
    with _offline_queue_lock:
        items = load_offline_queue()
        for item in items:
            if item["pid"] == pid and item.get("kind", "dicom") == kind:
                item["last_error"] = reason
                item["attempt_count"] = item.get("attempt_count", 0) + 1
                item.setdefault("retry_history", []).append({
                    "time": datetime.datetime.now().isoformat(timespec="seconds"),
                    "attempt": item["attempt_count"], "reason": reason,
                })
                _save_offline_queue()
                return
        now = time.time()
        items.append({
            "pid": pid,
            "destination_name": destination_name,
            "kind": kind,
            "queued_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "queued_at_epoch": now,
            "last_error": reason,
            "attempt_count": 1,
            "next_retry_at": now + RETRY_QUEUE_BASE_INTERVAL_SEC,
            "current_interval_sec": RETRY_QUEUE_BASE_INTERVAL_SEC,
            "retry_history": [{
                "time": datetime.datetime.now().isoformat(timespec="seconds"),
                "attempt": 1, "reason": reason,
            }],
        })
        _save_offline_queue()
        write_audit_log("QUEUE-CHANGED", f"pid={pid} action=enqueued dest={destination_name} kind={kind}")


def dequeue_offline(pid, kind=None):
    """Removes the queued entry for pid. If `kind` is given, only removes
    the entry matching that kind (so a "documents" retry succeeding
    doesn't also drop an unrelated "dicom" entry for the same patient, and
    vice versa). kind=None (the default, used by every pre-existing call
    site) preserves the original behavior of removing by pid alone."""
    with _offline_queue_lock:
        items = load_offline_queue()
        if kind is None:
            removed = next((i for i in items if i["pid"] == pid), None)
            remaining = [i for i in items if i["pid"] != pid]
        else:
            removed = next((i for i in items if i["pid"] == pid and i.get("kind", "dicom") == kind), None)
            remaining = [i for i in items if not (i["pid"] == pid and i.get("kind", "dicom") == kind)]
        if len(remaining) != len(items):
            _offline_queue_cache["value"] = remaining
            _save_offline_queue()
            write_audit_log("QUEUE-CHANGED", f"pid={pid} action=dequeued")
            if removed and removed.get("queued_at_epoch"):
                queue_wait_time_samples.append(time.time() - removed["queued_at_epoch"])


def get_offline_queue_items():
    """Read-only snapshot for the UI, oldest first (FIFO order is the
    literal list order, since we always append new items to the end and
    never reorder)."""
    with _offline_queue_lock:
        return list(load_offline_queue())


def get_offline_queue_summary():
    """Returns (queue_size, oldest_item_or_None, min_next_retry_epoch_or_None,
    min_current_interval_or_None) for the Offline Queue status panel."""
    items = get_offline_queue_items()
    if not items:
        return 0, None, None, None
    oldest = min(items, key=lambda i: i.get("queued_at_epoch", 0))
    soonest_retry = min((i.get("next_retry_at", 0) for i in items), default=None)
    soonest_item = min(items, key=lambda i: i.get("next_retry_at", float("inf")))
    return len(items), oldest, soonest_retry, soonest_item.get("current_interval_sec")


_offline_queue_worker_started = {"value": False}
_offline_queue_processing = {"value": False}


def _offline_queue_destination_believed_online(destination_name):
    """Uses the same destination_health_cache the PACS Health Monitor and
    Dashboard populate. If we've never checked this destination, assume
    it MIGHT be reachable so the queue still gets a first attempt rather
    than waiting indefinitely on a health check that hasn't run yet."""
    record = destination_health_cache.get(destination_name)
    if record is None:
        return True
    return bool(record.get("online"))


def offline_queue_worker_loop():
    """Background daemon: wakes up every OFFLINE_QUEUE_WORKER_TICK_SEC,
    retries whichever queued items are due AND whose destination looks
    reachable, on a fixed 5-minute cadence (RETRY_QUEUE_BASE_INTERVAL_SEC).
    Processes one item at a time to avoid hammering a struggling
    destination with parallel connection attempts."""
    while not app_shutdown_event.is_set():
        try:
            if not _offline_queue_processing["value"]:
                _offline_queue_processing["value"] = True
                try:
                    now = time.time()
                    for item in get_offline_queue_items():
                        if app_shutdown_event.is_set():
                            break
                        if item.get("next_retry_at", 0) > now:
                            continue
                        if not _offline_queue_destination_believed_online(item["destination_name"]):
                            continue
                        dest = get_destination_by_name(item["destination_name"])
                        if not dest:
                            continue  # destination profile was deleted; leave queued in case it's re-added
                        item_kind = item.get("kind", "dicom")
                        try:
                            if item_kind == "documents":
                                # D.4: retry only the document leg, never the
                                # full DICOM send loop.
                                push_patient_documents(item["pid"], dest)
                                with data_lock:
                                    ok = not bool(patient_data.get(item["pid"], {}).get("last_doc_transfer_error"))
                            else:
                                sent, total, ok = push_single_patient(item["pid"], destination=dest)
                        except Exception:
                            log_exception(f"Offline queue retry failed for {item['pid']}")
                            ok = False
                        if ok:
                            dequeue_offline(item["pid"], kind=item_kind)
                            notify_event("push_complete", "Offline Queue: Push Succeeded",
                                        f"{item['pid']} delivered to {item['destination_name']} after being queued.")
                        else:
                            with _offline_queue_lock:
                                items = load_offline_queue()
                                for it in items:
                                    if it["pid"] == item["pid"] and it.get("kind", "dicom") == item_kind:
                                        new_interval = min(
                                            it.get("current_interval_sec", RETRY_QUEUE_BASE_INTERVAL_SEC) * 2,
                                            RETRY_QUEUE_MAX_INTERVAL_SEC,
                                        )
                                        it["current_interval_sec"] = new_interval
                                        it["next_retry_at"] = time.time() + new_interval
                                        retry_interval_samples.append(new_interval)
                                        break
                                _save_offline_queue()
                finally:
                    _offline_queue_processing["value"] = False
        except Exception:
            log_exception("offline_queue_worker_loop iteration failed")
        app_shutdown_event.wait(OFFLINE_QUEUE_WORKER_TICK_SEC)


def start_offline_queue_worker_thread():
    if _offline_queue_worker_started["value"]:
        return
    _offline_queue_worker_started["value"] = True
    threading.Thread(target=offline_queue_worker_loop, daemon=True).start()


# =========================================================
# BANDWIDTH LIMITER
# =========================================================
# A simple token-bucket throttle shared by DICOM Push, Import, and Export
# (Export calls in once the ZIP Export feature is built). Config is a
# plain JSON preset (not secret) so it's easy for an admin to inspect/
# edit by hand if needed, same as routing_rules.json.

BANDWIDTH_PRESETS_MBPS = {
    "Unlimited": 0, "1 Mbps": 1, "5 Mbps": 5, "10 Mbps": 10,
    "25 Mbps": 25, "50 Mbps": 50, "100 Mbps": 100, "Custom": None,
}
DEFAULT_BANDWIDTH_CONFIG = {"preset": "Unlimited", "custom_mbps": 10}

_bandwidth_config_cache = {"value": None}


def load_bandwidth_config():
    if _bandwidth_config_cache["value"] is not None:
        return _bandwidth_config_cache["value"]
    cfg = dict(DEFAULT_BANDWIDTH_CONFIG)
    try:
        if os.path.exists(BANDWIDTH_CONFIG_FILE):
            with open(BANDWIDTH_CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            cfg.update(loaded)
    except Exception:
        log_exception("Failed to load bandwidth_config.json, using defaults")
    _bandwidth_config_cache["value"] = cfg
    return cfg


def get_effective_bandwidth_mbps(cfg=None):
    """Returns the effective Mbps limit (0 = unlimited) for the current
    config: a named preset, or the custom value when preset == 'Custom'."""
    cfg = cfg or load_bandwidth_config()
    preset = cfg.get("preset", "Unlimited")
    if preset == "Custom":
        try:
            return max(0.0, float(cfg.get("custom_mbps", 0)))
        except (TypeError, ValueError):
            return 0.0
    return BANDWIDTH_PRESETS_MBPS.get(preset, 0) or 0


def save_bandwidth_config(cfg):
    try:
        atomic_write(BANDWIDTH_CONFIG_FILE, json.dumps(cfg, indent=2))
        _bandwidth_config_cache["value"] = cfg
        bandwidth_limiter.reload()
        write_audit_log("BANDWIDTH-CONFIG-CHANGED",
                        f"preset={cfg.get('preset')} custom_mbps={cfg.get('custom_mbps')} "
                        f"effective_mbps={get_effective_bandwidth_mbps(cfg)}")
    except Exception:
        log_exception("Failed to save bandwidth_config.json")
        raise



# =========================================================
# APPLICATION SETTINGS (Priority 10 -- Settings & Performance)
# =========================================================
# Single JSON-backed store for every preference in the Settings tab.
# Not secret, so plain JSON (matches BANDWIDTH_CONFIG_FILE / VIEW_OPTIONS_FILE
# -- the app's existing convention for non-sensitive config). Every
# load_*/save_* elsewhere in the file follows this same shape; APP_SETTINGS
# below is the in-memory singleton every other part of the app reads
# through the get_*() helpers so there is exactly one source of truth
# and no parallel configuration system.

APP_SETTINGS_FILE = "app_settings.json"

# Refresh-rate choices offered in the UI, in milliseconds.
REFRESH_RATE_CHOICES_MS = [100, 250, 500, 1000, 2000, 5000]

# Performance Mode presets. Selecting a mode pushes these values into the
# corresponding settings (the user can still fine-tune afterward -- picking
# a mode just seeds sensible defaults, it doesn't lock the sliders).
PERFORMANCE_MODE_PRESETS = {
    "High Performance": {"refresh_rate_ms": 250, "max_worker_threads": 8, "max_simultaneous_pushes": 8},
    "Balanced":          {"refresh_rate_ms": 1000, "max_worker_threads": 4, "max_simultaneous_pushes": 4},
    "Power Saving":      {"refresh_rate_ms": 5000, "max_worker_threads": 2, "max_simultaneous_pushes": 2},
}

LOG_LEVEL_CHOICES = ["Errors", "Warnings", "Information", "Debug"]
_LOG_LEVEL_TO_LOGGING = {
    "Errors": logging.ERROR, "Warnings": logging.WARNING,
    "Information": logging.INFO, "Debug": logging.DEBUG,
}

DEFAULT_APP_SETTINGS = {
    # ---- Performance ----
    "refresh_rate_ms": REFRESH_INTERVAL_MS,
    "performance_mode": "Balanced",
    "disable_graphs": False,        # skip sparkline/chart rendering everywhere -- helps on slower machines
    # ---- Appearance ----
    "ui_scaling_pct": 100,          # 90/100/110/125/150
    "font_scale_pct": 100,          # multiplies FONT_SCALE sizes
    "compact_mode": False,
    "remember_window_geometry": True,
    "launch_maximized": False,
    # ---- Notifications ----
    "toast_notifications_enabled": True,
    "notification_sounds_enabled": True,
    "notification_duration_sec": 4,
    "critical_only_notifications": False,
    # ---- Startup & Behavior ----
    "launch_on_system_startup": False,
    "auto_start_receiver_on_launch": False,
    "auto_restart_receiver": False,
    "restore_previous_session": False,
    "minimize_to_tray_on_close": False,
    "confirm_before_close": True,
    "auto_check_for_updates": False,
    # ---- Logging & Diagnostics ----
    "logging_level": "Information",
    # ---- Accessibility ----
    "high_contrast_mode": False,
    "larger_click_targets": False,
    "reduced_motion": False,
    # ---- Advanced ----
    "max_worker_threads": DEFAULT_PUSH_WORKER_THREADS,
    "max_simultaneous_pushes": DEFAULT_PUSH_WORKER_THREADS,
    "network_timeout_sec": PUSH_ASSOC_TIMEOUT_SEC,
    "cache_size_limit_mb": LOG_ROTATION_MAX_BYTES // (1024 * 1024),
    "experimental_features_enabled": False,
    # ---- Transcoding ----
    # When a push destination's association only accepts a LOSSY
    # compressed transfer syntax for a SOP Class (e.g. it rejects the
    # file's native lossless encoding and every uncompressed option),
    # _transcode_dataset_for_target() will only actually perform that
    # lossy re-encode if this is explicitly turned on. Off by default:
    # silently turning a lossless image into a lossy one is a clinical/
    # legal decision, not something a router should do on an admin's
    # behalf without them opting in. See Settings tab / SOP editor.
    "allow_lossy_transcode": False,
    # §fix: see the debug_logger() gating note near the pynetdicom imports.
    # Off by default -- full PDU/DIMSE tracing is a diagnostic tool, not a
    # permanent production log stream.
    "verbose_dicom_protocol_logging": False,
    # ---- Document Transfer (global receiver gate) ----
    # The receiver only opens the document-transfer TCP port when this is
    # on. Each destination still has its own per-destination
    # doc_transfer_enabled flag (see load_destinations()) controlling
    # whether the PUSHER side attempts a transfer -- this setting only
    # controls whether the RECEIVER listens at all.
    "doc_transfer_enabled": False,
    "doc_transfer_receiver_port": "11244",
    # Shared-secret the receiver requires from any doc-transfer sender.
    # Empty (the default, for backward compatibility with existing
    # deployments) means authentication is off -- explicitly, via this
    # setting, rather than the previous hardcoded no-op. Once set, every
    # connecting sender's auth_key must match this exactly or the
    # connection is rejected before any file is accepted.
    "doc_transfer_receiver_auth_key": "",
    # A.5: configurable chunk size / connect timeout for the doc-transfer
    # socket. Fall back to the DOC_TRANSFER_CHUNK_SIZE / _CONNECT_TIMEOUT_SEC
    # constants if these are ever missing from a loaded settings file.
    "doc_transfer_chunk_size_kb": 64,       # matches DOC_TRANSFER_CHUNK_SIZE (65536 bytes)
    "doc_transfer_timeout_sec": 10,         # matches DOC_TRANSFER_CONNECT_TIMEOUT_SEC
    # Hard ceiling on any single file the doc-transfer protocol will
    # accept, enforced on the RECEIVER side (rejected before any bytes
    # are written to disk) and checked pre-flight on the PUSHER side too
    # (so an oversized file never even opens a connection). Falls back to
    # DOC_TRANSFER_MAX_FILE_SIZE_BYTES if missing from a loaded settings
    # file, same pattern as chunk size / timeout above.
    "doc_transfer_max_size_mb": 500,
    # D.3: admin idle-timeout auto-lock. 0 = disabled.
    "admin_idle_timeout_min": 15,
}

_app_settings_cache = {"value": None}


def load_app_settings():
    """Loads app_settings.json, filling in any keys missing (new settings
    added in a later version, or a first run with no file yet) from
    DEFAULT_APP_SETTINGS -- same forward-compatible merge pattern used by
    load_notifications_config()/load_bandwidth_config() elsewhere."""
    if _app_settings_cache["value"] is not None:
        return dict(_app_settings_cache["value"])
    cfg = dict(DEFAULT_APP_SETTINGS)
    try:
        if os.path.exists(APP_SETTINGS_FILE):
            with open(APP_SETTINGS_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                cfg.update(loaded)
    except Exception:
        log_exception("Failed to load app_settings.json -- using defaults")
    _app_settings_cache["value"] = dict(cfg)
    return cfg


def save_app_settings(cfg):
    try:
        atomic_write(APP_SETTINGS_FILE, json.dumps(cfg, indent=2))
        _app_settings_cache["value"] = dict(cfg)
        try:
            write_audit_log("APP-SETTINGS-CHANGED", "settings updated")
        except Exception:
            pass
    except Exception:
        log_exception("Failed to save app_settings.json")
        raise


# In-memory singleton every other function in the app reads through the
# get_*() helpers below. Mutated + persisted together by
# _update_app_setting() so the file and this dict never drift apart.
APP_SETTINGS = load_app_settings()


def _update_app_setting(key, value):
    """Sets one setting, persists the whole settings dict, and returns it.
    The single choke point every Settings-tab control writes through."""
    APP_SETTINGS[key] = value
    try:
        save_app_settings(APP_SETTINGS)
    except Exception:
        # save_app_settings() already logs the full traceback -- this is
        # the single choke point every settings change in the app goes
        # through, so also surface it to the operator: otherwise the UI
        # looks like the change took, but it silently reverts on restart.
        try:
            ui_event_queue.put(("toast", ("Setting Not Saved",
                                           f"Failed to save '{key}' to disk -- see app.log for details.")))
        except Exception:
            pass


# ---- Larger Click Targets (Accessibility) --------------------------------
# Single source of truth for the two dimensions this setting controls:
# worklist row height and nav-button height. Every tree/nav builder in the
# app reads these two functions instead of hardcoding a row/button height,
# so toggling the setting (and calling refresh_ui_theme()) changes every
# worklist and every nav entry in the whole app at once.
_BASE_WORKLIST_ROW_HEIGHT = 28
_LARGE_WORKLIST_ROW_HEIGHT = 40
_BASE_NAV_BUTTON_HEIGHT = 36
_LARGE_NAV_BUTTON_HEIGHT = 48


def get_worklist_row_height():
    base = _LARGE_WORKLIST_ROW_HEIGHT if APP_SETTINGS.get("larger_click_targets") else _BASE_WORKLIST_ROW_HEIGHT
    scale = max(0.5, APP_SETTINGS.get("ui_scaling_pct", 100) / 100.0)
    return max(18, round(base * scale))


def get_nav_button_height():
    return _LARGE_NAV_BUTTON_HEIGHT if APP_SETTINGS.get("larger_click_targets") else _BASE_NAV_BUTTON_HEIGHT
    return APP_SETTINGS


def get_refresh_interval_ms():
    """Single source of truth for how often the UI's periodic refresh
    loop (Dashboard, Receiver/Pusher monitoring, Queue, Logs, Destination
    Health, Statistics, ...) fires. Every app.after() reschedule of
    periodic_refresh()/pump_events() calls this instead of the old
    hardcoded REFRESH_INTERVAL_MS constant."""
    try:
        return int(APP_SETTINGS.get("refresh_rate_ms", REFRESH_INTERVAL_MS))
    except Exception:
        return REFRESH_INTERVAL_MS


def get_max_worker_threads():
    try:
        return max(1, min(int(APP_SETTINGS.get("max_worker_threads", DEFAULT_PUSH_WORKER_THREADS)),
                           MAX_PUSH_WORKER_THREADS))
    except Exception:
        return DEFAULT_PUSH_WORKER_THREADS


def get_max_simultaneous_pushes():
    try:
        return max(1, min(int(APP_SETTINGS.get("max_simultaneous_pushes", DEFAULT_PUSH_WORKER_THREADS)),
                           MAX_PUSH_WORKER_THREADS))
    except Exception:
        return DEFAULT_PUSH_WORKER_THREADS


def get_network_timeout_sec():
    try:
        return max(1, int(APP_SETTINGS.get("network_timeout_sec", PUSH_ASSOC_TIMEOUT_SEC)))
    except Exception:
        return PUSH_ASSOC_TIMEOUT_SEC


def get_cache_size_limit_bytes():
    try:
        return max(1, int(APP_SETTINGS.get("cache_size_limit_mb", LOG_ROTATION_MAX_BYTES // (1024 * 1024)))) * 1024 * 1024
    except Exception:
        return LOG_ROTATION_MAX_BYTES


def toasts_enabled():
    return bool(APP_SETTINGS.get("toast_notifications_enabled", True))


def notification_duration_ms():
    try:
        return max(1, int(APP_SETTINGS.get("notification_duration_sec", 4))) * 1000
    except Exception:
        return 4000


def apply_performance_mode(mode_name):
    """Seeds refresh_rate_ms / max_worker_threads / max_simultaneous_pushes
    from PERFORMANCE_MODE_PRESETS for the chosen mode, persists, and
    returns the updated settings dict. Individual sliders can still be
    changed afterward -- this is a starting point, not a lock."""
    preset = PERFORMANCE_MODE_PRESETS.get(mode_name)
    if preset is None:
        return APP_SETTINGS
    APP_SETTINGS["performance_mode"] = mode_name
    APP_SETTINGS.update(preset)
    save_app_settings(APP_SETTINGS)
    return APP_SETTINGS


def apply_logging_level(level_name=None):
    """Applies the chosen logging level to every logger the app writes
    through, immediately (no restart needed)."""
    level_name = level_name or APP_SETTINGS.get("logging_level", "Information")
    level = _LOG_LEVEL_TO_LOGGING.get(level_name, logging.INFO)
    try:
        app_logger.setLevel(level)
        for h in app_logger.handlers:
            h.setLevel(level)
    except Exception:
        pass
    try:
        logging.getLogger("pynetdicom").setLevel(max(level, logging.WARNING))
    except Exception:
        pass


def apply_ui_scaling():
    """Applies UI Scaling % and Font Size % immediately -- both are
    live-appliable via CTk's own scaling API, no restart required.

    ctk.set_widget_scaling only scales CustomTkinter widgets -- it has no
    effect on plain ttk widgets like the Receiver/Pusher worklist
    Treeviews, so those need their style (font, row height) and column
    widths rescaled by hand here too."""
    scale = max(0.5, APP_SETTINGS.get("ui_scaling_pct", 100) / 100.0)
    try:
        ctk.set_widget_scaling(scale)
    except Exception:
        pass
    try:
        style = ttk.Style()
        base_font_size = 11
        style.configure("Treeview", rowheight=get_worklist_row_height(),
                         font=(FONT_FAMILY, max(7, round(base_font_size * scale))))
        style.configure("Treeview.Heading",
                         font=(FONT_FAMILY, max(7, round(base_font_size * scale)), "bold"))
        # Receiver/Pusher worklists use their own style (see
        # build_worklist_tree) -- keep its size in sync too. Colors are
        # handled separately (build_worklist_tree / refresh_ui_theme).
        style.configure(WORKLIST_TREE_STYLE, rowheight=get_worklist_row_height() + 6,
                         font=(FONT_FAMILY, max(7, round(base_font_size * scale)) + 1))
        style.configure(f"{WORKLIST_TREE_STYLE}.Heading",
                         font=(FONT_FAMILY, max(7, round(base_font_size * scale)) + 1, "bold"))
    except Exception:
        pass
    # Worklist column widths scale too, so headings/values keep fitting
    # their text at any zoom level instead of the columns staying a fixed
    # pixel width while everything else on screen grows or shrinks.
    for _tree_name in ("rec_tree", "push_tree"):
        _tree = globals().get(_tree_name)
        if _tree is None:
            continue
        try:
            saved_widths = load_worklist_column_widths().get(_tree_name, {})
            for col in WL_COLUMNS:
                if col in saved_widths:
                    # User explicitly resized this column already -- respect
                    # their exact choice rather than compounding scale on
                    # top of it every time this runs.
                    _tree.column(col, width=saved_widths[col])
                else:
                    _tree.column(col, width=max(40, round(WL_WIDTHS[col] * scale)))
        except Exception:
            pass


def get_font_scale_multiplier():
    try:
        return max(0.5, APP_SETTINGS.get("font_scale_pct", 100) / 100.0)
    except Exception:
        return 1.0


class BandwidthLimiter:
    """Thread-safe token-bucket throttle. throttle(num_bytes) blocks
    (sleeps) just enough to keep the long-run average transfer rate at
    or under the configured limit. A no-op when the limit is Unlimited.
    Never freezes the UI -- this is only ever called from background
    push/import/export worker threads, never the main thread."""

    def __init__(self):
        self._lock = threading.RLock()
        self._bytes_per_sec = 0.0
        self._tokens = 0.0
        self._last_refill = time.time()
        self.reload()

    def reload(self):
        mbps = get_effective_bandwidth_mbps()
        with self._lock:
            self._bytes_per_sec = (mbps * 1_000_000) / 8.0 if mbps > 0 else 0.0
            self._tokens = self._bytes_per_sec  # start with a full bucket
            self._last_refill = time.time()

    def throttle(self, num_bytes):
        if num_bytes <= 0:
            return
        while True:
            with self._lock:
                limit = self._bytes_per_sec
                if limit <= 0:
                    return  # unlimited
                now = time.time()
                elapsed = now - self._last_refill
                self._tokens = min(limit, self._tokens + elapsed * limit)
                self._last_refill = now
                if self._tokens >= num_bytes:
                    self._tokens -= num_bytes
                    return
                deficit = num_bytes - self._tokens
                sleep_time = deficit / limit
            # Sleep in short slices (rather than one long sleep) so a
            # config change (e.g. admin switches to Unlimited mid-transfer)
            # takes effect quickly instead of only on the next file.
            time.sleep(min(sleep_time, 0.5))


bandwidth_limiter = BandwidthLimiter()

# =========================================================
# DOC-TRANSFER LIVE PROGRESS (Part 3 of the doc-transfer audit)
# =========================================================
# One entry per in-flight push_patient_documents() call, keyed by a
# unique transfer_id. Deliberately NOT a blocking modal -- this function
# is called from several background contexts (offline queue worker, routing-rule
# auto-push, manual resend) and a modal progress dialog would be wrong
# for the automatic ones. Instead this is a plain, thread-safe dict any
# UI can poll (the Doc Transfer tab does, on its existing refresh timer)
# to render live bars for whichever transfers happen to be running right
# now, without the send loop itself needing to know anything about Tk.
doc_transfer_progress_lock = threading.Lock()
doc_transfer_active_transfers = {}  # transfer_id -> dict, see _dt_progress_start()


def _dt_progress_start(pid, dest_name, files_total, bytes_total):
    transfer_id = f"{pid}:{dest_name}:{time.time()}"
    with doc_transfer_progress_lock:
        doc_transfer_active_transfers[transfer_id] = {
            "pid": pid, "destination_name": dest_name,
            "files_total": files_total, "files_done": 0,
            "current_filename": "",
            "bytes_total": max(1, bytes_total),  # avoid /0 for an all-empty-files edge case
            "bytes_sent": 0,
            "started_at": time.time(),
            "status": "running",
            "cancel_event": threading.Event(),
        }
    return transfer_id


def _dt_progress_update(transfer_id, **fields):
    with doc_transfer_progress_lock:
        entry = doc_transfer_active_transfers.get(transfer_id)
        if entry:
            entry.update(fields)


def _dt_progress_finish(transfer_id, status):
    """Marks the transfer done rather than deleting it immediately, so a
    fast UI poll still gets to show '100% / Done' or '/ Cancelled' for
    one refresh cycle instead of the row just vanishing mid-glance."""
    with doc_transfer_progress_lock:
        entry = doc_transfer_active_transfers.get(transfer_id)
        if entry:
            entry["status"] = status
            entry["finished_at"] = time.time()


def _dt_progress_sweep_finished(max_age_sec=8):
    """Drops entries that finished more than max_age_sec ago. Called
    from the same UI poll that reads this dict -- never runs on its own
    timer, so there's no extra background thread to reason about."""
    now = time.time()
    with doc_transfer_progress_lock:
        stale = [tid for tid, e in doc_transfer_active_transfers.items()
                 if e.get("status") != "running" and now - e.get("finished_at", now) > max_age_sec]
        for tid in stale:
            del doc_transfer_active_transfers[tid]


def cancel_doc_transfer(transfer_id):
    with doc_transfer_progress_lock:
        entry = doc_transfer_active_transfers.get(transfer_id)
        if entry:
            entry["cancel_event"].set()


# =========================================================
# BUILT-IN ZIP EXPORT
# =========================================================
# Supports exporting an entire patient, a single study, a single series,
# an explicit list of files, or the entire worklist, each optionally as
# a password-protected (AES-256, via pyzipper) ZIP, and optionally
# bundling the patient's report, a filtered log excerpt, a metadata
# summary, and/or a real DICOMDIR (via pydicom's FileSet). Runs entirely
# in a background thread with progress callbacks -- never blocks the UI.

def _filename_safe_pid(pid):
    """Filesystem-safe patient-ID fragment. This is the single
    sanitization routine for every place a (network-supplied, therefore
    untrusted) patient ID reaches the filesystem: get_patient_folder()
    below funnels DICOM C-STORE, folder import, and the doc-transfer
    (RDOC) protocol all through this one check, so fixing it here fixes
    all three at once instead of duplicating the same regex in each
    call site.
    Guards against:
    - path separators (/ and \\) and drive/UNC syntax (: ) -- collapsed
      to '_', so a raw pid can never introduce a new path segment.
    - a result of only dots ('.', '..', '...') -- every individual
      character in '..' passes the [A-Za-z0-9._-] allowlist below, so
      that has to be rejected explicitly or a pid of exactly '..' would
      still resolve to OUTPUT_DIR's parent.
    - empty / whitespace-only input -- falls back to 'UNKNOWN'.
    - pathological length -- capped well under OS path-length limits so
      an oversized pid fails cleanly here instead of crashing a later
      os.makedirs()/open() deep in the receive path.
    """
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(pid)).strip("_")
    if not safe or safe.strip(".") == "":
        return "UNKNOWN"
    return safe[:128]


def get_patient_folder(pid):
    return os.path.join(OUTPUT_DIR, _filename_safe_pid(pid))


def list_patient_dcm_files(pid):
    folder = get_patient_folder(pid)
    if not os.path.isdir(folder):
        return []
    return sorted(
        os.path.join(folder, f) for f in os.listdir(folder)
        if f.lower().endswith(".dcm") and os.path.isfile(os.path.join(folder, f))
    )


def build_file_index_for_patient(pid):
    """Reads each file's headers (not pixel data, so this stays fast even
    for large studies) to build a small index used by the Export tab's
    Study/Series pickers and by the metadata.json export option."""
    index = []
    for fpath in list_patient_dcm_files(pid):
        try:
            ds = pydicom.dcmread(fpath, stop_before_pixels=True, force=True)
            index.append({
                "path": fpath,
                "sop_uid": os.path.splitext(os.path.basename(fpath))[0],
                "study_uid": str(getattr(ds, "StudyInstanceUID", "")),
                "series_uid": str(getattr(ds, "SeriesInstanceUID", "")),
                "series_description": str(getattr(ds, "SeriesDescription", "")),
                "modality": str(getattr(ds, "Modality", "")),
                "size": os.path.getsize(fpath),
            })
        except Exception:
            log_exception(f"Failed to read metadata for export index: {fpath}")
    return index


def resolve_export_file_list(options):
    """Turns an export options dict into a concrete, deduplicated list of
    (pid, absolute_file_path) tuples covering exactly the requested scope."""
    scope = options["scope"]
    pairs = []

    if scope == "worklist":
        with data_lock:
            pids = list(patient_data.keys())
        for pid in pids:
            for fpath in list_patient_dcm_files(pid):
                pairs.append((pid, fpath))

    elif scope == "patient":
        pid = options["pid"]
        for fpath in list_patient_dcm_files(pid):
            pairs.append((pid, fpath))

    elif scope == "study":
        pid = options["pid"]
        for entry in build_file_index_for_patient(pid):
            if entry["study_uid"] == options["study_uid"]:
                pairs.append((pid, entry["path"]))

    elif scope == "series":
        pid = options["pid"]
        for entry in build_file_index_for_patient(pid):
            if entry["series_uid"] == options["series_uid"]:
                pairs.append((pid, entry["path"]))

    elif scope == "files":
        pid = options["pid"]
        for fpath in options.get("file_paths", []):
            pairs.append((pid, fpath))

    return pairs


def _ensure_dicomdir_required_fields(ds):
    """pydicom's FileSet.add() requires a handful of type-2 attributes
    (Study Date/Time/ID, Series Number, Instance Number) to build the
    standard directory records. Real scanner output almost always has
    these, but anonymized or hand-built secondary-capture files sometimes
    don't -- rather than silently dropping those instances from the
    DICOMDIR (which would make Include DICOMDIR quietly incomplete),
    fill in DICOM-legal empty/placeholder values so every instance in
    the export actually makes it into the directory."""
    if not hasattr(ds, "StudyDate") or not ds.StudyDate:
        ds.StudyDate = "19700101"
    if not hasattr(ds, "StudyTime") or not ds.StudyTime:
        ds.StudyTime = "000000"
    if not hasattr(ds, "StudyID") or not ds.StudyID:
        ds.StudyID = "1"
    if not hasattr(ds, "SeriesNumber") or ds.SeriesNumber in (None, ""):
        ds.SeriesNumber = 1
    if not hasattr(ds, "InstanceNumber") or ds.InstanceNumber in (None, ""):
        ds.InstanceNumber = 1
    return ds


def _build_dicomdir_staging(pairs, staging_root):
    """Uses pydicom's FileSet to build a real, spec-compliant DICOMDIR
    plus the standard PACS-media directory layout, so the exported ZIP
    can be opened directly by any DICOMDIR-aware viewer."""
    from pydicom.fileset import FileSet
    fs = FileSet()
    added = 0
    for pid, fpath in pairs:
        try:
            ds = pydicom.dcmread(fpath, force=True)
            _ensure_dicomdir_required_fields(ds)
            fs.add(ds)
            added += 1
        except Exception:
            log_exception(f"Failed to add {fpath} to DICOMDIR file-set")
    if added == 0:
        raise ValueError("None of the selected files could be added to a DICOMDIR (unreadable or incompatible).")
    fs.write(staging_root)


def _write_export_metadata_json(pairs, dest_root_in_zip, zf_write):
    """Writes a metadata.json summarizing exactly what's in this export:
    per-patient worklist record + a flat file index (SOP/Study/Series
    UID, modality, size). zf_write(arcname, data_bytes) abstracts over
    plain zipfile vs pyzipper so this works for both."""
    by_pid = {}
    for pid, fpath in pairs:
        by_pid.setdefault(pid, []).append(fpath)

    metadata = {"exported_at": datetime.datetime.now().isoformat(timespec="seconds"), "patients": {}}
    for pid, paths in by_pid.items():
        with data_lock:
            record = dict(patient_data.get(pid, {}))
        file_entries = []
        for fpath in paths:
            try:
                ds = pydicom.dcmread(fpath, stop_before_pixels=True, force=True)
                file_entries.append({
                    "sop_instance_uid": str(getattr(ds, "SOPInstanceUID", "")),
                    "study_uid": str(getattr(ds, "StudyInstanceUID", "")),
                    "series_uid": str(getattr(ds, "SeriesInstanceUID", "")),
                    "modality": str(getattr(ds, "Modality", "")),
                    "size_bytes": os.path.getsize(fpath),
                })
            except Exception:
                file_entries.append({"path": os.path.basename(fpath), "error": "unreadable"})
        metadata["patients"][pid] = {"worklist_record": record, "files": file_entries}

    zf_write(f"{dest_root_in_zip}metadata.json", json.dumps(metadata, indent=2, default=str).encode("utf-8"))


def _write_export_log_excerpt(pairs, dest_root_in_zip, zf_write):
    """Filters the structured JSONL logs down to just the patients in
    this export -- a full global log dump wouldn't make sense attached
    to one patient/study export, so this pulls only relevant entries."""
    pids = {pid for pid, _ in pairs}
    for log_path, label in ((RECEIVER_LOG_JSONL, "receiver"), (PUSH_LOG_JSONL, "push")):
        if not os.path.exists(log_path):
            continue
        matching_lines = []
        try:
            with open(log_path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    try:
                        rec = json.loads(line)
                    except Exception:
                        continue
                    if rec.get("patient_id") in pids:
                        matching_lines.append(line)
        except Exception:
            log_exception(f"Failed to filter {log_path} for export log excerpt")
            continue
        if matching_lines:
            zf_write(f"{dest_root_in_zip}logs/{label}_events_excerpt.jsonl", "".join(matching_lines).encode("utf-8"))


def run_export_job(options, progress_cb=None):
    """Does the actual export work. Always call from a background thread.
    `options` keys: scope ('patient'/'study'/'series'/'files'/'worklist'),
    pid, study_uid, series_uid, file_paths, dest_path, password,
    include_reports, include_logs, include_metadata, include_dicomdir.
    progress_cb(done, total, message) is called from THIS thread -- the
    caller is responsible for hopping back to the UI thread via app.after."""
    pairs = resolve_export_file_list(options)
    total_steps = len(pairs) + 1  # +1 for the final zip-write pass housekeeping
    if not pairs:
        raise ValueError("Nothing matched the selected export scope -- check the patient/study/series selection.")

    dest_path = options["dest_path"]
    password = options.get("password") or None
    use_password = bool(password)
    if use_password and not PYZIPPER_AVAILABLE:
        raise RuntimeError("Password-protected export requires the 'pyzipper' package. Install it with: pip install pyzipper")

    staging_dir = tempfile.mkdtemp(prefix="rapps_export_")
    try:
        # DICOMDIR needs its own directory layout (FileSet writes files +
        # DICOMDIR into staging_dir itself), everything else is added to
        # the ZIP directly from the original patient folders.
        if options.get("include_dicomdir"):
            if progress_cb:
                progress_cb(0, total_steps, "Building DICOMDIR…")
            _build_dicomdir_staging(pairs, staging_dir)

        if use_password:
            zf = pyzipper.AESZipFile(dest_path, "w", compression=pyzipper.ZIP_DEFLATED,
                                      encryption=pyzipper.WZ_AES)
            zf.setpassword(password.encode("utf-8"))
            zf_write = lambda arcname, data: zf.writestr(arcname, data)
            zf_write_file = lambda arcname, path: zf.write(path, arcname=arcname)
        else:
            zf = zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED)
            zf_write = lambda arcname, data: zf.writestr(arcname, data)
            zf_write_file = lambda arcname, path: zf.write(path, arcname=arcname)

        try:
            done = 0
            for pid, fpath in pairs:
                try:
                    bandwidth_limiter.throttle(os.path.getsize(fpath))
                except OSError:
                    pass
                arcname = f"{_filename_safe_pid(pid)}/{os.path.basename(fpath)}"
                zf_write_file(arcname, fpath)
                done += 1
                if progress_cb:
                    progress_cb(done, total_steps, f"Adding {pid}/{os.path.basename(fpath)}")

            pids_in_export = sorted({pid for pid, _ in pairs})

            if options.get("include_reports"):
                if progress_cb:
                    progress_cb(done, total_steps, "Adding reports…")
                for pid in pids_in_export:
                    rpath = get_report_path(pid)
                    if os.path.isfile(rpath):
                        zf_write_file(f"{_filename_safe_pid(pid)}/Reports/{os.path.basename(rpath)}", rpath)

            if options.get("include_logs"):
                if progress_cb:
                    progress_cb(done, total_steps, "Adding log excerpt…")
                _write_export_log_excerpt(pairs, "", zf_write)

            if options.get("include_metadata"):
                if progress_cb:
                    progress_cb(done, total_steps, "Adding metadata.json…")
                _write_export_metadata_json(pairs, "", zf_write)

            if options.get("include_dicomdir"):
                if progress_cb:
                    progress_cb(done, total_steps, "Adding DICOMDIR…")
                dicomdir_path = os.path.join(staging_dir, "DICOMDIR")
                if os.path.isfile(dicomdir_path):
                    # DICOMDIR's internal file references are relative to
                    # its own location (e.g. "PT000000\ST000000\..."), so
                    # it and the PATIENT/STUDY/SERIES/IMAGE tree
                    # FileSet.write() produced must stay siblings -- both
                    # go under one shared subfolder here rather than
                    # DICOMDIR at the zip root with the tree nested deeper
                    # (which would break every reference inside it).
                    for root, _dirs, files in os.walk(staging_dir):
                        for fname in files:
                            full = os.path.join(root, fname)
                            rel = os.path.join("DICOMDIR_export", os.path.relpath(full, staging_dir))
                            zf_write_file(rel, full)

            if progress_cb:
                progress_cb(total_steps, total_steps, "Finalizing ZIP…")
        finally:
            zf.close()
    finally:
        shutil.rmtree(staging_dir, ignore_errors=True)

    write_audit_log(
        "EXPORT",
        f"scope={options['scope']} patients={len(pids_in_export)} "
        f"files={len(pairs)} password_protected={use_password} dest={dest_path}"
    )
    return len(pairs), pids_in_export


# =========================================================
# PDF REPORT GENERATOR
# =========================================================
# Reports are built from the structured JSONL logs (receiver_events /
# push_events), NOT from in-memory daily_stats -- daily_stats only knows
# "today", but a report needs to cover a date range that could span
# rotated/archived logs too, so this reads both the live JSONL files and
# any matching archives in LOG_ARCHIVE_DIR.

def _parse_jsonl_text_in_range(text, start_dt, end_dt):
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rec = json.loads(line)
            ts = datetime.datetime.fromisoformat(rec.get("timestamp", ""))
        except Exception:
            continue
        if start_dt <= ts <= end_dt:
            yield rec


def _iter_jsonl_records_in_range(base_filename, start_dt, end_dt):
    """Yields every parsed record with a timestamp in [start_dt, end_dt]
    from the live log file AND any rotated archive of it -- so a Monthly
    or Custom Date Range report still works correctly even after the
    immutable-log rotation system has archived older entries away."""
    if os.path.exists(base_filename):
        try:
            with open(base_filename, "r", encoding="utf-8", errors="replace") as f:
                yield from _parse_jsonl_text_in_range(f.read(), start_dt, end_dt)
        except Exception:
            log_exception(f"Failed reading {base_filename} for report")

    base_stem = os.path.splitext(os.path.basename(base_filename))[0]
    if os.path.isdir(LOG_ARCHIVE_DIR):
        for fname in os.listdir(LOG_ARCHIVE_DIR):
            if not (fname.startswith(base_stem) and fname.lower().endswith(".zip")):
                continue
            archive_path = os.path.join(LOG_ARCHIVE_DIR, fname)
            try:
                with zipfile.ZipFile(archive_path, "r") as zf:
                    for name in zf.namelist():
                        if name == os.path.basename(base_filename):
                            with zf.open(name) as member:
                                text = member.read().decode("utf-8", errors="replace")
                                yield from _parse_jsonl_text_in_range(text, start_dt, end_dt)
            except Exception:
                log_exception(f"Failed reading archived log {archive_path} for report")


LOG_TEXT_TO_JSONL = {
    RECEIVER_LOG: RECEIVER_LOG_JSONL,
    PUSH_LOG: PUSH_LOG_JSONL,
    AUDIT_LOG: AUDIT_LOG_JSONL,
    APP_LOG: APP_LOG_JSONL,
}

LOG_STRUCT_RANGE_OPTIONS = ["Today", "Last 7 Days", "Last 30 Days", "All Time"]


def _date_range_for_log_option(option):
    """Same semantics as the worklist's date filter (Today / Last 7 Days /
    Last 30 Days / All Time), but returning (start_dt, end_dt) datetimes
    for _iter_jsonl_records_in_range."""
    now = datetime.datetime.now()
    if option == "Today":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif option == "Last 7 Days":
        start = now - datetime.timedelta(days=7)
    elif option == "Last 30 Days":
        start = now - datetime.timedelta(days=30)
    else:  # All Time
        start = datetime.datetime(2000, 1, 1)
    return start, now


def _log_record_is_failure(rec):
    """The three jsonl schemas don't share one field name for outcome:
    receiver events use result=SUCCESS/FAILURE, push events use
    final_status=OK/FAILED, audit events use result=INFO/... plus an
    optional explicit result. This normalizes all three."""
    if rec.get("result") == "FAILURE":
        return True
    if rec.get("final_status") == "FAILED":
        return True
    if rec.get("error_message"):
        return True
    return False


def get_structured_log_records(module_text_path, search_text="", destination="All",
                                severity="All", range_option="All Time", max_records=500,
                                event_type="All"):
    """Structured (parsed-jsonl) equivalent of the Logs tab's raw text
    tail. Reuses _iter_jsonl_records_in_range exactly as the PDF Report
    feature already does (live file + rotated archives), so this is a
    second consumer of existing infrastructure, not a new log source.

    B.3: `event_type` matches either the push/receiver record's own
    "event_type" field (e.g. "DOC-TRANSFER") or an audit record's
    "action" field (e.g. "DOC-TRANSFER-OK") -- whichever the given log
    file actually populates -- so DOC-TRANSFER* values work as filters
    across the Push, Receiver, and Audit log files alike."""
    jsonl_file = LOG_TEXT_TO_JSONL.get(module_text_path, RECEIVER_LOG_JSONL)
    start_dt, end_dt = _date_range_for_log_option(range_option)
    records = list(_iter_jsonl_records_in_range(jsonl_file, start_dt, end_dt))

    search_lower = (search_text or "").strip().lower()
    out = []
    for rec in records:
        is_failure = _log_record_is_failure(rec)
        if severity == "Errors Only" and not is_failure:
            continue
        if severity == "Success Only" and is_failure:
            continue
        if destination != "All" and rec.get("destination_name") != destination:
            continue
        if event_type != "All" and rec.get("event_type") != event_type and rec.get("action") != event_type:
            continue
        if search_lower:
            haystack = " ".join(str(rec.get(k, "")) for k in (
                "patient_id", "patient_name", "institution", "destination_name",
                "error_message", "event", "action", "details")).lower()
            if search_lower not in haystack:
                continue
        out.append(rec)

    out.sort(key=lambda r: r.get("timestamp", ""), reverse=True)
    return out[:max_records]


def compute_report_stats(start_dt, end_dt):
    """Aggregates everything the PDF Report (and, later, the Performance
    Metrics page) needs for a given date range."""
    receiver_records = list(_iter_jsonl_records_in_range(RECEIVER_LOG_JSONL, start_dt, end_dt))
    push_records = list(_iter_jsonl_records_in_range(PUSH_LOG_JSONL, start_dt, end_dt))

    ok_receives = [r for r in receiver_records if r.get("result") == "SUCCESS"]
    studies_received = len({r["study_uid"] for r in ok_receives if r.get("study_uid")})
    # Duplicate C-STORE events (a study re-pushed from the same or another
    # sender) are still logged for audit purposes -- see the "duplicate"
    # field written by handle_store()/import_folder() -- but must not be
    # counted as additional received images, or this figure would inflate
    # the same way the worklist "#" column used to.
    unique_ok_receives = [r for r in ok_receives if not r.get("duplicate")]
    images_received = len(unique_ok_receives)

    studies_sent_uids = set()
    images_sent = 0
    images_failed = 0
    failed_push_count = 0
    transfer_speeds = []
    dest_stats = defaultdict(lambda: {"sent": 0, "failed": 0})

    for r in push_records:
        dest_name = r.get("destination_name") or "(unknown)"
        status = r.get("final_status")
        sent_n = r.get("images_sent", 0) or 0
        failed_n = r.get("images_failed", 0) or 0
        images_sent += sent_n
        images_failed += failed_n
        dest_stats[dest_name]["sent"] += sent_n
        if status != "OK":
            failed_push_count += 1
            dest_stats[dest_name]["failed"] += failed_n or (r.get("images_attempted", 0) or 0)
        elif r.get("study_uid"):
            studies_sent_uids.add(r["study_uid"])
        speed = r.get("transfer_speed_files_per_sec")
        if isinstance(speed, (int, float)) and speed > 0:
            transfer_speeds.append(speed)

    total_push_attempts = images_sent + images_failed
    success_rate = (images_sent / total_push_attempts * 100) if total_push_attempts else 0.0
    avg_speed = (sum(transfer_speeds) / len(transfer_speeds)) if transfer_speeds else 0.0

    top_modalities = Counter(r.get("modality") for r in unique_ok_receives if r.get("modality")).most_common(5)
    top_institutions = Counter(r.get("institution") for r in unique_ok_receives if r.get("institution")).most_common(5)

    queue_size, oldest_item, next_retry, interval = get_offline_queue_summary()

    try:
        total_b, used_b, free_b = shutil.disk_usage(os.path.abspath("."))
        disk_used_pct = (used_b / total_b * 100) if total_b else 0.0
        disk_free_gb = free_b / (1024 ** 3)
    except Exception:
        disk_used_pct, disk_free_gb = 0.0, 0.0

    uptime_sec = int(time.time() - APP_START_TIME)

    return {
        "start_dt": start_dt, "end_dt": end_dt,
        "studies_received": studies_received, "images_received": images_received,
        "studies_sent": len(studies_sent_uids), "images_sent": images_sent,
        "images_failed": images_failed, "failed_push_count": failed_push_count,
        "success_rate": success_rate, "avg_transfer_speed": avg_speed,
        "top_modalities": top_modalities, "top_institutions": top_institutions,
        "destination_stats": dict(dest_stats),
        "queue_size": queue_size, "queue_oldest": oldest_item,
        "disk_used_pct": disk_used_pct, "disk_free_gb": disk_free_gb,
        "uptime_sec": uptime_sec,
    }


def _fmt_uptime(uptime_sec):
    hours, rem = divmod(int(uptime_sec), 3600)
    minutes, _ = divmod(rem, 60)
    days, hours = divmod(hours, 24)
    return f"{days}d {hours}h {minutes}m" if days else f"{hours}h {minutes}m"


# =========================================================
# PERFORMANCE METRICS
# =========================================================

_perf_net_io_baseline = {"time": None, "bytes_sent": None, "bytes_recv": None}
_perf_disk_io_baseline = {"time": None, "read_bytes": None, "write_bytes": None}


def compute_doc_transfer_stats(start, end):
    """B.1/B.2/B.4/B.5/B.6 shared helper: derives all document-transfer
    analytics from the existing structured push/receiver JSONL logs
    (filtered to event_type == "DOC-TRANSFER") -- no new counters are
    added to push_job/receiver_state, consistent with how every other
    perf card is already computed."""
    push_records = [r for r in _iter_jsonl_records_in_range(PUSH_LOG_JSONL, start, end)
                    if r.get("event_type") == "DOC-TRANSFER"]
    attempts = len(push_records)
    successes = sum(1 for r in push_records if r.get("final_status") == "OK")
    failures = attempts - successes
    durations = [r["transfer_time_sec"] for r in push_records
                 if isinstance(r.get("transfer_time_sec"), (int, float))]
    docs_sent_total = sum(int(r.get("files_sent") or 0) for r in push_records)
    docs_failed_total = sum(int(r.get("files_failed") or 0) for r in push_records)
    elapsed_sec = max((end - start).total_seconds(), 1.0)

    by_destination = {}
    for r in push_records:
        name = r.get("destination_name") or "(unknown)"
        entry = by_destination.setdefault(name, {"docs_sent": 0, "docs_failed": 0})
        entry["docs_sent"] += int(r.get("files_sent") or 0)
        entry["docs_failed"] += int(r.get("files_failed") or 0)

    return {
        "attempts": attempts,
        "successes": successes,
        "failures": failures,
        "docs_sent_total": docs_sent_total,
        "docs_failed_total": docs_failed_total,
        "docs_per_sec": docs_sent_total / elapsed_sec,
        "success_rate_pct": (successes / attempts * 100.0) if attempts else 100.0,
        "avg_transfer_time_sec": (sum(durations) / len(durations)) if durations else 0.0,
        "by_destination": by_destination,
    }


def compute_doc_transfer_activity_stats(start, end):
    """Part 3 (doc-transfer audit) dashboard cards: per-FILE granularity
    (event_type == DOC-TRANSFER-FILE on the push side, DOC-TRANSFER on
    the receive side -- both are one row per file). Deliberately
    separate from compute_doc_transfer_stats() above, which is per-BATCH
    and already relied on by the Performance tab -- mixing the two would
    either double-count or require touching working code unnecessarily."""
    sent = [r for r in _iter_jsonl_records_in_range(PUSH_LOG_JSONL, start, end)
            if r.get("event_type") == "DOC-TRANSFER-FILE"]
    received = [r for r in _iter_jsonl_records_in_range(RECEIVER_LOG_JSONL, start, end)
                if r.get("event_type") == "DOC-TRANSFER"]
    all_records = sent + received

    total = len(all_records)
    succeeded = sum(1 for r in all_records if r.get("result") == "SUCCESS" or r.get("final_status") == "OK")
    failed = total - succeeded
    speeds = [r["transfer_speed_mbps"] for r in all_records if isinstance(r.get("transfer_speed_mbps"), (int, float))]
    durations = [r["duration_sec"] for r in all_records if isinstance(r.get("duration_sec"), (int, float))]
    retries_total = sum(int(r.get("retry_count") or 0) for r in sent)
    sizes = [r["file_size"] for r in all_records if isinstance(r.get("file_size"), (int, float))]
    largest = max(sizes) if sizes else 0

    # "Pending" reuses the same last_doc_transfer_error signal the
    # worklist's Docs column already relies on (see refresh_inspector /
    # populate_tree) -- a patient with a live doc-transfer error has
    # local documents that haven't successfully made it out yet.
    with data_lock:
        pending = sum(1 for d in patient_data.values() if d.get("last_doc_transfer_error"))

    return {
        "documents_today": total,
        "success_rate_pct": (succeeded / total * 100.0) if total else 100.0,
        "avg_speed_mbps": (sum(speeds) / len(speeds)) if speeds else 0.0,
        "avg_duration_sec": (sum(durations) / len(durations)) if durations else 0.0,
        "failed": failed,
        "pending": pending,
        "retries": retries_total,
        "largest_transfer_bytes": largest,
        "sent_records": sent,
        "received_records": received,
    }


DOC_TRANSFER_FAILURE_RATE_ALERT_THRESHOLD_PCT = 30.0
DOC_TRANSFER_FAILURE_RATE_ALERT_MIN_ATTEMPTS = 3
_doc_transfer_failure_alert_throttle = [0.0]


def _check_doc_transfer_failure_rate_alert():
    """B.6: fires a notify_event() if the rolling doc-transfer failure
    rate over the last hour crosses a threshold. Piggybacks entirely on
    refresh_performance_tab's existing refresh cadence -- no new polling
    thread. Throttled to at most one notification per 15 minutes so a
    sustained outage doesn't spam the notification channels."""
    try:
        end = datetime.datetime.now()
        start = end - datetime.timedelta(hours=1)
        stats = compute_doc_transfer_stats(start, end)
        if stats["attempts"] < DOC_TRANSFER_FAILURE_RATE_ALERT_MIN_ATTEMPTS:
            return
        failure_rate_pct = (stats["failures"] / stats["attempts"]) * 100.0
        if failure_rate_pct < DOC_TRANSFER_FAILURE_RATE_ALERT_THRESHOLD_PCT:
            return
        now = time.time()
        if now - _doc_transfer_failure_alert_throttle[0] < 900:  # 15 min
            return
        _doc_transfer_failure_alert_throttle[0] = now
        notify_event(
            "doc_transfer_failure_rate",
            "Document Transfer Failure Rate High",
            f"{failure_rate_pct:.0f}% of document transfers failed in the last hour "
            f"({stats['failures']}/{stats['attempts']} attempts).",
        )
    except Exception:
        log_exception("Failed to evaluate doc-transfer failure-rate alert")


def compute_live_performance_metrics(window_minutes=60):
    """Combines a recent rolling-window read of the structured logs (for
    throughput/timing averages) with live psutil sampling (for CPU/RAM/
    disk-I/O/network) into the full Performance Metrics stat set."""
    end = datetime.datetime.now()
    start = end - datetime.timedelta(minutes=window_minutes)
    stats = compute_report_stats(start, end)
    elapsed_sec = max((end - start).total_seconds(), 1.0)

    receiver_records = list(_iter_jsonl_records_in_range(RECEIVER_LOG_JSONL, start, end))
    push_records = list(_iter_jsonl_records_in_range(PUSH_LOG_JSONL, start, end))

    receive_durations = [r["duration_sec"] for r in receiver_records if isinstance(r.get("duration_sec"), (int, float))]
    push_durations = [r["transfer_time_sec"] for r in push_records if isinstance(r.get("transfer_time_sec"), (int, float))]
    assoc_times = [r["association_time_sec"] for r in push_records if isinstance(r.get("association_time_sec"), (int, float))]

    images_per_sec = (stats["images_received"] + stats["images_sent"]) / elapsed_sec
    studies_per_sec = (stats["studies_received"] + stats["studies_sent"]) / elapsed_sec
    mb_per_sec = stats["avg_transfer_speed"] * 0.5  # rough images/sec -> MB/sec, same heuristic as the Dashboard

    avg_receive_time = (sum(receive_durations) / len(receive_durations)) if receive_durations else 0.0
    avg_push_time = (sum(push_durations) / len(push_durations)) if push_durations else 0.0
    avg_association_time = (sum(assoc_times) / len(assoc_times)) if assoc_times else 0.0
    avg_queue_time = (sum(queue_wait_time_samples) / len(queue_wait_time_samples)) if queue_wait_time_samples else 0.0
    avg_retry_time = (sum(retry_interval_samples) / len(retry_interval_samples)) if retry_interval_samples else 0.0

    cpu_pct = ram_pct = 0.0
    if PSUTIL_AVAILABLE:
        try:
            cpu_pct = psutil.cpu_percent(interval=None)
            ram_pct = psutil.virtual_memory().percent
        except Exception:
            pass

    disk_read_bps = disk_write_bps = 0.0
    if PSUTIL_AVAILABLE:
        try:
            counters = psutil.disk_io_counters()
            now_t = time.time()
            if counters is not None:
                if _perf_disk_io_baseline["time"] is not None:
                    dt = max(now_t - _perf_disk_io_baseline["time"], 0.001)
                    disk_read_bps = max(0.0, (counters.read_bytes - _perf_disk_io_baseline["read_bytes"]) / dt)
                    disk_write_bps = max(0.0, (counters.write_bytes - _perf_disk_io_baseline["write_bytes"]) / dt)
                _perf_disk_io_baseline.update(time=now_t, read_bytes=counters.read_bytes, write_bytes=counters.write_bytes)
        except Exception:
            pass

    net_up_bps = net_down_bps = 0.0
    if PSUTIL_AVAILABLE:
        try:
            counters = psutil.net_io_counters()
            now_t = time.time()
            if _perf_net_io_baseline["time"] is not None:
                dt = max(now_t - _perf_net_io_baseline["time"], 0.001)
                net_up_bps = max(0.0, (counters.bytes_sent - _perf_net_io_baseline["bytes_sent"]) / dt)
                net_down_bps = max(0.0, (counters.bytes_recv - _perf_net_io_baseline["bytes_recv"]) / dt)
            _perf_net_io_baseline.update(time=now_t, bytes_sent=counters.bytes_sent, bytes_recv=counters.bytes_recv)
        except Exception:
            pass

    db_size_bytes = os.path.getsize(CSV_FILE) if os.path.exists(CSV_FILE) else 0

    doc_stats = compute_doc_transfer_stats(start, end)

    return {
        "window_minutes": window_minutes,
        "images_per_sec": images_per_sec, "studies_per_sec": studies_per_sec, "mb_per_sec": mb_per_sec,
        "avg_receive_time": avg_receive_time, "avg_push_time": avg_push_time,
        "avg_association_time": avg_association_time, "avg_queue_time": avg_queue_time,
        "avg_retry_time": avg_retry_time,
        "cpu_pct": cpu_pct, "ram_pct": ram_pct,
        "disk_read_bps": disk_read_bps, "disk_write_bps": disk_write_bps,
        "network_up_bps": net_up_bps, "network_down_bps": net_down_bps,
        "database_size_bytes": db_size_bytes,
        "uptime_sec": int(time.time() - APP_START_TIME),
        "docs_per_sec": doc_stats["docs_per_sec"],
        "doc_success_rate_pct": doc_stats["success_rate_pct"],
        "avg_doc_transfer_time": doc_stats["avg_transfer_time_sec"],
    }


PERF_HISTORY_LEN = 60
perf_history = {
    "images_per_sec": deque(maxlen=PERF_HISTORY_LEN),
    "studies_per_sec": deque(maxlen=PERF_HISTORY_LEN),
    "cpu_pct": deque(maxlen=PERF_HISTORY_LEN),
    "ram_pct": deque(maxlen=PERF_HISTORY_LEN),
    "disk_io_kbps": deque(maxlen=PERF_HISTORY_LEN),
    "network_kbps": deque(maxlen=PERF_HISTORY_LEN),
}


def export_performance_metrics(dest_path):
    """Exports the current metrics snapshot + the rolling history buffers
    behind the graphs to a JSON file the user can archive or analyze
    elsewhere."""
    metrics = compute_live_performance_metrics()
    payload = {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "snapshot": metrics,
        "history": {k: list(v) for k, v in perf_history.items()},
    }
    with open(dest_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)
    write_audit_log("METRICS-EXPORTED", f"dest={dest_path}")


def _render_bar_chart_png(labels, values, title, color="#2f8eff"):
    """Renders a simple bar chart to PNG bytes via matplotlib (headless,
    Agg backend) for embedding in the PDF via reportlab's Image flowable."""
    fig, ax = plt.subplots(figsize=(6, 3), dpi=150)
    if labels:
        ax.bar(labels, values, color=color)
        ax.set_title(title, fontsize=11)
        ax.tick_params(axis="x", labelrotation=30, labelsize=8)
        ax.tick_params(axis="y", labelsize=8)
    else:
        ax.text(0.5, 0.5, "No data in this date range", ha="center", va="center", fontsize=10)
        ax.set_xticks([]); ax.set_yticks([])
        ax.set_title(title, fontsize=11)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf


def generate_report_pdf(start_dt, end_dt, dest_path, report_label):
    """Builds the actual PDF. Always call from a background thread --
    this does real file I/O and chart rendering. Returns the stats dict
    used, in case the caller wants to reuse it (e.g. for an email body)."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("PDF report generation requires 'reportlab' and 'matplotlib'. Install with: pip install reportlab matplotlib")

    stats = compute_report_stats(start_dt, end_dt)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("R-Apps DICOM Receiver + AutoRouter", styles["Title"]))
    elements.append(Paragraph(f"{report_label} Report", styles["Heading2"]))
    elements.append(Paragraph(
        f"{start_dt.strftime('%Y-%m-%d %H:%M')} &ndash; {end_dt.strftime('%Y-%m-%d %H:%M')}"
        f" &nbsp;|&nbsp; Generated {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        styles["Normal"],
    ))
    elements.append(Spacer(1, 0.25 * inch))

    summary_rows = [
        ["Metric", "Value"],
        ["Studies Received", str(stats["studies_received"])],
        ["Images Received", str(stats["images_received"])],
        ["Studies Sent", str(stats["studies_sent"])],
        ["Images Sent", str(stats["images_sent"])],
        ["Failed Studies", str(stats["failed_push_count"])],
        ["Success Rate", f"{stats['success_rate']:.1f}%"],
        ["Average Transfer Speed", f"{stats['avg_transfer_speed']:.2f} images/sec"],
        ["Offline Queue Size", str(stats["queue_size"])],
        ["Oldest Queue Item", stats["queue_oldest"]["queued_at"] if stats["queue_oldest"] else "—"],
        ["Disk Usage", f"{stats['disk_used_pct']:.1f}% used, {stats['disk_free_gb']:.1f} GB free"],
        ["Application Uptime", _fmt_uptime(stats["uptime_sec"])],
    ]
    summary_table = Table(summary_rows, colWidths=[2.8 * inch, 3.2 * inch])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#232733")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 11),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f4f5f7")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.3 * inch))

    # ---- Top Modalities ----
    elements.append(Paragraph("Top Modalities", styles["Heading3"]))
    mod_labels = [m for m, _ in stats["top_modalities"]]
    mod_values = [c for _, c in stats["top_modalities"]]
    chart_buf = _render_bar_chart_png(mod_labels, mod_values, "Images Received by Modality", color="#2f8eff")
    elements.append(RLImage(chart_buf, width=5.5 * inch, height=2.75 * inch))
    elements.append(Spacer(1, 0.2 * inch))

    # ---- Top Institutions ----
    elements.append(Paragraph("Top Institutions", styles["Heading3"]))
    if stats["top_institutions"]:
        inst_rows = [["Institution", "Images Received"]] + [[name, str(c)] for name, c in stats["top_institutions"]]
    else:
        inst_rows = [["Institution", "Images Received"], ["(no data in this date range)", "0"]]
    inst_table = Table(inst_rows, colWidths=[4 * inch, 2 * inch])
    inst_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#232733")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(inst_table)
    elements.append(Spacer(1, 0.3 * inch))
    elements.append(PageBreak())

    # ---- Destination Statistics ----
    elements.append(Paragraph("Destination Statistics", styles["Heading3"]))
    dest_names = list(stats["destination_stats"].keys())
    dest_sent = [stats["destination_stats"][n]["sent"] for n in dest_names]
    dest_failed = [stats["destination_stats"][n]["failed"] for n in dest_names]
    chart_buf2 = _render_bar_chart_png(dest_names, dest_sent, "Images Sent by Destination", color="#2ecc71")
    elements.append(RLImage(chart_buf2, width=5.5 * inch, height=2.75 * inch))
    elements.append(Spacer(1, 0.15 * inch))

    if dest_names:
        doc_stats_by_dest = compute_doc_transfer_stats(start_dt, end_dt)["by_destination"]
        dest_rows = [["Destination", "Sent", "Failed", "Docs Sent", "Docs Failed"]] + [
            [n, str(stats["destination_stats"][n]["sent"]), str(stats["destination_stats"][n]["failed"]),
             str(doc_stats_by_dest.get(n, {}).get("docs_sent", 0)),
             str(doc_stats_by_dest.get(n, {}).get("docs_failed", 0))]
            for n in dest_names
        ]
    else:
        dest_rows = [["Destination", "Sent", "Failed", "Docs Sent", "Docs Failed"],
                    ["(no push activity in this date range)", "0", "0", "0", "0"]]
    dest_table = Table(dest_rows, colWidths=[2.3 * inch, 1 * inch, 1 * inch, 1 * inch, 1 * inch])
    dest_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#232733")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(dest_table)

    doc = SimpleDocTemplate(dest_path, pagesize=letter,
                            title=f"{report_label} Report", author="R-Apps DICOM Receiver + AutoRouter")
    doc.build(elements)

    write_audit_log("REPORT-CREATED", f"type={report_label} range={start_dt.date()}..{end_dt.date()} dest={dest_path}")
    return stats


def print_pdf_file(path):
    """Best-effort cross-platform print-to-default-printer. Returns
    (ok, message)."""
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path, "print")
        elif system == "Darwin":
            subprocess.run(["lpr", path], check=True)
        else:
            subprocess.run(["lpr", path], check=True)
        write_audit_log("REPORT-PRINTED", f"path={path}")
        return True, "Sent to default printer."
    except Exception as e:
        log_exception(f"Failed to print {path}")
        return False, str(e)


def email_pdf_file(path, subject, body, to_addrs_override=None):
    """Emails the generated PDF as an attachment using the same SMTP
    settings configured in the Notifications tab. Returns (ok, message)."""
    cfg = load_notifications_config()
    email_cfg = cfg["email"]
    to_addrs = to_addrs_override or [a.strip() for a in email_cfg.get("to_addrs", "").split(",") if a.strip()]
    if not email_cfg.get("smtp_host") or not to_addrs:
        return False, "SMTP host and recipient(s) must be configured in the Notifications tab first."
    try:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.application import MIMEApplication

        msg = MIMEMultipart()
        msg["Subject"] = f"[R-Apps DICOM] {subject}"
        msg["From"] = email_cfg.get("from_addr") or email_cfg.get("smtp_user") or "noreply@rapps.local"
        msg["To"] = ", ".join(to_addrs)
        msg.attach(MIMEText(body))
        with open(path, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")
        attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
        msg.attach(attachment)

        with smtplib.SMTP(email_cfg["smtp_host"], int(email_cfg.get("smtp_port", 587)), timeout=15) as server:
            if email_cfg.get("use_tls", True):
                server.starttls()
            if email_cfg.get("smtp_user"):
                server.login(email_cfg["smtp_user"], email_cfg.get("smtp_password", ""))
            server.sendmail(msg["From"], to_addrs, msg.as_string())
        write_audit_log("REPORT-EMAILED", f"path={path} to={to_addrs}")
        return True, f"Emailed to {', '.join(to_addrs)}."
    except Exception as e:
        log_exception(f"Failed to email report {path}")
        return False, str(e)


# =========================================================
# BACKUP & RESTORE
# =========================================================
# One ZIP covers every category the spec calls out. Several categories
# map to overlapping files by design (e.g. "Configuration" and
# "Application Preferences" both include bandwidth/notifications
# settings) -- BACKUP_MANIFEST lists them per-category for the history/
# UI display, and _collect_backup_arcnames() dedupes before writing so
# nothing is zipped twice.

def _reports_folder_paths():
    """Every per-patient Reports/ folder under OUTPUT_DIR -- these live
    inside the received-DICOM tree, not as a single top-level file, so
    they need their own discovery step rather than a fixed path list."""
    paths = []
    if os.path.isdir(OUTPUT_DIR):
        for pid in os.listdir(OUTPUT_DIR):
            reports_dir = os.path.join(OUTPUT_DIR, pid, "Reports")
            if os.path.isdir(reports_dir):
                for fname in os.listdir(reports_dir):
                    fpath = os.path.join(reports_dir, fname)
                    if os.path.isfile(fpath):
                        paths.append(fpath)
    return paths


def _log_archive_paths():
    paths = []
    if os.path.isdir(LOG_ARCHIVE_DIR):
        for fname in os.listdir(LOG_ARCHIVE_DIR):
            fpath = os.path.join(LOG_ARCHIVE_DIR, fname)
            if os.path.isfile(fpath):
                paths.append(fpath)
    return paths


def _raw_backup_manifest():
    """The full set of possible manifest paths, regardless of whether
    they currently exist on disk. Used as the restore allow-list --
    get_backup_manifest() (existence-filtered) would wrongly exclude
    exactly the files restore is meant to bring back (ones that are
    currently MISSING)."""
    return {
        "Configuration": [BANDWIDTH_CONFIG_FILE, NOTIFICATIONS_CONFIG, TLS_CONFIG_FILE,
                          LOG_RETENTION_CONFIG_FILE, ADMIN_AUTH_FILE],
        "Encryption Keys": [KEY_FILE],
        "Routing Rules": [ROUTING_RULES_FILE],
        "Destination Profiles": [DESTINATIONS_CONFIG, PUSH_CONFIG, RECEIVER_CONFIG],
        "TLS Configuration": [TLS_CONFIG_FILE],
        "CSV Database": [CSV_FILE],
        "Logs": [RECEIVER_LOG, PUSH_LOG, APP_LOG, RECEIVER_LOG_JSONL, PUSH_LOG_JSONL, APP_LOG_JSONL],
        "Audit Logs": [AUDIT_LOG, AUDIT_LOG_JSONL],
        "Application Preferences": [BANDWIDTH_CONFIG_FILE, NOTIFICATIONS_CONFIG, LOG_RETENTION_CONFIG_FILE],
    }


def get_backup_manifest():
    """Returns {category_label: [relative_paths...]} -- only paths that
    actually currently exist, so an empty/unused category (e.g. no TLS
    config yet) just shows as empty rather than erroring. Reports and Log
    Archives are discovered dynamically since they're per-patient/per-
    rotation folders, not fixed top-level files."""
    manifest = dict(_raw_backup_manifest())
    manifest["Reports"] = _reports_folder_paths()
    manifest["Logs"] = manifest["Logs"] + _log_archive_paths()
    return {cat: [p for p in paths if os.path.isfile(p)] for cat, paths in manifest.items()}


def run_backup_job(dest_path, triggered_by="manual", progress_cb=None):
    """Writes the full backup ZIP, records it in backup_history.json, and
    runs automatic validation (re-opening the ZIP and confirming every
    manifest entry that should be there actually made it in, readable).
    Always call from a background thread."""
    manifest = get_backup_manifest()
    added_arcnames = set()
    all_paths = []
    for category, paths in manifest.items():
        for p in paths:
            if p not in added_arcnames:
                added_arcnames.add(p)
                all_paths.append(p)

    total = len(all_paths) + 1
    with zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED) as zf:
        manifest_payload = {
            "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "categories": {cat: paths for cat, paths in manifest.items()},
        }
        zf.writestr("BACKUP_MANIFEST.json", json.dumps(manifest_payload, indent=2))
        for i, p in enumerate(all_paths):
            try:
                zf.write(p, arcname=p)
            except Exception:
                log_exception(f"Failed to add {p} to backup")
            if progress_cb:
                progress_cb(i + 1, total, f"Backing up {p}")

    ok, message = validate_backup_zip(dest_path, expected_paths=all_paths)

    size_bytes = os.path.getsize(dest_path) if os.path.exists(dest_path) else 0
    record = {
        "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
        "path": dest_path,
        "size_bytes": size_bytes,
        "triggered_by": triggered_by,
        "file_count": len(all_paths),
        "validation_ok": ok,
        "validation_message": message,
    }
    _append_backup_history(record)
    write_audit_log("BACKUP-COMPLETED" if ok else "BACKUP-VALIDATION-FAILED",
                    f"path={dest_path} files={len(all_paths)} triggered_by={triggered_by} validation={message}")
    notify_event("backup_completed" if ok else "backup_failed",
                "Backup Completed" if ok else "Backup Failed",
                f"{os.path.basename(dest_path)} ({len(all_paths)} files) -- {message}")
    return record


def validate_backup_zip(dest_path, expected_paths=None):
    """Automatic Validation: confirms the ZIP itself isn't corrupt
    (testzip) and that every file we intended to include actually made
    it in. Returns (ok, message)."""
    try:
        with zipfile.ZipFile(dest_path, "r") as zf:
            bad_file = zf.testzip()
            if bad_file:
                return False, f"Corrupt member in archive: {bad_file}"
            names = set(zf.namelist())
            if expected_paths:
                missing = [p for p in expected_paths if p not in names]
                if missing:
                    return False, f"{len(missing)} expected file(s) missing from archive (e.g. {missing[0]})"
            return True, f"Validated OK -- {len(names)} entries."
    except Exception as e:
        log_exception(f"Backup validation failed for {dest_path}")
        return False, f"Validation error: {e}"


def _append_backup_history(record):
    try:
        history = load_backup_history()
        history.append(record)
        atomic_write(BACKUP_HISTORY_FILE, json.dumps(history, indent=2))
    except Exception:
        log_exception("Failed to append backup history")


def load_backup_history():
    try:
        if os.path.exists(BACKUP_HISTORY_FILE):
            with open(BACKUP_HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        log_exception("Failed to load backup_history.json")
    return []


def restore_backup(backup_path, progress_cb=None):
    """Restores every file in the backup ZIP to its original relative
    path. For safety: (1) validates the ZIP first and refuses to restore
    a corrupt one, (2) takes an automatic 'pre-restore safety backup' of
    the CURRENT state first so a bad restore is itself recoverable, (3)
    only ever writes to paths that appeared in this app's own manifest
    categories (never arbitrary paths from the zip, even though the zip
    is one this app created -- defense in depth against a
    tampered/corrupted archive)."""
    ok, message = validate_backup_zip(backup_path)
    if not ok:
        raise ValueError(f"Refusing to restore -- backup failed validation: {message}")

    # Safety net: back up current state before overwriting anything.
    os.makedirs(BACKUP_DIR, exist_ok=True)
    safety_path = os.path.join(BACKUP_DIR, f"pre_restore_safety_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")
    try:
        run_backup_job(safety_path, triggered_by="pre-restore-safety")
    except Exception:
        log_exception("Failed to create pre-restore safety backup -- continuing with restore anyway")

    allowed_paths = set()
    for paths in _raw_backup_manifest().values():
        allowed_paths.update(paths)

    restored = []
    with zipfile.ZipFile(backup_path, "r") as zf:
        names = [n for n in zf.namelist() if n != "BACKUP_MANIFEST.json"]
        total = len(names)
        for i, name in enumerate(names):
            # Only restore paths that are plausibly ours: no traversal
            # tricks, and it must be a file this app's own manifest could
            # have produced (either a known top-level config file, or
            # under received_dicoms/.../Reports/ or log_archives/).
            normalized = name.replace("\\", "/")
            is_safe_relative = not normalized.startswith("/") and ".." not in normalized.split("/")
            is_known_category_path = (
                normalized in allowed_paths
                or normalized.startswith(OUTPUT_DIR + "/")
                or normalized.startswith(LOG_ARCHIVE_DIR + "/")
            )
            if is_safe_relative and is_known_category_path:
                target_dir = os.path.dirname(normalized)
                if target_dir:
                    os.makedirs(target_dir, exist_ok=True)
                with zf.open(name) as src, open(normalized, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                restored.append(normalized)
            if progress_cb:
                progress_cb(i + 1, total, f"Restoring {name}")

    # Any in-memory config caches must be invalidated so the restored
    # files are actually picked up instead of stale cached values.
    _bandwidth_config_cache["value"] = None
    _notifications_config_cache["value"] = None
    _checkpoint_cache["value"] = None
    _offline_queue_cache["value"] = None

    write_audit_log("RESTORE-COMPLETED", f"path={backup_path} files_restored={len(restored)} safety_backup={safety_path}")
    notify_event("backup_completed", "Restore Completed",
                f"{len(restored)} file(s) restored from {os.path.basename(backup_path)}. A restart is recommended.")
    return restored, safety_path


# ---------------------------------------------------------
# Scheduled backups
# ---------------------------------------------------------

DEFAULT_BACKUP_SCHEDULE = {"enabled": False, "frequency": "Daily", "hour": 2}  # 2 AM local time by default
_backup_schedule_cache = {"value": None}
_last_scheduled_backup_date = {"value": None}
_backup_scheduler_started = {"value": False}


EXPORT_TEMPLATE_CONFIG_FILE = "export_template_config.json"  # not secret -- remembers last-used Export tab checkboxes/scope, never the password itself
DEFAULT_EXPORT_TEMPLATE = {
    "scope": "Entire Patient",
    "include_reports": False,
    "include_logs": False,
    "include_metadata": False,
    "include_dicomdir": False,
    "had_password": False,  # informational only -- the password ITSELF is never persisted
}
_export_template_cache = {"value": None}


def load_export_template():
    if _export_template_cache["value"] is not None:
        return _export_template_cache["value"]
    cfg = dict(DEFAULT_EXPORT_TEMPLATE)
    try:
        if os.path.exists(EXPORT_TEMPLATE_CONFIG_FILE):
            with open(EXPORT_TEMPLATE_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
    except Exception:
        log_exception("Failed to load export_template_config.json")
    _export_template_cache["value"] = cfg
    return cfg


def save_export_template(cfg):
    try:
        atomic_write(EXPORT_TEMPLATE_CONFIG_FILE, json.dumps(cfg, indent=2))
    except Exception:
        log_exception("Failed to save export_template_config.json")
    _export_template_cache["value"] = cfg


def load_backup_schedule():
    if _backup_schedule_cache["value"] is not None:
        return _backup_schedule_cache["value"]
    cfg = dict(DEFAULT_BACKUP_SCHEDULE)
    try:
        if os.path.exists(BACKUP_SCHEDULE_CONFIG_FILE):
            with open(BACKUP_SCHEDULE_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
    except Exception:
        log_exception("Failed to load backup_schedule_config.json")
    _backup_schedule_cache["value"] = cfg
    return cfg


def save_backup_schedule(cfg):
    atomic_write(BACKUP_SCHEDULE_CONFIG_FILE, json.dumps(cfg, indent=2))
    _backup_schedule_cache["value"] = cfg
    write_audit_log("BACKUP-SCHEDULE-CHANGED", f"enabled={cfg['enabled']} frequency={cfg['frequency']} hour={cfg['hour']}")


def _is_scheduled_backup_due(cfg, now):
    if not cfg.get("enabled"):
        return False
    if now.hour != cfg.get("hour", 2):
        return False
    today_key = now.date().isoformat()
    if cfg.get("frequency") == "Weekly":
        today_key = f"{now.isocalendar()[0]}-W{now.isocalendar()[1]}"
    return _last_scheduled_backup_date["value"] != today_key


def _compute_next_scheduled_backup_time(cfg, now=None):
    """Estimates when the next scheduled backup will fire, using the same
    hour-of-day + already-ran-this-period bookkeeping _is_scheduled_backup_due()
    checks, so the two can't drift out of sync. The scheduler itself
    doesn't pin scheduled backups to a specific weekday (it just compares
    ISO-week keys), so the Weekly case's estimate is the next Monday's
    slot -- the nearest deterministic guess consistent with that."""
    if not cfg.get("enabled"):
        return None
    now = now or datetime.datetime.now()
    hour = cfg.get("hour", 2)
    today_key = now.date().isoformat()
    this_week_key = f"{now.isocalendar()[0]}-W{now.isocalendar()[1]}"
    period_key = this_week_key if cfg.get("frequency") == "Weekly" else today_key
    already_ran_this_period = _last_scheduled_backup_date["value"] == period_key
    candidate = now.replace(hour=hour, minute=0, second=0, microsecond=0)
    if not already_ran_this_period and candidate > now:
        return candidate
    if cfg.get("frequency") == "Weekly":
        days_until_monday = (7 - now.weekday()) % 7 or 7
        candidate = (now + datetime.timedelta(days=days_until_monday)).replace(
            hour=hour, minute=0, second=0, microsecond=0)
    else:
        candidate = (now + datetime.timedelta(days=1)).replace(hour=hour, minute=0, second=0, microsecond=0)
    return candidate


def refresh_next_backup_indicator():
    if not admin_tabs_active["value"] or "backup_next_run_lbl" not in globals():
        return
    cfg = load_backup_schedule()
    next_time = _compute_next_scheduled_backup_time(cfg)
    if next_time is None:
        backup_next_run_lbl.configure(text="Backups disabled")
    else:
        backup_next_run_lbl.configure(text=f"Next backup: {next_time.strftime('%Y-%m-%d %H:%M')}")


def backup_scheduler_loop():
    """Background daemon: checks once a minute whether a scheduled
    backup is due (matches the configured hour-of-day, and hasn't
    already run today/this-week)."""
    while not app_shutdown_event.is_set():
        try:
            cfg = load_backup_schedule()
            now = datetime.datetime.now()
            if _is_scheduled_backup_due(cfg, now):
                os.makedirs(BACKUP_DIR, exist_ok=True)
                dest_path = os.path.join(BACKUP_DIR, f"scheduled_backup_{now.strftime('%Y%m%d_%H%M%S')}.zip")
                try:
                    run_backup_job(dest_path, triggered_by="scheduled")
                except Exception:
                    log_exception("Scheduled backup failed")
                today_key = now.date().isoformat()
                if cfg.get("frequency") == "Weekly":
                    today_key = f"{now.isocalendar()[0]}-W{now.isocalendar()[1]}"
                _last_scheduled_backup_date["value"] = today_key
        except Exception:
            log_exception("Backup scheduler loop iteration failed")
        app_shutdown_event.wait(60)


def start_backup_scheduler_thread():
    if _backup_scheduler_started["value"]:
        return
    _backup_scheduler_started["value"] = True
    threading.Thread(target=backup_scheduler_loop, daemon=True).start()


# =========================================================
# SCHEDULED REPORT EMAIL  (8.2)
# =========================================================
# Same lightweight "check once a minute, fire once per day/week at the
# configured hour" pattern as backup_scheduler_loop()/_is_scheduled_backup_due()
# above, applied to periodically generating + emailing a report PDF instead
# of a backup ZIP. Uses the exact same SMTP settings already loaded via
# load_notifications_config() that the manual "Email Report" button uses.

REPORT_EMAIL_SCHEDULE_CONFIG_FILE = "report_email_schedule_config.json"  # not secret -- just cadence settings
DEFAULT_REPORT_EMAIL_SCHEDULE = {"enabled": False, "frequency": "Daily", "hour": 6}
_report_email_schedule_cache = {"value": None}
_last_scheduled_report_email_date = {"value": None}
_report_email_scheduler_started = {"value": False}


def load_report_email_schedule():
    if _report_email_schedule_cache["value"] is not None:
        return _report_email_schedule_cache["value"]
    cfg = dict(DEFAULT_REPORT_EMAIL_SCHEDULE)
    try:
        if os.path.exists(REPORT_EMAIL_SCHEDULE_CONFIG_FILE):
            with open(REPORT_EMAIL_SCHEDULE_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
    except Exception:
        log_exception("Failed to load report_email_schedule_config.json")
    _report_email_schedule_cache["value"] = cfg
    return cfg


def save_report_email_schedule(cfg):
    atomic_write(REPORT_EMAIL_SCHEDULE_CONFIG_FILE, json.dumps(cfg, indent=2))
    _report_email_schedule_cache["value"] = cfg
    write_audit_log("REPORT-EMAIL-SCHEDULE-CHANGED", f"enabled={cfg['enabled']} frequency={cfg['frequency']} hour={cfg['hour']}")


def _is_scheduled_report_email_due(cfg, now):
    if not cfg.get("enabled"):
        return False
    if now.hour != cfg.get("hour", 6):
        return False
    today_key = now.date().isoformat()
    if cfg.get("frequency") == "Weekly":
        today_key = f"{now.isocalendar()[0]}-W{now.isocalendar()[1]}"
    return _last_scheduled_report_email_date["value"] != today_key


def _run_scheduled_report_email(cfg, now):
    if cfg.get("frequency") == "Weekly":
        start_dt = (now - datetime.timedelta(days=7)).replace(hour=0, minute=0, second=0, microsecond=0)
        label = "Weekly"
    else:
        start_dt = (now - datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        label = "Daily"
    end_dt = now
    os.makedirs(SCHEDULED_REPORTS_DIR, exist_ok=True)
    dest_path = os.path.join(SCHEDULED_REPORTS_DIR, f"scheduled_report_{label.lower()}_{now.strftime('%Y%m%d_%H%M%S')}.pdf")
    generate_report_pdf(start_dt, end_dt, dest_path, label)
    ok, msg = email_pdf_file(
        dest_path, f"{label} Report ({now.strftime('%Y-%m-%d')})",
        f"Attached is the scheduled {label.lower()} R-Apps DICOM report, covering "
        f"{start_dt.strftime('%Y-%m-%d %H:%M')} to {end_dt.strftime('%Y-%m-%d %H:%M')}.")
    if ok:
        write_audit_log("REPORT-EMAIL-SCHEDULED-SENT", f"label={label} path={dest_path}")
    else:
        log_exception(f"Scheduled report email failed: {msg}")
        notify_event("push_complete", "Scheduled Report Email Failed", f"{label} report could not be emailed: {msg}")


def report_email_scheduler_loop():
    """Background daemon: checks once a minute whether a scheduled report
    email is due (same hour-of-day + not-already-run-today/this-week
    bookkeeping as backup_scheduler_loop())."""
    while not app_shutdown_event.is_set():
        try:
            cfg = load_report_email_schedule()
            now = datetime.datetime.now()
            if _is_scheduled_report_email_due(cfg, now):
                try:
                    _run_scheduled_report_email(cfg, now)
                except Exception:
                    log_exception("Scheduled report email generation failed")
                today_key = now.date().isoformat()
                if cfg.get("frequency") == "Weekly":
                    today_key = f"{now.isocalendar()[0]}-W{now.isocalendar()[1]}"
                _last_scheduled_report_email_date["value"] = today_key
        except Exception:
            log_exception("Report email scheduler loop iteration failed")
        app_shutdown_event.wait(60)


def start_report_email_scheduler_thread():
    if _report_email_scheduler_started["value"]:
        return
    _report_email_scheduler_started["value"] = True
    threading.Thread(target=report_email_scheduler_loop, daemon=True).start()


# =========================================================
# LDAP / ACTIVE DIRECTORY AUTHENTICATION
# =========================================================
# Real LDAP/AD support via ldap3 (bind-based auth + group-membership
# lookup), with automatic user import and configurable group-to-role
# mapping. Local Admin PIN authentication (see verify_admin_pin() above)
# remains fully intact and always available as a fallback -- LDAP is
# purely additive and never required.
#
# Role mapping note: this app's existing UI has two privilege tiers --
# Admin mode (unlocks every admin-only tab) and User mode (Receiver/
# Pusher only). The group-mapping system below resolves an LDAP/AD user
# to one of five named roles (Administrators, Technicians, Radiologists,
# Support, Guest) exactly as configured, and that resolved role is
# recorded against the user (roster, audit trail, session identity).
# Only the "Administrators" role actually unlocks Admin mode today, since
# building five distinct permission tiers throughout the UI is a much
# larger undertaking than the authentication layer itself -- the other
# four roles all land in the existing User mode. The mapping and the
# resolved role are real and fully configurable; it's the enforcement
# granularity that's currently binary, and that's stated plainly here and
# in the LDAP tab's UI rather than implied to be more granular than it is.

LDAP_ROLE_NAMES = ["Administrators", "Technicians", "Radiologists", "Support", "Guest"]

DEFAULT_LDAP_CONFIG = {
    "enabled": False,
    "server_uri": "",             # e.g. ldap://dc1.company.local:389 or ldaps://dc1.company.local:636
    "use_ssl": False,
    "domain": "",                 # e.g. company.local -- used to build username@domain for AD binds
    "bind_dn": "",                # service account used for searches, e.g. svc-rapps@company.local
    "bind_password": "",
    "user_search_base": "",       # e.g. ou=Users,dc=company,dc=local
    "user_search_filter": "(sAMAccountName={username})",
    "user_bind_dn_template": "", # generic/OpenLDAP-style: e.g. "cn={username},ou=Users,dc=company,dc=local".
                                 # Leave blank for AD (uses username@domain instead -- see `domain` above).
    "group_search_base": "",
    "display_name_attr": "displayName",
    "email_attr": "mail",
    "member_of_attr": "memberOf",
    # Maps an LDAP/AD group CN (not full DN -- just the CN, e.g. "RAPPS-Admins")
    # to one of LDAP_ROLE_NAMES.
    "group_mappings": {},
    "default_role_if_unmapped": "Guest",
}

_ldap_config_cache = {"value": None}


def load_ldap_config():
    if _ldap_config_cache["value"] is not None:
        return _ldap_config_cache["value"]
    cfg = json.loads(json.dumps(DEFAULT_LDAP_CONFIG))
    data = decrypt_and_load(LDAP_CONFIG_FILE)
    if data:
        try:
            loaded = json.loads(data)
            cfg.update(loaded)
        except Exception:
            log_exception("Failed to parse ldap_config.enc, using defaults")
    _ldap_config_cache["value"] = cfg
    return cfg


def save_ldap_config(cfg):
    encrypt_and_save(LDAP_CONFIG_FILE, json.dumps(cfg))
    _ldap_config_cache["value"] = cfg
    write_audit_log("LDAP-CONFIG-CHANGED", f"enabled={cfg.get('enabled')} server={cfg.get('server_uri')}")


def load_ldap_users():
    try:
        if os.path.exists(LDAP_USERS_FILE):
            with open(LDAP_USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        log_exception("Failed to load ldap_users.json")
    return {}


def _save_ldap_users(users):
    try:
        atomic_write(LDAP_USERS_FILE, json.dumps(users, indent=2))
    except Exception:
        log_exception("Failed to save ldap_users.json")


def _resolve_role_from_groups(cfg, group_dns):
    """group_dns is whatever the directory returned for memberOf (a list
    of full DNs, e.g. 'CN=RAPPS-Admins,OU=Groups,DC=company,DC=local').
    Matches on CN against the configured group_mappings, first match
    wins (in the order LDAP_ROLE_NAMES lists them, admin-first, so a
    user in multiple mapped groups gets the most privileged one)."""
    cns = set()
    for dn in group_dns or []:
        for part in str(dn).split(","):
            part = part.strip()
            if part.upper().startswith("CN="):
                cns.add(part[3:])
    mappings = cfg.get("group_mappings", {})
    matched_roles = {mappings[cn] for cn in cns if cn in mappings}
    for role in LDAP_ROLE_NAMES:
        if role in matched_roles:
            return role
    return cfg.get("default_role_if_unmapped", "Guest")


def _upsert_ldap_user_record(username, display_name, email, role):
    users = load_ldap_users()
    prior = users.get(username, {})
    users[username] = {
        "username": username, "display_name": display_name, "email": email,
        "role": role, "last_login": datetime.datetime.now().isoformat(timespec="seconds"),
        "revoked": prior.get("revoked", False),  # a fresh successful directory bind shouldn't silently clear a revoke
    }
    _save_ldap_users(users)


def _set_ldap_user_revoked(username, revoked):
    """7.2 -- flips a local active/revoked flag on the stored LDAP user
    record, independent of directory sync state. Checked by
    ldap_authenticate() so a revoked account is blocked locally even
    though it would still bind successfully against the directory itself."""
    users = load_ldap_users()
    if username not in users:
        return False
    users[username]["revoked"] = bool(revoked)
    _save_ldap_users(users)
    write_audit_log("LDAP-ACCESS-REVOKED" if revoked else "LDAP-ACCESS-RESTORED", f"username={username}")
    return True


def ldap_authenticate(username, password):
    """Attempts a real LDAP/AD bind as `username`/`password`, then (using
    the service account) looks up group membership to resolve a role.
    Returns (ok, message, user_info_or_None). Automatic User Import: on
    success, the user's basic directory info + resolved role is recorded
    in ldap_users.json -- this IS the automatic import (each successful
    login imports/updates that user), separate from the bulk
    ldap_import_all_users() below for importing the whole directory
    up front."""
    if not LDAP3_AVAILABLE:
        return False, "LDAP support requires the 'ldap3' package. Install with: pip install ldap3", None

    cfg = load_ldap_config()
    if not cfg.get("enabled"):
        return False, "LDAP/AD authentication is not enabled.", None
    if not cfg.get("server_uri"):
        return False, "LDAP server is not configured.", None

    bind_username = username
    if cfg.get("user_bind_dn_template"):
        bind_username = cfg["user_bind_dn_template"].format(username=username)
    elif cfg.get("domain") and "@" not in username and "\\" not in username:
        bind_username = f"{username}@{cfg['domain']}"

    try:
        server = Server(cfg["server_uri"], use_ssl=cfg.get("use_ssl", False), get_info=None)
        # Step 1: the user's own credentials must bind successfully --
        # this IS the authentication check, delegated entirely to the
        # directory rather than reimplementing password verification.
        user_conn = Connection(server, user=bind_username, password=password, authentication=SIMPLE)
        if not user_conn.bind():
            return False, "Invalid domain username or password.", None
        user_conn.unbind()

        if load_ldap_users().get(username, {}).get("revoked"):
            write_audit_log("LDAP-LOGIN-BLOCKED", f"username={username} reason=revoked", username=username)
            return False, "Access for this account has been revoked. Contact an administrator.", None

        # Step 2: use the service account (if configured) to look up the
        # user's directory attributes + group membership for role
        # resolution. Falls back to the user's own credentials if no
        # separate service account is configured.
        search_conn = Connection(
            server,
            user=cfg["bind_dn"] or bind_username,
            password=cfg["bind_password"] or password,
            authentication=SIMPLE,
        )
        if not search_conn.bind():
            return False, "Domain login succeeded, but the search/service account could not bind to look up group membership.", None

        search_filter = cfg["user_search_filter"].format(username=username)
        search_conn.search(
            search_base=cfg.get("user_search_base", ""),
            search_filter=search_filter,
            search_scope=SUBTREE,
            attributes=[cfg["display_name_attr"], cfg["email_attr"], cfg["member_of_attr"]],
        )
        display_name, email, group_dns = username, "", []
        if search_conn.entries:
            entry = search_conn.entries[0]
            display_name = str(getattr(entry, cfg["display_name_attr"], username) or username)
            email = str(getattr(entry, cfg["email_attr"], "") or "")
            member_of = getattr(entry, cfg["member_of_attr"], [])
            group_dns = list(member_of) if member_of else []
        search_conn.unbind()

        role = _resolve_role_from_groups(cfg, group_dns)
        _upsert_ldap_user_record(username, display_name, email, role)
        write_audit_log("LDAP-LOGIN-SUCCESS", f"username={username} role={role}", username=username)
        return True, f"Welcome, {display_name} ({role}).", {
            "username": username, "display_name": display_name, "email": email, "role": role,
        }
    except Exception as e:
        log_exception(f"LDAP authentication failed for {username}")
        write_audit_log("LDAP-LOGIN-FAILED", f"username={username} error={e}", username=username)
        notify_event("authentication_failures", "Authentication Failure", f"LDAP login failed for {username}: {e}")
        return False, f"LDAP error: {e}", None


def ldap_import_all_users(progress_cb=None):
    """Bulk directory import: searches the entire configured
    user_search_base for every account matching a generic filter,
    resolves each one's role from group membership, and records them all
    in ldap_users.json -- without requiring each person to log in first.
    Returns (count_imported, error_or_None)."""
    if not LDAP3_AVAILABLE:
        return 0, "LDAP support requires the 'ldap3' package. Install with: pip install ldap3"
    cfg = load_ldap_config()
    if not cfg.get("server_uri") or not cfg.get("bind_dn"):
        return 0, "Server URI and a service/bind account must be configured first."

    try:
        server = Server(cfg["server_uri"], use_ssl=cfg.get("use_ssl", False), get_info=None)
        conn = Connection(server, user=cfg["bind_dn"], password=cfg["bind_password"], authentication=SIMPLE)
        if not conn.bind():
            return 0, "Service account could not bind -- check bind DN / password."

        # Generic "every user account" filter for AD; works for most
        # standard schemas. Directories with a different objectClass
        # convention can adjust user_search_filter's template instead,
        # this bulk import intentionally uses a broader filter than the
        # per-login lookup since there's no specific username to match.
        bulk_filter = "(&(objectClass=user)(objectCategory=person))"
        conn.search(
            search_base=cfg.get("user_search_base", ""),
            search_filter=bulk_filter,
            search_scope=SUBTREE,
            attributes=["sAMAccountName", cfg["display_name_attr"], cfg["email_attr"], cfg["member_of_attr"]],
        )
        count = 0
        total = len(conn.entries)
        for i, entry in enumerate(conn.entries):
            username = str(getattr(entry, "sAMAccountName", "") or "")
            if not username:
                continue
            display_name = str(getattr(entry, cfg["display_name_attr"], username) or username)
            email = str(getattr(entry, cfg["email_attr"], "") or "")
            member_of = getattr(entry, cfg["member_of_attr"], [])
            group_dns = list(member_of) if member_of else []
            role = _resolve_role_from_groups(cfg, group_dns)
            _upsert_ldap_user_record(username, display_name, email, role)
            count += 1
            if progress_cb:
                progress_cb(i + 1, total, f"Imported {username}")
        conn.unbind()
        write_audit_log("LDAP-BULK-IMPORT", f"count={count}")
        return count, None
    except Exception as e:
        log_exception("LDAP bulk user import failed")
        return 0, str(e)



# Dark Mode is this application's single, permanent theme -- there is no
# Light theme and no user-facing switch. The palette below is the one
# and only source of truth for every color in the GUI; the High Contrast
# Mode accessibility setting adjusts it (see _compute_theme_palette()),
# but it never changes away from dark.

DARK_THEME_PALETTE = {
    "bg": "#0e1015", "surface": "#181b21", "heading_bg": "#232733",
    "accent": "#2f8eff", "accent_hover": "#1f6fd6", "text": "#e8eaed",
    "text_muted": "#8b93a7", "neutral_btn": "#3a3f4c", "neutral_btn_hover": "#4b5263",
    "danger": "#f04747", "danger_hover": "#c0392b", "success": "#2ecc71", "success_hover": "#27ae60",
    "warning": "#e67e22", "warning_hover": "#c9701c",
    "odd_row": "#1d212a", "stale": "#ff5555", "search_highlight": "#3a3510",
    "segmented_hover": "#2c3140",
}


def _compute_theme_palette():
    """Returns the active color palette: the app's one permanent Dark
    Mode palette, widened for contrast when the High Contrast Mode
    accessibility setting (Settings > Accessibility) is enabled. This is
    evaluated once at startup, before any widgets are built, so every
    widget in the app is constructed with the correct colors from the
    start -- there's nothing to switch later."""
    palette = dict(DARK_THEME_PALETTE)
    if APP_SETTINGS.get("high_contrast_mode"):
        palette["text_muted"] = "#c9cfd9"       # much closer to full-white than the default muted gray
        palette["heading_bg"] = "#2c313d"        # lighter surface -> stronger separation from bg
        palette["neutral_btn"] = "#454c5c"
        palette["neutral_btn_hover"] = "#5a6478"
        palette["accent"] = "#5aa3ff"            # brighter accent for stronger selection/focus contrast
        palette["segmented_hover"] = "#3a4152"
    return palette




# =========================================================
# ANONYMIZATION (DE-IDENTIFICATION)
# =========================================================
# A practical, non-exhaustive de-identification profile loosely aligned
# with DICOM PS3.15 Basic Application Level Confidentiality Profile.
# This is NOT a certified de-identification tool -- it is meant for
# teaching / research workflows, not regulatory compliance. Patient
# identity is replaced with a stable pseudonym derived from the original
# PatientID so the same patient maps consistently within one push job.

ANONYMIZE_TAGS_BLANK = [
    "PatientBirthDate", "PatientBirthTime", "PatientAddress",
    "PatientTelephoneNumbers", "OtherPatientIDs", "OtherPatientNames",
    "PatientMotherBirthName", "MilitaryRank", "BranchOfService",
    "MedicalRecordLocator", "ReferringPhysicianAddress",
    "ReferringPhysicianTelephoneNumbers", "InstitutionAddress",
    "PersonAddress", "PersonTelephoneNumbers",
]

ANONYMIZE_TAGS_REMOVE_IF_PRESENT = [
    "InstitutionalDepartmentName",
]


def _pseudonym_for(pid):
    import hashlib
    digest = hashlib.sha256(pid.encode("utf-8")).hexdigest()[:10].upper()
    return f"ANON-{digest}"


def anonymize_dataset(ds, original_pid):
    """Mutate a pydicom Dataset in place to remove direct identifiers.
    Safe to call even if some tags are absent (best-effort -- logs any
    failure rather than silently leaving an identifier un-anonymized;
    see §3.3). Still proceeds best-effort on a failure (matching prior
    behavior) rather than aborting the push outright -- if you want a
    failed anonymization step to hard-block the push instead, that's a
    deliberate behavior change worth its own sign-off rather than a
    silent side-effect of this cleanup."""
    pseudonym = _pseudonym_for(original_pid)
    try:
        ds.PatientID = pseudonym
        ds.PatientName = pseudonym
    except Exception:
        # This is the core de-identification guarantee -- log loudly.
        log_exception(f"anonymize_dataset: FAILED to replace PatientID/PatientName "
                      f"for original pid={original_pid!r} -- original identifier may "
                      f"still be present on the dataset about to be sent")

    for tag in ANONYMIZE_TAGS_BLANK:
        if hasattr(ds, tag):
            try:
                setattr(ds, tag, "")
            except Exception:
                log_exception(f"anonymize_dataset: failed to blank tag {tag!r} (pid={original_pid!r})")

    for tag in ANONYMIZE_TAGS_REMOVE_IF_PRESENT:
        if hasattr(ds, tag):
            try:
                delattr(ds, tag)
            except Exception:
                log_exception(f"anonymize_dataset: failed to remove tag {tag!r} (pid={original_pid!r})")

    # Strip private tags and curves/overlays which can carry burned-in PHI
    try:
        ds.remove_private_tags()
    except Exception:
        log_exception(f"anonymize_dataset: failed to remove private tags (pid={original_pid!r})")

    return ds

# =========================================================
# CSV PERSISTENCE
# =========================================================

CSV_FIELDS = [
    "Patient ID", "Patient Name", "Institution", "Study UID",
    "Modality", "Receive Time", "Image Count", "Document Count", "Source", "Status",
    "Sent Time", "Pushed To", "Last Error", "Retry Count",
    "Report Exists", "Report Created Date", "Report Last Opened",
]


def autosave_csv():
    with data_lock:
        try:
            with atomic_open_for_write(CSV_FILE, mode="w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(CSV_FIELDS)
                for pid, d in patient_data.items():
                    writer.writerow([
                        pid,
                        d.get("patient_name", ""),
                        d.get("institution", ""),
                        d.get("study_uid", ""),
                        d.get("modality", ""),
                        d.get("time", ""),
                        d.get("count", 0),
                        d.get("doc_count", 0),
                        d.get("source", STATUS_RECEIVED),
                        d.get("status", STATUS_PENDING),
                        d.get("sent_time", ""),
                        d.get("push_target", ""),
                        d.get("last_error", ""),
                        d.get("retry_count", 0),
                        d.get("report_exists", False),
                        d.get("report_created_date", ""),
                        d.get("report_last_opened", ""),
                    ])
        except Exception:
            log_exception("Failed to autosave worklist CSV")


def load_csv():
    if not os.path.exists(CSV_FILE):
        return
    try:
        with open(CSV_FILE, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            header = reader.fieldnames or []
            missing = [col for col in CSV_FIELDS if col not in header]
            if missing:
                # §3.8 fix: previously a header mismatch raised a KeyError
                # on the first data row, was swallowed by the broad
                # except below, and silently left patient_data empty (or
                # partially populated) with no indication anything had
                # gone wrong. Refuse to guess at a mismatched worklist --
                # surface it loudly instead, and leave the CSV file on
                # disk untouched so it can be recovered/repaired by hand.
                err = (f"Worklist CSV header does not match the expected "
                       f"format (missing column(s): {', '.join(missing)}). "
                       f"Refusing to load a mismatched worklist -- "
                       f"{CSV_FILE} was left untouched. Restore it from a "
                       f"backup or fix the header by hand.")
                app_logger.error(err)
                try:
                    write_audit_log("WORKLIST-LOAD-FAILED", err)
                except Exception:
                    pass
                try:
                    ui_event_queue.put(("toast", ("Worklist Failed to Load", err)))
                except Exception:
                    pass
                return

            skipped = 0
            for row_num, row in enumerate(reader, start=2):  # header is line 1
                try:
                    pid = row["Patient ID"]
                    if not pid:
                        continue
                    patient_data[pid] = {
                        "patient_name": format_dicom_person_name(row.get("Patient Name", "")),
                        "institution": row.get("Institution", ""),
                        "study_uid": row.get("Study UID", ""),
                        "modality": row.get("Modality", ""),
                        "time": row.get("Receive Time", ""),
                        "count": int(row.get("Image Count", 0) or 0),
                        "doc_count": int(row.get("Document Count", 0) or 0),
                        "source": row.get("Source", STATUS_RECEIVED),
                        "status": row.get("Status", STATUS_PENDING),
                        "sent_time": row.get("Sent Time", ""),
                        "push_target": row.get("Pushed To", ""),
                        "last_error": row.get("Last Error", ""),
                        "retry_count": int(row.get("Retry Count", 0) or 0),
                        "report_exists": str(row.get("Report Exists", "")).strip() in ("True", "1", "yes"),
                        "report_created_date": row.get("Report Created Date", ""),
                        "report_last_opened": row.get("Report Last Opened", ""),
                    }
                except Exception:
                    # One malformed row (bad int, etc.) must not cost every
                    # row after it -- log which row/pid and keep going.
                    skipped += 1
                    log_exception(f"Failed to load worklist CSV row {row_num} "
                                  f"(Patient ID={row.get('Patient ID', '?')!r}) -- skipping this row")
            if skipped:
                app_logger.warning("Worklist CSV load finished with %d row(s) skipped due to errors.", skipped)
    except Exception:
        log_exception("Failed to load worklist CSV")
    _bump_data_version()

# =========================================================
# LOGGING (receiver / push error logs + audit trail)
# =========================================================
#
# Enterprise diagnostic logging design:
#   * The original plain-text .log files (RECEIVER_LOG / PUSH_LOG /
#     AUDIT_LOG) are preserved byte-for-byte in format and behavior --
#     anything that already tails/parses them keeps working exactly as
#     before.
#   * Alongside them, every event is ALSO written as one JSON object per
#     line to a matching *_events.jsonl file. This is where the expanded
#     enterprise field set lives (StudyUID, SeriesUID, SOP Instance UID,
#     Calling/Called AE, source IP/port, transfer syntax, file size,
#     receive/push duration, retry counts, transfer speed, TLS state,
#     full exception details, username/computer name for audit events,
#     etc.) without disturbing the existing plain-text format.
#   * Logs are immutable: there is no UI action anywhere in the app that
#     truncates or deletes a log file. rotate_logs_if_needed() moves a
#     log (both the .log and its .jsonl twin) into LOG_ARCHIVE_DIR as a
#     timestamped .zip once it is older than the configured retention
#     window or exceeds LOG_ROTATION_MAX_BYTES, then starts a fresh log.
#     Archives themselves are never auto-deleted.

_log_write_lock = threading.RLock()


def _computer_name():
    try:
        return platform.node() or "UNKNOWN-HOST"
    except Exception:
        return "UNKNOWN-HOST"


def _os_username():
    try:
        return os.environ.get("USERNAME") or os.environ.get("USER") or "unknown"
    except Exception:
        return "unknown"


def _write_jsonl(path, record: dict):
    """Append one JSON object as a line. Never raises -- logging must
    never be able to crash the calling operation."""
    try:
        record.setdefault("timestamp", datetime.datetime.now().isoformat(timespec="seconds"))
        with _log_write_lock:
            with open(path, "a", encoding="utf-8") as f:
                f.write(json.dumps(record, default=str) + "\n")
    except Exception:
        log_exception(f"Failed to write structured log entry to {path}")


def _log(path, pid, pname, extra, error):
    try:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with _log_write_lock:
            with open(path, "a", encoding="utf-8") as f:
                f.write(f"[{now}] PATIENT_ID={pid} | PATIENT_NAME={pname} | {extra} ERROR={error}\n")
    except Exception:
        log_exception(f"Failed to write log entry to {path}")


def write_receiver_log(patient_id, patient_name, institution, error, **details):
    """Accepts plain values (not a fake dataset object) so callers don't
    need the previous type()-hack workaround.

    `**details` is new and fully optional/backward-compatible -- existing
    call sites that only pass (patient_id, patient_name, institution, error)
    keep working unchanged. When callers *do* supply extras (study_uid,
    series_uid, sop_instance_uid, calling_ae, called_ae, source_ip,
    source_port, transfer_syntax, sop_class, modality, file_size,
    duration_sec, save_location, result, error_code, stack_trace, ...)
    they're captured in the structured .jsonl log for enterprise
    troubleshooting."""
    _log(RECEIVER_LOG, patient_id, patient_name, f"INSTITUTION={institution} | ", error)
    record = {
        "event": "receive",
        "patient_id": patient_id,
        "patient_name": patient_name,
        "institution": institution,
        "error_message": error or None,
        "result": details.pop("result", "FAILURE" if error else "SUCCESS"),
    }
    record.update(details)
    _write_jsonl(RECEIVER_LOG_JSONL, record)
    rotate_logs_if_needed()


def write_push_log(pid, pname, error, **details):
    """See write_receiver_log docstring -- **details is additive/optional.
    Typical extras: destination_name, destination_ae, destination_ip,
    destination_port, calling_ae, called_ae, study_uid, series_uid,
    images_attempted, images_sent, images_failed, retry_count,
    transfer_speed_mbps, transfer_time_sec, compression_used, tls_enabled,
    association_time_sec, final_status, failure_reason, exception_details."""
    _log(PUSH_LOG, pid, pname, "", error)
    record = {
        "event": "push",
        "patient_id": pid,
        "patient_name": pname,
        "error_message": error or None,
        "final_status": details.pop("final_status", "FAILED" if error else "OK"),
    }
    record.update(details)
    _write_jsonl(PUSH_LOG_JSONL, record)
    rotate_logs_if_needed()


def _audit_username():
    """Prefers the authenticated identity (LDAP domain login) over the
    raw OS username, so audit records reflect who actually logged in
    through this app rather than just which OS account the process
    happens to run under."""
    return current_identity.get("username") or _os_username()


def write_audit_log(event_type, detail, **details):
    """Connection / action audit trail: who connected, when, what happened.
    Append-only, human-readable, separate from error logs.

    **details is additive/optional; when omitted, username/computer_name/
    result are still auto-populated so every audit record satisfies the
    enterprise audit schema (Timestamp, Username, Computer Name, IP
    Address, Action, Result, Details) even from older call sites."""
    try:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with _log_write_lock:
            with open(AUDIT_LOG, "a", encoding="utf-8") as f:
                f.write(f"[{now}] {event_type}: {detail}\n")
    except Exception:
        log_exception("Failed to write audit log entry")

    record = {
        "event": "audit",
        "action": event_type,
        "details": detail,
        "username": details.pop("username", _audit_username()) if details else _audit_username(),
        "computer_name": _computer_name(),
        "result": details.pop("result", "INFO") if details else "INFO",
    }
    record.update(details)
    _write_jsonl(AUDIT_LOG_JSONL, record)
    rotate_logs_if_needed()

    # Generic "Configuration Changes" notification hook -- covers every
    # config persistence point (destinations, routing rules, SOP editor,
    # log retention, notification settings, admin PIN) without needing a
    # separate notify_event() call sprinkled at each one individually.
    if event_type in CONFIG_CHANGE_AUDIT_EVENTS:
        try:
            notify_event("configuration_changes", "Configuration Changed", f"{event_type}: {detail}")
        except Exception:
            log_exception("notify_event failed from write_audit_log config-change hook")


def tail_log_file(path, max_lines=500):
    """Return the last `max_lines` lines of a log file, newest last.
    Used by the in-GUI log viewer. Safe if the file doesn't exist yet."""
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        return lines[-max_lines:]
    except Exception:
        log_exception(f"Failed to tail log file {path}")
        return []


# ---------------------------------------------------------
# Retention configuration
# ---------------------------------------------------------

def load_log_retention_days():
    """Returns the configured retention window in days (30/60/90/180/365).
    Not secret, so it's a plain JSON file like routing_rules.json/
    tls_config.json rather than an encrypted config."""
    try:
        if os.path.exists(LOG_RETENTION_CONFIG_FILE):
            with open(LOG_RETENTION_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            days = int(cfg.get("retention_days", DEFAULT_LOG_RETENTION_DAYS))
            if days in VALID_LOG_RETENTION_DAYS:
                return days
    except Exception:
        log_exception("Failed to load log retention config")
    return DEFAULT_LOG_RETENTION_DAYS


def save_log_retention_days(days):
    if days not in VALID_LOG_RETENTION_DAYS:
        raise ValueError(f"Retention must be one of {VALID_LOG_RETENTION_DAYS}")
    try:
        atomic_write(LOG_RETENTION_CONFIG_FILE, json.dumps({"retention_days": days}, indent=2))
        write_audit_log("LOG-RETENTION-CHANGED", f"retention_days={days}")
    except Exception:
        log_exception("Failed to save log retention config")
        raise


# ---------------------------------------------------------
# Rotation / archival
# ---------------------------------------------------------

_last_rotation_check = [0.0]
_ROTATION_CHECK_INTERVAL_SEC = 300  # avoid stat()'ing every log on every write


def _log_file_age_days(path):
    try:
        mtime = os.path.getmtime(path)
        return (time.time() - mtime) / 86400.0
    except Exception:
        return 0.0


def archive_log_file(log_path, jsonl_path=None, reason="rotation"):
    """Compress log_path (+ its .jsonl twin, if any) into a timestamped
    .zip inside LOG_ARCHIVE_DIR, then truncate-and-restart the live log
    file so writes can continue immediately. The archive is never deleted
    by anything in this application."""
    try:
        os.makedirs(LOG_ARCHIVE_DIR, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.splitext(os.path.basename(log_path))[0]
        archive_name = f"{base}_{stamp}.zip"
        archive_path = os.path.join(LOG_ARCHIVE_DIR, archive_name)

        with _log_write_lock:
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as zf:
                if os.path.exists(log_path) and os.path.getsize(log_path) > 0:
                    zf.write(log_path, arcname=os.path.basename(log_path))
                if jsonl_path and os.path.exists(jsonl_path) and os.path.getsize(jsonl_path) > 0:
                    zf.write(jsonl_path, arcname=os.path.basename(jsonl_path))
            # Start fresh live logs (never delete -- just rotate content out)
            if os.path.exists(log_path):
                open(log_path, "w").close()
            if jsonl_path and os.path.exists(jsonl_path):
                open(jsonl_path, "w").close()

        write_audit_log("LOG-ROTATED", f"log={log_path} archive={archive_path} reason={reason}")
        return archive_path
    except Exception:
        log_exception(f"Failed to archive log file {log_path}")
        return None


def rotate_logs_if_needed(force=False):
    """Cheap, frequently-called guard: only actually stats the log files
    every _ROTATION_CHECK_INTERVAL_SEC seconds (or immediately if
    force=True, e.g. from a manual 'Archive Now' button or the daily
    background timer)."""
    now = time.time()
    if not force and (now - _last_rotation_check[0]) < _ROTATION_CHECK_INTERVAL_SEC:
        return
    _last_rotation_check[0] = now

    retention_days = load_log_retention_days()
    pairs = list(zip(ALL_LOG_FILES, ALL_LOG_FILES_JSONL))
    for log_path, jsonl_path in pairs:
        try:
            needs_rotation = False
            if os.path.exists(log_path):
                if os.path.getsize(log_path) >= get_cache_size_limit_bytes():
                    needs_rotation = True
                elif _log_file_age_days(log_path) >= retention_days:
                    needs_rotation = True
            if needs_rotation:
                reason = "size-threshold" if os.path.exists(log_path) and os.path.getsize(log_path) >= get_cache_size_limit_bytes() else "retention-window"
                archive_log_file(log_path, jsonl_path, reason=reason)
        except Exception:
            log_exception(f"Rotation check failed for {log_path}")


def list_log_archives():
    """Returns archive metadata sorted newest-first: [(filename, path,
    size_bytes, created_dt), ...]. Used by the Logs tab archive browser."""
    results = []
    try:
        if not os.path.isdir(LOG_ARCHIVE_DIR):
            return results
        for fname in os.listdir(LOG_ARCHIVE_DIR):
            if not fname.lower().endswith(".zip"):
                continue
            fpath = os.path.join(LOG_ARCHIVE_DIR, fname)
            try:
                size = os.path.getsize(fpath)
                created = datetime.datetime.fromtimestamp(os.path.getmtime(fpath))
            except Exception:
                size, created = 0, None
            results.append((fname, fpath, size, created))
        results.sort(key=lambda r: r[3] or datetime.datetime.min, reverse=True)
    except Exception:
        log_exception("Failed to list log archives")
    return results


def read_archived_log(archive_path, max_lines=1000):
    """Extracts and returns the text content of an archived log .zip for
    in-GUI viewing (concatenates whatever members it contains, newest
    logic mirrors tail_log_file)."""
    try:
        chunks = []
        with zipfile.ZipFile(archive_path, "r") as zf:
            for name in zf.namelist():
                with zf.open(name) as member:
                    text = member.read().decode("utf-8", errors="replace")
                    chunks.append(f"----- {name} -----\n{text}")
        combined = "\n".join(chunks)
        lines = combined.splitlines(keepends=True)
        return "".join(lines[-max_lines:]) if lines else "(archive is empty)"
    except Exception:
        log_exception(f"Failed to read archived log {archive_path}")
        return f"(failed to read archive: {archive_path})"


def export_logs_zip(dest_path, include_archives=True):
    """Bundles the live logs (plain + structured) and, optionally, every
    archived log into a single ZIP the user can save anywhere. This is
    the ONLY way logs leave the app other than viewing them -- there is
    no deletion path."""
    try:
        with zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname in ALL_LOG_FILES + ALL_LOG_FILES_JSONL:
                if os.path.exists(fname):
                    zf.write(fname, arcname=os.path.join("live", fname))
            if include_archives and os.path.isdir(LOG_ARCHIVE_DIR):
                for fname in os.listdir(LOG_ARCHIVE_DIR):
                    if fname.lower().endswith(".zip"):
                        zf.write(os.path.join(LOG_ARCHIVE_DIR, fname), arcname=os.path.join("archives", fname))
        write_audit_log("LOG-EXPORTED", f"dest={dest_path} include_archives={include_archives}")
        return True
    except Exception:
        log_exception(f"Failed to export logs to {dest_path}")
        return False

# =========================================================
# WORKLIST HELPERS
# =========================================================

def upsert_patient(pid, pname, institution, study_uid, modality, source, status=None,
                    is_document=False, is_duplicate=False):
    """Create or update a worklist entry. Thread-safe.
    `source` is either the calling AE Title (network receive), the literal
    string "Imported" (folder import), or any caller-supplied label.
    `is_document` -- True when the instance being recorded is one of the
    encapsulated-document SOP Classes (see DOCUMENT_SOP_CLASS_UIDS): it
    increments "doc_count" instead of "count", so encapsulated PDFs/CDA/
    mesh documents are tallied separately from actual images instead of
    silently inflating the image count.
    `is_duplicate` -- True when the caller already determined (by checking
    whether the destination file existed before this write -- see
    handle_store()/import_folder()) that this SOP Instance UID was already
    saved for this patient. Every "#"/image-count display in the app reads
    "count"/"doc_count" directly off this entry, so a duplicate MUST NOT
    increment either counter: the file on disk is being overwritten in
    place (files are named "{SOPInstanceUID}.dcm"), not added to, so the
    true number of images on disk does not change and the displayed count
    must not either. This is what previously let a re-pushed study double-
    count in the Receiver worklist (e.g. 295 real files showing as 570)."""
    with data_lock:
        if pid in patient_data:
            if not is_duplicate:
                if is_document:
                    patient_data[pid]["doc_count"] = patient_data[pid].get("doc_count", 0) + 1
                else:
                    patient_data[pid]["count"] += 1
            patient_data[pid]["time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            patient_data[pid]["source"] = source
            # New images arriving for an already-Sent study should NOT silently
            # stay marked Sent -- that hides the fact there's new unsent data.
            # A genuine duplicate (same SOP Instance UID already on disk)
            # isn't new data, so it must not demote an already-Sent study
            # back to Pending -- that would falsely suggest there's
            # something new to push.
            if not is_duplicate and patient_data[pid].get("status") == STATUS_SENT:
                patient_data[pid]["status"] = STATUS_PENDING
        else:
            patient_data[pid] = {
                "patient_name": pname,
                "institution": institution,
                "study_uid": study_uid,
                "modality": modality,
                "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "count": 0 if is_document else 1,
                "doc_count": 1 if is_document else 0,
                "source": source,
                "status": status or (STATUS_IMPORTED if source == "Imported" else STATUS_RECEIVED),
                "sent_time": "",
                "push_target": "",
                "last_error": "",
                "retry_count": 0,
                "report_exists": False,
                "report_created_date": "",
                "report_last_opened": "",
            }
    autosave_csv()
    _bump_data_version()
    ui_event_queue.put(("refresh", None))


def set_status(pid, status):
    with data_lock:
        if pid in patient_data:
            patient_data[pid]["status"] = status
    autosave_csv()
    _bump_data_version()
    ui_event_queue.put(("refresh", None))


def set_fields(pid, **fields):
    with data_lock:
        if pid in patient_data:
            patient_data[pid].update(fields)
    autosave_csv()
    _bump_data_version()
    ui_event_queue.put(("refresh", None))


def get_studies_for_patient(pid):
    """Group locally-held files for a patient by StudyInstanceUID by
    reading the DICOM files on disk. Used for study-level grouping in the
    worklist tree. Returns dict: study_uid -> {"count": n, "files": [...],
    "modality": ..., "date": ...}."""
    folder = get_patient_folder(pid)
    studies = {}
    if not os.path.isdir(folder):
        return studies
    for fname in os.listdir(folder):
        if not fname.lower().endswith(".dcm"):
            continue
        fpath = os.path.join(folder, fname)
        try:
            ds = pydicom.dcmread(fpath, stop_before_pixels=True, force=True)
            study_uid = str(getattr(ds, "StudyInstanceUID", "UNKNOWN_STUDY"))
            modality = str(getattr(ds, "Modality", ""))
            study_date = str(getattr(ds, "StudyDate", ""))
        except Exception:
            study_uid, modality, study_date = "UNKNOWN_STUDY", "", ""
        entry = studies.setdefault(study_uid, {"count": 0, "files": [], "modality": modality, "date": study_date})
        entry["count"] += 1
        entry["files"].append(fpath)
    return studies


def compute_patient_disk_counts(pid):
    """Ground truth for how many DICOM files actually sit in this
    patient's folder right now, split into images vs. encapsulated
    documents. Every ``.dcm`` file is named ``{SOPInstanceUID}.dcm`` (see
    handle_store() / import_folder()), so re-receiving or re-importing the
    same SOP Instance overwrites the existing file rather than adding a
    new one -- this is the authoritative count "count"/"doc_count" are
    meant to track, and what reconcile_patient_image_counts() uses to
    self-heal any drift."""
    image_count = 0
    doc_count = 0
    for fpath in list_patient_dcm_files(pid):
        try:
            ds = pydicom.dcmread(fpath, stop_before_pixels=True, force=True)
            sop_class_uid = str(
                getattr(getattr(ds, "file_meta", None), "MediaStorageSOPClassUID", "")
                or getattr(ds, "SOPClassUID", "")
            )
        except Exception:
            sop_class_uid = ""
        if sop_class_uid in DOCUMENT_SOP_CLASS_UIDS:
            doc_count += 1
        else:
            image_count += 1
    return image_count, doc_count


def reconcile_patient_image_counts():
    """Self-heals patient_data[pid]["count"]/["doc_count"] against the
    files actually on disk for every known patient.

    upsert_patient() now keeps these counters in sync going forward (it
    only increments on a genuinely new SOP Instance UID -- see its
    docstring), but that alone doesn't fix counts that were already
    inflated by duplicate transfers before this fix, e.g. from a worklist
    CSV saved by an older version of the app. Run once, in the background,
    at startup so any such drift is corrected without requiring the user
    to re-import or manually reset anything. Safe to call any time --
    it's a pure resync against disk, not a mutation of any DICOM file."""
    try:
        with data_lock:
            pids = list(patient_data.keys())
        changed = False
        for pid in pids:
            image_count, doc_count = compute_patient_disk_counts(pid)
            with data_lock:
                entry = patient_data.get(pid)
                if entry is None:
                    continue
                if entry.get("count") != image_count or entry.get("doc_count") != doc_count:
                    entry["count"] = image_count
                    entry["doc_count"] = doc_count
                    changed = True
        if changed:
            autosave_csv()
            _bump_data_version()
            ui_event_queue.put(("refresh", None))
            app_logger.info("Reconciled worklist image/document counts against files on disk.")
    except Exception:
        log_exception("Failed to reconcile patient image counts against disk")


def start_count_reconciliation_thread():
    threading.Thread(target=reconcile_patient_image_counts, daemon=True, name="CountReconcile").start()


# =========================================================
# Shared filesystem helpers (still used by report generation, etc.)
# =========================================================

def _safe_document_op(op_label, func, *args, **kwargs):
    """Runs a filesystem operation and maps common failure modes to a
    friendly message. Returns (ok: bool, result_or_message)."""
    try:
        result = func(*args, **kwargs)
        return True, result
    except FileNotFoundError as e:
        return False, f"{op_label} failed: file or folder not found ({e.filename or e})."
    except FileExistsError as e:
        return False, f"{op_label} failed: a file with that name already exists ({e.filename or e})."
    except PermissionError as e:
        return False, (f"{op_label} failed: permission denied. The file may be open in "
                       f"another program, or you don't have access rights ({e.filename or e}).")
    except IsADirectoryError as e:
        return False, f"{op_label} failed: expected a file but found a folder ({e})."
    except OSError as e:
        # Covers locked files on Windows (WinError 32), disk-full, etc.
        return False, f"{op_label} failed: {e.strerror or e}."
    except Exception as e:
        log_exception(f"{op_label} failed unexpectedly")
        return False, f"{op_label} failed: {e}"


def format_dicom_person_name(raw):
    """DICOM 'PN' (Person Name) values are stored component-separated by
    the '^' character (e.g. 'Doe^John^A' for Last^First^Middle) -- that's
    a DICOM formatting rule, not literal text the patient's name contains.
    This converts that into a normal human-readable, space-separated
    name for display/storage everywhere in the UI (worklist, reports,
    history, exports, etc.). Safe to call on already-clean strings too."""
    if raw is None:
        return ""
    text = str(raw)
    if "^" not in text:
        return text.strip()
    parts = [p.strip() for p in text.split("^") if p.strip()]
    return " ".join(parts)


def open_document(path):
    """Launches a file with the OS default associated application.
    Returns (ok, message)."""
    if not os.path.isfile(path):
        return False, f"File not found: {os.path.basename(path)}. It may have been moved or deleted."
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path)  # noqa: S606 - intentional, OS-level file association
        elif system == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True, ""
    except PermissionError:
        return False, "Permission denied opening the file. It may be locked by another program."
    except Exception as e:
        log_exception(f"Failed to open document {path}")
        return False, f"Could not open the file with the system default application: {e}"


def open_patient_folder(pid):
    """Opens the OS file browser directly at a patient's folder (as
    opposed to open_containing_folder, which reveals a specific FILE's
    parent directory). Returns (ok, message). Same platform-branch
    pattern as open_document/open_containing_folder above."""
    folder = get_patient_folder(pid)
    if not os.path.isdir(folder):
        return False, f"No folder found on disk for patient {pid}."
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(folder)  # noqa: S606 - intentional, OS-level folder open
        elif system == "Darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
        return True, ""
    except Exception as e:
        log_exception(f"Failed to open patient folder for {pid}")
        return False, f"Could not open the patient folder: {e}"


def open_containing_folder(path):
    """Opens the OS file browser at the given file's location. Returns (ok, message)."""
    folder = os.path.dirname(path)
    if not os.path.isdir(folder):
        return False, "Folder no longer exists."
    try:
        system = platform.system()
        if system == "Windows":
            subprocess.Popen(f'explorer /select,"{path}"')
        elif system == "Darwin":
            subprocess.Popen(["open", "-R", path])
        else:
            subprocess.Popen(["xdg-open", folder])
        return True, ""
    except Exception as e:
        log_exception(f"Failed to open containing folder for {path}")
        return False, f"Could not open the containing folder: {e}"


# =========================================================
# PATIENT REPORT MANAGER (backend)
# =========================================================
# Separate feature from the Patient Document Manager above. Exactly ONE
# primary radiology report per patient, auto-generated from a
# professional template the first time it's requested, then opened
# in-place (in Microsoft Word, or whatever the OS has associated with
# .docx) on every subsequent request — never regenerated, never
# overwritten, no Save As dialog, because the file always lives at a
# single fixed path:
#
#   received_dicoms/<PatientID>/Reports/Radiology_Report.docx
#
# This reuses _safe_document_op / open_document / open_containing_folder
# from the Document Manager section above rather than duplicating
# filesystem-error handling, per the existing architecture.

REPORTS_SUBDIR = "Reports"
REPORT_FILENAME = "Radiology_Report.docx"  # base name only; get_report_path() embeds the patient ID

# "Report Created Date" is recorded by US at the moment we generate the
# file rather than read back from OS filesystem timestamps: st_ctime is
# "metadata change time" (not creation time) on Linux, and gets touched
# by every Word autosave on some filesystems/platforms. Recording it
# ourselves, once, at creation time is the only way to keep it accurate
# for the life of the report. "Last Modified" and "Report File Size",
# by contrast, are always read live from the file itself (a single cheap
# os.stat call), since Word is the one changing those, not us.


def get_reports_folder(pid):
    return os.path.join(get_patient_folder(pid), REPORTS_SUBDIR)


def get_report_path(pid):
    return os.path.join(get_reports_folder(pid), f"Radiology_Report_{_filename_safe_pid(pid)}.docx")


def ensure_reports_folder(pid):
    """Creates the patient + Reports folders if missing. Returns
    (ok, folder_path_or_error_message)."""
    def _make():
        folder = get_reports_folder(pid)
        os.makedirs(folder, exist_ok=True)
        return folder
    return _safe_document_op("Create Reports folder", _make)


# =========================================================
# GENERIC ATTACHMENTS (Part 2 of the doc-transfer audit)
# =========================================================
# Unlike Report/History above (which are exactly one fixed file each,
# with their own dedicated create/edit workflow), Attachments is an
# open folder: any number of files, any extension, optionally nested in
# subfolders that get preserved end to end. The doc-transfer wire
# protocol already carried an arbitrary relative_path + size per file
# from day one (see _doc_transfer_safe_dest_path / DOC_TRANSFER_MAGIC
# framing) -- the only thing that was ever hardcoded to two filenames
# was the SENDER's candidate list in push_patient_documents(). This
# section is the folder-scanning replacement for that; the wire
# protocol itself did not need to change.

ATTACHMENTS_SUBDIR = "Attachments"

# Files the scanner deliberately ignores: in-progress .part transfers
# (never a real attachment, always transient), Office lock files
# ("~$Document.docx"), and dotfiles (.DS_Store and friends).
_ATTACHMENT_IGNORED_SUFFIXES = (".part",)
_ATTACHMENT_IGNORED_PREFIXES = ("~$", ".")


def get_attachments_folder(pid):
    return os.path.join(get_patient_folder(pid), ATTACHMENTS_SUBDIR)


def ensure_attachments_folder(pid):
    """Creates the patient + Attachments folders if missing. Returns
    (ok, folder_path_or_error_message)."""
    def _make():
        folder = get_attachments_folder(pid)
        os.makedirs(folder, exist_ok=True)
        return folder
    return _safe_document_op("Create Attachments folder", _make)


def list_patient_attachments(pid):
    """Walks this patient's Attachments folder recursively. Returns a
    list of (relative_path, absolute_path, size_bytes) tuples -- one per
    real file, any filename, any extension, any subfolder depth.
    relative_path is always "Attachments/..." with forward slashes
    regardless of OS, ready to hand straight to the doc-transfer wire
    protocol or to _doc_transfer_safe_dest_path(). This is the single
    source of truth both push_patient_documents() (what to send) and
    the Inspector panel (what to show) read from, so they can never
    drift out of sync with each other."""
    folder = get_attachments_folder(pid)
    if not os.path.isdir(folder):
        return []
    out = []
    for root, _dirs, files in os.walk(folder):
        for fname in files:
            if fname.startswith(_ATTACHMENT_IGNORED_PREFIXES) or fname.endswith(_ATTACHMENT_IGNORED_SUFFIXES):
                continue
            abspath = os.path.join(root, fname)
            rel_within = os.path.relpath(abspath, folder)
            relative_path = f"{ATTACHMENTS_SUBDIR}/{rel_within}".replace(os.sep, "/")
            try:
                size = os.path.getsize(abspath)
            except OSError:
                size = 0
            out.append((relative_path, abspath, size))
    out.sort(key=lambda t: t[0].lower())
    return out


def add_patient_attachments(pid, source_paths, subfolder=""):
    """Copies each entry in source_paths into this patient's Attachments
    folder. Entries can be files OR directories -- a directory is walked
    recursively and every file inside it is added with the directory's
    own name (plus its internal subfolder structure) preserved, which is
    what actually satisfies 'preserve folder structure when appropriate'
    for a dropped/picked folder rather than a loose file. An explicit
    subfolder prefix (e.g. 'Consent Forms') applies on top of that.
    Destination filenames go through the same sanitization as every
    other filesystem write in this app; a same-name collision is
    numbered rather than silently overwritten. Returns
    (ok_count, [(source_path, error_message), ...]) so the caller can
    report partial failures without losing the rest of a batch."""
    ok, folder_or_err = ensure_attachments_folder(pid)
    if not ok:
        return 0, [(p, folder_or_err) for p in source_paths]
    attachments_root = folder_or_err

    base_sub = ""
    if subfolder:
        safe_sub = re.sub(r"[^A-Za-z0-9._ -]+", "_", subfolder).strip("_ ")
        if safe_sub:
            base_sub = safe_sub

    # Expand any directories into (file_path, extra_subfolder) pairs
    # before doing any actual copying, so a bad entry deep in one
    # dropped folder can't abort files from a different, valid entry in
    # the same batch.
    expanded = []
    errors = []
    for src in source_paths:
        if os.path.isdir(src):
            folder_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", os.path.basename(src.rstrip("/\\"))).strip("_ ") or "folder"
            found_any = False
            for root, _dirs, files in os.walk(src):
                rel_inside = os.path.relpath(root, src)
                for fname in files:
                    if fname.startswith(_ATTACHMENT_IGNORED_PREFIXES) or fname.endswith(_ATTACHMENT_IGNORED_SUFFIXES):
                        continue
                    extra = folder_name if rel_inside == "." else os.path.join(folder_name, rel_inside)
                    expanded.append((os.path.join(root, fname), extra))
                    found_any = True
            if not found_any:
                errors.append((src, "Folder is empty"))
        elif os.path.isfile(src):
            expanded.append((src, ""))
        else:
            errors.append((src, "Not a file or folder"))

    ok_count = 0
    for src, extra_sub in expanded:
        try:
            dest_root = attachments_root
            if base_sub:
                dest_root = os.path.join(dest_root, base_sub)
            if extra_sub:
                dest_root = os.path.join(dest_root, extra_sub)
            safe_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", os.path.basename(src)).strip("_ ") or "attachment"
            os.makedirs(dest_root, exist_ok=True)
            dest = os.path.join(dest_root, safe_name)
            base, ext = os.path.splitext(dest)
            n = 1
            while os.path.exists(dest):
                dest = f"{base} ({n}){ext}"
                n += 1
            shutil.copy2(src, dest)
            ok_count += 1
        except Exception as e:
            errors.append((src, str(e)))
    if ok_count:
        write_audit_log("ATTACHMENT-ADDED", f"pid={pid} count={ok_count}")
    return ok_count, errors


def remove_patient_attachment(pid, relative_path):
    """Deletes one attachment by its relative_path (as returned by
    list_patient_attachments). Reuses _doc_transfer_safe_dest_path's
    containment check rather than duplicating path-validation logic --
    it never trusts relative_path enough to delete outside the
    Attachments folder, same guarantee the receive path already has."""
    ok, dest_or_err = _doc_transfer_safe_dest_path(pid, relative_path)
    if not ok:
        return False, dest_or_err
    try:
        if os.path.isfile(dest_or_err):
            os.remove(dest_or_err)
            write_audit_log("ATTACHMENT-REMOVED", f"pid={pid} file={relative_path}")
            return True, ""
        return False, "File not found"
    except Exception as e:
        return False, str(e)


def get_report_metadata(pid):
    """Returns a dict for the 'internally track' requirement:
    exists, created_date, last_modified, last_opened, file_size_bytes.
    Touches the filesystem only if the report file is actually present
    (a single os.stat call) — never scans the folder."""
    with data_lock:
        d = patient_data.get(pid, {})
        created_date = d.get("report_created_date", "")
        last_opened = d.get("report_last_opened", "")

    path = get_report_path(pid)
    if not os.path.isfile(path):
        return {
            "exists": False, "created_date": created_date, "last_modified": "",
            "last_opened": last_opened, "file_size_bytes": 0, "path": path,
        }
    try:
        stat = os.stat(path)
        last_modified = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        size = stat.st_size
    except OSError:
        last_modified, size = "", 0
    return {
        "exists": True, "created_date": created_date, "last_modified": last_modified,
        "last_opened": last_opened, "file_size_bytes": size, "path": path,
    }


def _set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _add_paragraph_bottom_border(paragraph, sz=4, color="BFBFBF"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    edge = OxmlElement('w:bottom')
    edge.set(qn('w:val'), 'single')
    edge.set(qn('w:sz'), str(sz))
    edge.set(qn('w:space'), '1')
    edge.set(qn('w:color'), color)
    pBdr.append(edge)
    pPr.append(pBdr)


def _blank_fill_line(doc, height_pt=18):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(height_pt)
    _add_paragraph_bottom_border(p)
    return p


def _section_heading(doc, text):
    h = doc.add_heading(text, level=2)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        run.font.size = Pt(13)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(4)
    return h


def build_radiology_report_docx(path, pid):
    """Generates the professional hospital-style radiology report
    template, pre-filled with this patient's worklist info, and saves
    it to `path`. Only ever called for a BRAND NEW report (or as an
    explicit, user-confirmed recovery from a missing/corrupt file) —
    never on an existing, healthy report."""
    with data_lock:
        d = dict(patient_data.get(pid, {}))

    now = datetime.datetime.now()
    doc = DocxDocument()

    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)

    section = doc.sections[0]
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)

    institution = d.get("institution", "") or "Hospital / Clinic Name"

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(institution)
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    dept = doc.add_paragraph()
    dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = dept.add_run("Department of Radiology")
    r.font.size = Pt(12)
    r.italic = True

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("OFFICIAL RADIOLOGY REPORT")
    r.bold = True
    r.font.size = Pt(13)
    sub.paragraph_format.space_after = Pt(6)
    _add_paragraph_bottom_border(sub, sz=12, color="1F4E79")

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    info_heading = doc.add_paragraph()
    r = info_heading.add_run("Patient Information")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    rows = [
        ("Patient Name", d.get("patient_name", "") or "—", "Patient ID", pid),
        ("Institution", institution, "Study UID", d.get("study_uid", "") or "—"),
        ("Modality", d.get("modality", "") or "—", "Date Received", d.get("time", "") or "—"),
        ("Report Date", now.strftime("%Y-%m-%d"), "Report Time", now.strftime("%H:%M:%S")),
    ]
    table = doc.add_table(rows=len(rows), cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    widths = [Inches(1.3), Inches(2.4), Inches(1.3), Inches(2.4)]
    for r_idx, (l1, v1, l2, v2) in enumerate(rows):
        cells = table.rows[r_idx].cells
        for c_idx, text in enumerate((l1, v1, l2, v2)):
            cells[c_idx].width = widths[c_idx]
            run = cells[c_idx].paragraphs[0].add_run(text)
            if c_idx in (0, 2):
                run.bold = True
                _set_cell_shading(cells[c_idx], "DCE6F1")

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    for heading, blank_lines in [
        ("Clinical History", 3),
        ("Examination", 3),
        ("Findings", 5),
        ("Impression", 3),
        ("Recommendations", 3),
    ]:
        _section_heading(doc, heading)
        for _ in range(blank_lines):
            _blank_fill_line(doc)

    _section_heading(doc, "Radiologist")
    for label in ("Name:", "Signature:", "Date:"):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(14)
        run = p.add_run(label + "  ")
        run.bold = True
        _add_paragraph_bottom_border(p)

    # Digital signature block: intentionally left blank. Structured as a
    # 3-column table (not free text) so a future cryptographic-signing
    # feature can populate/lock these specific cells without changing
    # the document's overall structure, per the spec's forward-compat
    # requirement.
    _section_heading(doc, "Digital Signature")
    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.style = 'Table Grid'
    for i, h in enumerate(("Digital Signature", "Date", "Time")):
        cell = sig_table.rows[0].cells[i]
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        _set_cell_shading(cell, "F2F2F2")
        cell.add_paragraph("")  # blank line reserved for future signature content

    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run("This report is electronically generated.")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    doc.save(path)


def _is_docx_readable(path):
    """Cheap corruption pre-check: a .docx is a zip archive, so a quick
    zipfile validity check catches truncated/corrupt files without
    actually rendering or parsing the document content (which the spec
    says this app should never do)."""
    try:
        return zipfile.is_zipfile(path)
    except OSError:
        return False


def ensure_report(pid):
    """Guarantees exactly one primary report exists for this patient and
    returns a 3-state result without ever touching a healthy existing
    file:
        ("ok", path)       -- report exists (or was just created) and is usable
        ("corrupt", path)  -- report exists but fails the zip-validity check;
                              caller must get explicit user confirmation
                              before any recovery action is taken
        ("error", message) -- could not create/access the report at all
    This is the ONLY function that decides whether to generate a new
    report — and it only ever does so when the file is genuinely absent."""
    ok, folder_or_err = ensure_reports_folder(pid)
    if not ok:
        return "error", folder_or_err

    path = get_report_path(pid)

    if os.path.isfile(path):
        if not _is_docx_readable(path):
            return "corrupt", path
        return "ok", path  # existing, healthy report — never regenerated, never touched

    # Missing (first-time, or deleted) -> generate fresh from the template.
    ok, err = _safe_document_op("Generate report", build_radiology_report_docx, path, pid)
    if not ok:
        return "error", err
    set_fields(pid, report_exists=True,
              report_created_date=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    write_audit_log("REPORT-CREATED", f"pid={pid} path={path}")
    return "ok", path


def recover_corrupt_report(pid):
    """Explicit, user-confirmed recovery path: renames the unreadable
    file aside as a timestamped backup (never deletes it) and generates
    a fresh template in its place. Returns (ok, path_or_error)."""
    path = get_report_path(pid)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"Radiology_Report_corrupt_{ts}.docx"
    backup_path = os.path.join(get_reports_folder(pid), backup_name)

    ok, err = _safe_document_op("Back up corrupt report", os.rename, path, backup_path)
    if not ok:
        return False, err
    write_audit_log("REPORT-CORRUPT-BACKED-UP", f"pid={pid} backup={backup_path}")

    ok, err = _safe_document_op("Generate report", build_radiology_report_docx, path, pid)
    if not ok:
        return False, err
    set_fields(pid, report_exists=True,
              report_created_date=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    write_audit_log("REPORT-RECREATED-AFTER-CORRUPTION", f"pid={pid} path={path}")
    return True, path


def open_report(pid):
    """The single entry point the UI calls for 'Open Report'. Returns
    (status, path_or_message) where status is 'ok', 'corrupt', or
    'error' — mirroring ensure_report's contract so the UI layer can
    show the corruption-recovery confirmation dialog when needed."""
    status, result = ensure_report(pid)
    if status != "ok":
        return status, result

    path = result
    ok, err = open_document(path)  # reuses the Document Manager's OS-launch + error handling
    if not ok:
        write_audit_log("REPORT-OPEN-FAILED", f"pid={pid} path={path} error={err}")
        return "error", err

    set_fields(pid, report_last_opened=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    write_audit_log("REPORT-OPENED", f"pid={pid} path={path}")
    return "ok", path


# =========================================================
# PATIENT HISTORY MANAGER (backend)
# =========================================================
# Exactly one plain-text clinical history file per patient, following the
# same "create once, then just open" contract as the Report manager above
#   received_dicoms/<PatientID>/Reports/Patient_History.txt
# so it lives right alongside that patient's report, and reuses the same
# ensure_reports_folder / open_document / _safe_document_op plumbing.

HISTORY_FILENAME = "Patient_History.txt"  # base name only; get_history_path() embeds the patient ID


def get_history_path(pid):
    return os.path.join(get_reports_folder(pid), f"Patient_History_{_filename_safe_pid(pid)}.txt")


def ensure_history_file(pid):
    """Guarantees a history .txt file exists for this patient (creating an
    empty, friendly starter file the first time) and returns
    (ok, path_or_error_message). Never overwrites an existing file."""
    ok, folder_or_err = ensure_reports_folder(pid)
    if not ok:
        return False, folder_or_err

    path = get_history_path(pid)
    if os.path.isfile(path):
        return True, path

    def _make():
        with data_lock:
            pname = patient_data.get(pid, {}).get("patient_name", "")
        header = (
            f"Patient History\n"
            f"Patient ID: {pid}\n"
            f"Patient Name: {pname}\n"
            f"Created: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"{'-' * 50}\n\n"
        )
        atomic_write(path, header)
        return path

    ok, err = _safe_document_op("Create patient history file", _make)
    if not ok:
        return False, err
    set_fields(pid, history_exists=True,
              history_created_date=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    write_audit_log("HISTORY-CREATED", f"pid={pid} path={path}")
    return True, path


def open_history(pid):
    """The single entry point the UI calls for 'Open History'. Returns
    (ok, path_or_message)."""
    ok, result = ensure_history_file(pid)
    if not ok:
        return False, result

    path = result
    ok2, err = open_document(path)
    if not ok2:
        write_audit_log("HISTORY-OPEN-FAILED", f"pid={pid} path={path} error={err}")
        return False, err

    set_fields(pid, history_last_opened=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    write_audit_log("HISTORY-OPENED", f"pid={pid} path={path}")
    return True, path


# =========================================================
# DOCUMENT TRANSFER (Radiology_Report.docx / Patient_History.txt)
# =========================================================
# Transfers the pusher's real, edited Reports/ files to the receiver over
# one small TCP connection per patient, so the receiver stops fabricating
# its own blank copies via ensure_report()/ensure_history_file() -- those
# two functions are unmodified and, since they never overwrite an existing
# file, correctly leave alone whatever this feature writes to disk first.
#
# Wire protocol (see spec): magic + version + JSON connection header, then
# per-file JSON header + raw bytes, all length-prefixed. v1 verifies size
# only -- the optional sha256 field is carried but unenforced (see
# plug-in-point comments below).

# A.3: bounds how many doc-transfer connections are actively being handled
# (i.e. past accept()) at once. Additional connections are still accept()ed
# immediately (so the loop never stalls) but are rejected right away.
DOC_TRANSFER_MAX_CONCURRENT_CONNECTIONS = 8
_doc_transfer_connection_semaphore = threading.Semaphore(DOC_TRANSFER_MAX_CONCURRENT_CONNECTIONS)

# A.4: per-patient-ID lock registry so two connections for the same pid can
# never write to the same .part file concurrently. Protected by its own
# module-level lock since the dict itself is mutated from multiple threads.
_doc_transfer_pid_locks = {}
_doc_transfer_pid_locks_guard = threading.Lock()


def _doc_transfer_acquire_pid_lock(pid):
    with _doc_transfer_pid_locks_guard:
        lock = _doc_transfer_pid_locks.get(pid)
        if lock is None:
            lock = threading.Lock()
            _doc_transfer_pid_locks[pid] = lock
    lock.acquire()
    return lock


def _doc_transfer_release_pid_lock(pid, lock):
    lock.release()
    # Opportunistic cleanup: if nobody else is waiting on this lock right
    # now, drop it from the registry so a long-running receiver doesn't
    # accumulate one Lock object per patient ID forever.
    with _doc_transfer_pid_locks_guard:
        current = _doc_transfer_pid_locks.get(pid)
        if current is lock and lock.acquire(blocking=False):
            lock.release()
            del _doc_transfer_pid_locks[pid]


def _doc_transfer_receiver_disk_gb():
    """A.5-adjacent helper: how much free disk space the receiver should
    check against before writing a doc-transfer file. Reuses the same
    get_free_disk_gb()/LOW_DISK_WARNING_GB the DICOM receive path already
    warns with (D.5) -- just applied as a hard pre-write rejection here
    instead of a post-write warning."""
    return get_free_disk_gb(OUTPUT_DIR if os.path.isdir(OUTPUT_DIR) else ".")


def _doc_transfer_backup_existing(dest_path):
    """D.6: if dest_path already exists, copy it to a single '<name>.bak'
    sibling before it gets overwritten by the incoming file. Keeps only the
    most recent backup per file (no unbounded accumulation). Returns
    (overwrite_occurred, backup_created) -- usually the same value, but a
    backup can legitimately fail (e.g. disk full) even when an overwrite
    is genuinely about to happen, so the Receiver Activity page tracks
    them as two separate columns rather than one collapsed flag."""
    overwrite_occurred = os.path.isfile(dest_path)
    backup_created = False
    if overwrite_occurred:
        try:
            shutil.copy2(dest_path, dest_path + ".bak")
            backup_created = True
        except Exception:
            log_exception(f"Failed to back up previous file before overwrite: {dest_path}")
    return overwrite_occurred, backup_created


def _doc_transfer_recv_exact(sock, n):
    """Reads exactly n bytes from a socket or raises ConnectionError.
    Shared by both the pusher (client) and receiver (server) sides."""
    if n == 0:
        return b""
    buf = bytearray()
    while len(buf) < n:
        chunk = sock.recv(min(DOC_TRANSFER_CHUNK_SIZE, n - len(buf)))
        if not chunk:
            raise ConnectionError("Connection closed before expected data was received")
        buf.extend(chunk)
    return bytes(buf)


def _doc_transfer_drain(sock, n):
    """Reads and discards exactly n bytes, keeping a multi-file connection
    correctly framed after a per-file rejection (e.g. unsafe path)."""
    remaining = n
    while remaining > 0:
        chunk = sock.recv(min(DOC_TRANSFER_CHUNK_SIZE, remaining))
        if not chunk:
            raise ConnectionError("Connection closed while draining rejected file")
        remaining -= len(chunk)


def _doc_transfer_send_reply(sock, ok, reason=""):
    sock.sendall(bytes([0x01 if ok else 0x00]))
    if not ok:
        reason_bytes = (reason or "").encode("utf-8")
        sock.sendall(len(reason_bytes).to_bytes(2, "big") + reason_bytes)


def _doc_transfer_safe_dest_path(pid, relative_path):
    """Validates relative_path resolves to a location still inside this
    patient's folder (defends against ../ traversal), then returns that
    validated path as-is. IMPORTANT: this preserves whatever top-level
    folder the sender declared (Reports/, Attachments/, or a nested
    Attachments/<subfolder>/...) and the full subfolder structure under
    it -- it must NOT flatten to get_reports_folder(pid) + basename,
    which would silently misfile every Attachment into the Reports
    folder (losing the Attachments/ prefix entirely) and discard any
    subfolder structure. This happened to be invisible for Report/
    History before Attachments existed, since those relative_paths
    already started with "Reports/" -- same result either way for that
    one case, which is why this stayed unnoticed. Returns
    (ok, dest_path_or_error_message)."""
    try:
        raw = relative_path or ""
        if raw.endswith("/") or raw.endswith("\\"):
            return False, f"Rejected relative_path with no filename: {relative_path!r}"
        normalized = os.path.normpath(raw)
        patient_folder = os.path.abspath(get_patient_folder(pid))
        candidate = os.path.abspath(os.path.join(patient_folder, normalized))
        if candidate != patient_folder and not candidate.startswith(patient_folder + os.sep):
            return False, f"Rejected relative_path outside patient folder: {relative_path!r}"
        basename = os.path.basename(normalized)
        if not basename or basename in (".", ".."):
            return False, f"Rejected relative_path with no filename: {relative_path!r}"
        return True, candidate
    except Exception as e:
        return False, f"Path validation failed: {e}"


def _doc_transfer_receive_one_file(conn, pid, authenticated=False, source_addr=None):
    """Receives one file header + payload on an already-handshaked
    connection. Writes to <path>.part and atomically os.replace()s into
    place only after the byte count matches the declared size -- a
    dropped connection can never leave a corrupt-but-present file behind.
    `authenticated` and `source_addr` are purely for the activity log /
    trust indicator (Part 3 of the doc-transfer audit) -- they don't
    affect anything about how the file itself is received."""
    filename = ""
    started = time.time()
    source_ip = source_addr[0] if source_addr else None
    with data_lock:
        d = patient_data.get(pid, {})
        pname = d.get("patient_name", "")
        institution = d.get("institution", "")

    def _log_kwargs(**extra):
        base = {"authenticated": authenticated, "source_ip": source_ip}
        base.update(extra)
        return base

    try:
        header_len = int.from_bytes(_doc_transfer_recv_exact(conn, 2), "big")
        file_header = json.loads(_doc_transfer_recv_exact(conn, header_len).decode("utf-8"))
        relative_path = str(file_header.get("relative_path", ""))
        size = int(file_header.get("size", 0) or 0)
        # Real sha256 verification plug-in point: after the size check
        # below passes, before the .part -> final rename, hash the .part
        # file and compare against file_header.get("sha256") here.
        filename = os.path.basename(relative_path) or "(unknown)"
        if relative_path.startswith(REPORTS_SUBDIR + "/") and filename == os.path.basename(get_report_path(pid)):
            transfer_type = "Report"
        elif relative_path.startswith(REPORTS_SUBDIR + "/") and filename == os.path.basename(get_history_path(pid)):
            transfer_type = "History"
        elif relative_path.startswith(ATTACHMENTS_SUBDIR + "/"):
            transfer_type = "Attachment"
        else:
            transfer_type = "Other"

        _settings = load_app_settings()
        try:
            max_size_bytes = max(1, int(_settings.get("doc_transfer_max_size_mb", 500))) * 1024 * 1024
        except Exception:
            max_size_bytes = DOC_TRANSFER_MAX_FILE_SIZE_BYTES

        if size < 0:
            # A negative declared size can't be trusted enough to safely
            # drain and keep the connection framed for any further files
            # -- close outright rather than guess.
            err = "Rejected: invalid (negative) file size in header"
            _doc_transfer_send_reply(conn, False, err)
            write_receiver_log(pid, pname, institution, err,
                               result="FAILURE", event_type="DOC-TRANSFER", filename=filename,
                               **_log_kwargs(transfer_type=transfer_type))
            return

        if size > max_size_bytes:
            err = (f"Rejected: {filename} is {size / (1024 * 1024):.1f} MB, "
                   f"exceeds the {max_size_bytes / (1024 * 1024):.0f} MB doc-transfer limit")
            # size is well-formed (just too big) -- still drain it so a
            # multi-file connection stays correctly framed for the next
            # file instead of having to be torn down.
            _doc_transfer_drain(conn, size)
            _doc_transfer_send_reply(conn, False, err)
            write_receiver_log(pid, pname, institution, err,
                               result="FAILURE", event_type="DOC-TRANSFER",
                               filename=filename, file_size=size,
                               **_log_kwargs(transfer_type=transfer_type))
            ui_event_queue.put(("toast", ("Document Rejected",
                                          f"{pid}: {filename} exceeds the {max_size_bytes // (1024 * 1024)} MB limit")))
            return

        ok, dest_path_or_err = _doc_transfer_safe_dest_path(pid, relative_path)
        if not ok:
            _doc_transfer_drain(conn, size)
            _doc_transfer_send_reply(conn, False, dest_path_or_err)
            write_receiver_log(pid, pname, institution, dest_path_or_err,
                               result="FAILURE", event_type="DOC-TRANSFER", filename=filename,
                               **_log_kwargs(transfer_type=transfer_type))
            write_audit_log("DOC-TRANSFER-FAILED", f"pid={pid} file={filename}")
            return

        dest_path = dest_path_or_err

        # D.5: reject before touching disk if free space is already below
        # the same LOW_DISK_WARNING_GB threshold the DICOM receive path
        # warns on, mirroring that check but as a hard pre-write rejection.
        free_gb = _doc_transfer_receiver_disk_gb()
        if free_gb is not None and free_gb < LOW_DISK_WARNING_GB:
            err = f"Rejected: low disk space ({free_gb:.2f} GB free)"
            _doc_transfer_drain(conn, size)
            _doc_transfer_send_reply(conn, False, err)
            write_receiver_log(pid, pname, institution, err,
                               result="FAILURE", event_type="DOC-TRANSFER", filename=filename,
                               **_log_kwargs(transfer_type=transfer_type))
            write_audit_log("DOC-TRANSFER-FAILED", f"pid={pid} file={filename}")
            return

        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        part_path = dest_path + ".part"

        try:
            recv_chunk_size = max(1024, int(_settings.get("doc_transfer_chunk_size_kb", 64)) * 1024)
        except Exception:
            recv_chunk_size = DOC_TRANSFER_CHUNK_SIZE

        written = 0
        with open(part_path, "wb") as f:
            remaining = size
            while remaining > 0:
                chunk = conn.recv(min(recv_chunk_size, remaining))
                if not chunk:
                    raise ConnectionError("Connection closed mid-file")
                f.write(chunk)
                written += len(chunk)
                remaining -= len(chunk)
                bandwidth_limiter.throttle(len(chunk))

        if written != size:
            try:
                os.remove(part_path)
            except OSError:
                pass
            err = f"Size mismatch: expected {size}, received {written}"
            _doc_transfer_send_reply(conn, False, err)
            write_receiver_log(pid, pname, institution, err,
                               result="FAILURE", event_type="DOC-TRANSFER",
                               filename=filename, file_size=written,
                               **_log_kwargs(transfer_type=transfer_type))
            write_audit_log("DOC-TRANSFER-FAILED", f"pid={pid} file={filename}")
            return

        overwrite_occurred, backup_created = _doc_transfer_backup_existing(dest_path)  # D.6
        os.replace(part_path, dest_path)

        if filename == os.path.basename(get_report_path(pid)):
            set_fields(pid, report_exists=True, last_doc_transfer_error="")
        elif filename == os.path.basename(get_history_path(pid)):
            set_fields(pid, history_exists=True, last_doc_transfer_error="")
        else:
            set_fields(pid, last_doc_transfer_error="")

        duration_sec = round(time.time() - started, 3)
        speed_mbps = round((size / (1024 * 1024)) / duration_sec, 2) if duration_sec > 0 and size > 0 else None
        write_receiver_log(pid, pname, institution, "",
                           result="SUCCESS", event_type="DOC-TRANSFER",
                           filename=filename, file_size=size,
                           duration_sec=duration_sec,
                           transfer_speed_mbps=speed_mbps,
                           save_location=dest_path,
                           overwrite_occurred=overwrite_occurred,
                           backup_created=backup_created,
                           **_log_kwargs(transfer_type=transfer_type))
        write_audit_log("DOC-TRANSFER-RECEIVED", f"pid={pid} file={filename}")
        ui_event_queue.put(("toast", ("Document Received", f"{pid}: {filename}")))
        _doc_transfer_send_reply(conn, True)

    except Exception as e:
        log_exception(f"Failed receiving document for pid={pid}")
        try:
            write_receiver_log(pid, pname, institution, str(e),
                               result="FAILURE", event_type="DOC-TRANSFER", filename=filename,
                               **_log_kwargs(transfer_type=locals().get("transfer_type", "Unknown")))
            write_audit_log("DOC-TRANSFER-FAILED", f"pid={pid} file={filename}")
        except Exception:
            # The failure-reporting itself failed (e.g. disk full, which
            # may be the very reason the original write failed too).
            # log_exception() above already wrote a traceback for the
            # ORIGINAL failure, but that meta-failure needs its own
            # record or the incident could otherwise vanish from every
            # log if app.log and the receiver/audit logs share a volume.
            log_exception(f"Also failed to write receiver/audit log entries for the "
                          f"above failure (pid={pid}, file={filename})")
        try:
            _doc_transfer_send_reply(conn, False, str(e))
        except Exception:
            pass


def _reject_doc_transfer_connection_busy(conn, addr):
    """A.3: used when DOC_TRANSFER_MAX_CONCURRENT_CONNECTIONS is already
    saturated. Still performs the handshake read so the wire stays framed
    correctly, then replies with a REJECT (server busy) instead of READY,
    and closes -- never counted against the semaphore, never runs the
    per-file loop."""
    try:
        conn.settimeout(DOC_TRANSFER_CONNECT_TIMEOUT_SEC)
        magic = _doc_transfer_recv_exact(conn, 4)
        if magic != DOC_TRANSFER_MAGIC:
            _doc_transfer_send_reply(conn, False, "Bad magic")
            return
        version = _doc_transfer_recv_exact(conn, 1)[0]
        if version != DOC_TRANSFER_PROTO_VERSION:
            _doc_transfer_send_reply(conn, False, f"Unsupported protocol version {version}")
            return
        header_len = int.from_bytes(_doc_transfer_recv_exact(conn, 2), "big")
        _doc_transfer_recv_exact(conn, header_len)  # drain, ignore contents
        _doc_transfer_send_reply(conn, False, "Server busy, try again")
    except Exception:
        pass
    finally:
        try:
            conn.close()
        except Exception:
            pass


def _handle_doc_transfer_connection(conn, addr):
    """Handles one pusher's document-transfer connection for one patient.
    Runs on its own short-lived worker thread (see doc_transfer_accept_loop)
    -- never blocks the accept loop, never touches Tk directly."""
    pid = ""
    pid_lock = None
    try:
        conn.settimeout(DOC_TRANSFER_CONNECT_TIMEOUT_SEC)
        magic = _doc_transfer_recv_exact(conn, 4)
        if magic != DOC_TRANSFER_MAGIC:
            _doc_transfer_send_reply(conn, False, "Bad magic")
            return
        version = _doc_transfer_recv_exact(conn, 1)[0]
        if version != DOC_TRANSFER_PROTO_VERSION:
            _doc_transfer_send_reply(conn, False, f"Unsupported protocol version {version}")
            return
        header_len = int.from_bytes(_doc_transfer_recv_exact(conn, 2), "big")
        try:
            header = json.loads(_doc_transfer_recv_exact(conn, header_len).decode("utf-8"))
        except Exception:
            _doc_transfer_send_reply(conn, False, "Malformed connection header")
            return

        pid = str(header.get("patient_id", "")).strip()
        try:
            file_count = int(header.get("file_count", 0) or 0)
        except (TypeError, ValueError):
            file_count = -1

        # Auth key enforcement: if this receiver has a key configured
        # (doc_transfer_receiver_auth_key, see Settings), every sender
        # must present the exact same key or the connection is rejected
        # right here, before READY and before any per-file loop runs. An
        # empty receiver key means auth is deliberately off (backward
        # compatible with existing zero-config deployments) -- that is
        # now an explicit, visible setting rather than a silent no-op.
        # secrets.compare_digest is used instead of == to avoid leaking
        # timing information about how much of the key matched.
        expected_key = str(load_app_settings().get("doc_transfer_receiver_auth_key", "") or "")
        authenticated = False
        if expected_key:
            presented_key = str(header.get("auth_key", "") or "")
            if not secrets.compare_digest(presented_key, expected_key):
                _doc_transfer_send_reply(conn, False, "Authentication failed")
                write_audit_log("DOC-TRANSFER-AUTH-FAILED",
                                f"pid={pid or 'unknown'} addr={addr[0] if addr else 'unknown'}")
                return
            authenticated = True  # key was required AND matched

        if not pid or file_count < 0:
            _doc_transfer_send_reply(conn, False, "Invalid patient_id/file_count")
            return

        # A.4: serialize doc-transfer writes per patient ID so two
        # concurrent connections for the same pid can never race on the
        # same .part file.
        pid_lock = _doc_transfer_acquire_pid_lock(pid)

        conn.sendall(bytes([0x01]))  # READY

        # A.1: file_count is kept for logging/back-compat but the loop is now
        # continuation-byte driven so the pusher can retry a single file
        # (re-sending its header+payload as "another one coming") without
        # having to guess how many total attempts it will need up front.
        # 0x01 = another file header follows; 0x00 = pusher is done.
        while True:
            cont = _doc_transfer_recv_exact(conn, 1)
            if cont[0] == 0x00:
                break
            _doc_transfer_receive_one_file(conn, pid, authenticated=authenticated, source_addr=addr)

    except Exception as e:
        log_exception(f"Document-transfer connection failed (pid={pid})")
        try:
            write_receiver_log(pid, "", "", str(e), result="FAILURE", event_type="DOC-TRANSFER")
        except Exception:
            pass
    finally:
        if pid_lock is not None:
            try:
                _doc_transfer_release_pid_lock(pid, pid_lock)
            except Exception:
                pass
        try:
            conn.close()
        except Exception:
            pass
        # A.3: release the concurrency slot acquired by the accept loop
        # before this handler was dispatched.
        try:
            _doc_transfer_connection_semaphore.release()
        except Exception:
            pass


def doc_transfer_accept_loop(port, stop_event):
    """Daemon accept-loop for the document-transfer TCP port. Started
    alongside start_receiver_server() only when the global doc_transfer_enabled
    setting is on; stopped symmetrically by stop_receiver_server(). Each
    accepted connection is handed to its own short-lived worker thread so
    a slow/stalled pusher can never block new connections.

    A.6: binds "" (all families) via a dual-stack IPv6 socket so both IPv4
    and IPv6 pushers can connect, mirroring the fact that outbound
    connections elsewhere in this app (socket.create_connection) already
    resolve whichever family is reachable. If this platform doesn't
    support a dual-stack bind, falls back to the original IPv4-only
    binding rather than failing to listen at all."""
    # TLS plug-in point: wrap the listening socket / accepted connections
    # here with build_ssl_context_for_server(load_tls_config()) once
    # document-socket TLS is implemented.
    server_sock = None
    try:
        try:
            server_sock = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
            server_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                # Allow IPv4 connections on the same dual-stack socket where
                # the OS supports it; some platforms default this read-only.
                server_sock.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 0)
            except (AttributeError, OSError):
                pass
            server_sock.bind(("", int(port)))
        except OSError:
            # No usable IPv6 stack on this host -- fall back to the
            # original IPv4-only bind rather than not listening at all.
            if server_sock is not None:
                try:
                    server_sock.close()
                except Exception:
                    pass
            server_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            server_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            server_sock.bind(("0.0.0.0", int(port)))

        server_sock.listen(DOC_TRANSFER_ACCEPT_BACKLOG)
        server_sock.settimeout(1.0)  # periodic wake-up so stop_event is honored promptly
        write_audit_log("DOC-TRANSFER-SERVER-START", f"port={port}")
        while not stop_event.is_set():
            try:
                conn, addr = server_sock.accept()
            except socket.timeout:
                continue
            except OSError:
                break
            # A.3: still accept() immediately so the loop never stalls, but
            # if we're already at max concurrency, reject this one right
            # away instead of handing it to a worker thread.
            if _doc_transfer_connection_semaphore.acquire(blocking=False):
                threading.Thread(target=_handle_doc_transfer_connection,
                                  args=(conn, addr), daemon=True).start()
            else:
                threading.Thread(target=_reject_doc_transfer_connection_busy,
                                  args=(conn, addr), daemon=True).start()
    except Exception as e:
        log_exception("Document-transfer accept loop crashed")
        write_audit_log("DOC-TRANSFER-SERVER-CRASH", str(e))
    finally:
        try:
            server_sock.close()
        except Exception:
            pass
        write_audit_log("DOC-TRANSFER-SERVER-STOP", f"port={port}")


def push_patient_documents(pid, cfg):
    """Sends this patient's Report/History files AND every file in their
    Attachments folder (any name, any type, subfolders preserved) to
    `cfg`'s document-transfer endpoint, over one TCP connection. No-op
    (no log lines, no connection attempt) if this destination doesn't
    have document transfer enabled, or if there's nothing local to send.
    A failure here is recorded only via last_doc_transfer_error -- it
    never changes the patient's DICOM status (STATUS_SENT/STATUS_FAILED
    stay driven by image delivery only)."""
    if not cfg or not cfg.get("doc_transfer_enabled"):
        return

    candidates = []
    report_path = get_report_path(pid)
    if os.path.isfile(report_path):
        candidates.append((f"{REPORTS_SUBDIR}/{os.path.basename(report_path)}", report_path))
    history_path = get_history_path(pid)
    if os.path.isfile(history_path):
        candidates.append((f"{REPORTS_SUBDIR}/{os.path.basename(history_path)}", history_path))
    # Generic attachments: unlimited filenames, arbitrary types, folder
    # structure preserved via list_patient_attachments()'s relative_path.
    candidates.extend((rel, abspath) for rel, abspath, _size in list_patient_attachments(pid))

    if not candidates:
        return

    # Pre-flight size check: reject oversized files locally before ever
    # opening a connection, mirroring the same doc_transfer_max_size_mb
    # limit the receiver enforces server-side. Fails fast with a clear
    # UI-facing error instead of wasting a connection attempt (or worse,
    # streaming most of a large file) on something guaranteed to bounce.
    try:
        max_size_bytes = max(1, int(load_app_settings().get("doc_transfer_max_size_mb", 500))) * 1024 * 1024
    except Exception:
        max_size_bytes = DOC_TRANSFER_MAX_FILE_SIZE_BYTES

    oversized = []
    sendable = []
    for relative_path, fpath in candidates:
        try:
            fsize = os.path.getsize(fpath)
        except OSError:
            fsize = 0
        (oversized if fsize > max_size_bytes else sendable).append((relative_path, fpath, fsize))

    if oversized:
        names = ", ".join(f"{os.path.basename(p)} ({s / (1024 * 1024):.1f} MB)" for _, p, s in oversized)
        err = (f"{len(oversized)} file(s) exceed the {max_size_bytes // (1024 * 1024)} MB "
               f"doc-transfer limit and were not sent: {names}")
        write_audit_log("DOC-TRANSFER-REJECTED",
                        f"pid={pid} dest={cfg.get('name', '')} reason=too_large files={names}")
        if not sendable:
            set_fields(pid, last_doc_transfer_error=err)
            write_push_log(pid, "", err, final_status="FAILED", destination_name=cfg.get("name", ""),
                           files_attempted=len(candidates), files_sent=0, files_failed=len(candidates),
                           event_type="DOC-TRANSFER", push_job_id=push_job.get("started_at"))
            ui_event_queue.put(("toast", ("Document Transfer Blocked", err)))
            return
        # Otherwise keep going with the sendable subset below; the
        # oversized ones are folded into files_failed once transfer_time
        # is computed further down so the final tally isn't silently short.

    candidates = [(rp, fp) for rp, fp, _ in sendable]

    dest_name = cfg.get("name", "")
    host = cfg.get("ip") if cfg.get("doc_transfer_use_dicom_host", True) else cfg.get("doc_transfer_ip")
    ok, port_val = validate_port(cfg.get("doc_transfer_port"), "Document Transfer Port")

    if not ok or not host:
        err = "Document transfer misconfigured: missing host or invalid port"
        set_fields(pid, last_doc_transfer_error=err)
        write_push_log(pid, "", err, final_status="FAILED", destination_name=dest_name,
                       files_attempted=len(candidates), files_sent=0, files_failed=len(candidates),
                       event_type="DOC-TRANSFER", push_job_id=push_job.get("started_at"))
        write_audit_log("DOC-TRANSFER-FAILED", f"pid={pid} dest={dest_name}")
        return

    # A.5: read configurable chunk size / connect timeout from APP_SETTINGS,
    # falling back to the original constants if the settings are missing
    # (older settings file, or settings not yet migrated).
    _settings = load_app_settings()
    try:
        chunk_size = max(1024, int(_settings.get("doc_transfer_chunk_size_kb", 64)) * 1024)
    except Exception:
        chunk_size = DOC_TRANSFER_CHUNK_SIZE
    try:
        connect_timeout = max(1, int(_settings.get("doc_transfer_timeout_sec", DOC_TRANSFER_CONNECT_TIMEOUT_SEC)))
    except Exception:
        connect_timeout = DOC_TRANSFER_CONNECT_TIMEOUT_SEC

    started = time.time()
    files_sent = 0
    files_failed = 0
    last_error = ""
    transfer_id = _dt_progress_start(pid, dest_name, len(candidates),
                                     sum(s for _, _, s in sendable))
    overall_bytes_sent = 0
    cancel_event = doc_transfer_active_transfers[transfer_id]["cancel_event"]

    # A.1: up to 2 retries (3 attempts total) per file, short fixed backoff,
    # all on the same open connection -- only the send-one-file step retries,
    # never the handshake/connection itself. If the socket is dead (sendall
    # raises), that's a connection-level failure: no retry, fall through to
    # the outer except exactly as before.
    DOC_TRANSFER_MAX_FILE_ATTEMPTS = 3
    DOC_TRANSFER_RETRY_BACKOFF_SEC = 1

    cancelled = False
    try:
        # TLS plug-in point: wrap this socket with
        # build_ssl_context_for_client(load_tls_config(), port=port_val)
        # once document-socket TLS is implemented.
        with socket.create_connection((host, port_val), timeout=connect_timeout) as sock:
            sock.settimeout(connect_timeout)
            header = {
                "patient_id": pid,
                "file_count": len(candidates),
                "auth_key": cfg.get("doc_transfer_auth_key", ""),
            }
            header_bytes = json.dumps(header).encode("utf-8")
            sock.sendall(DOC_TRANSFER_MAGIC + bytes([DOC_TRANSFER_PROTO_VERSION]) +
                        len(header_bytes).to_bytes(2, "big") + header_bytes)

            reply = _doc_transfer_recv_exact(sock, 1)
            if reply[0] != 0x01:
                reason_len = int.from_bytes(_doc_transfer_recv_exact(sock, 2), "big")
                reason = _doc_transfer_recv_exact(sock, reason_len).decode("utf-8", errors="replace")
                raise ConnectionError(f"Receiver rejected connection: {reason}")

            for relative_path, fpath in candidates:
                file_ok = False
                connection_dead = False
                file_started = time.time()
                file_size = 0
                attempts_used = 0
                _dt_progress_update(transfer_id, current_filename=os.path.basename(relative_path))
                for attempt in range(1, DOC_TRANSFER_MAX_FILE_ATTEMPTS + 1):
                    attempts_used = attempt
                    try:
                        sock.sendall(bytes([0x01]))  # continuation: one more file header follows
                        size = os.path.getsize(fpath)
                        file_size = size
                        file_header = {"relative_path": relative_path, "size": size, "sha256": ""}
                        file_header_bytes = json.dumps(file_header).encode("utf-8")
                        sock.sendall(len(file_header_bytes).to_bytes(2, "big") + file_header_bytes)
                        with open(fpath, "rb") as f:
                            while True:
                                if cancel_event.is_set():
                                    raise DocTransferCancelled("Cancelled by user")
                                chunk = f.read(chunk_size)
                                if not chunk:
                                    break
                                sock.sendall(chunk)
                                bandwidth_limiter.throttle(len(chunk))
                                overall_bytes_sent += len(chunk)
                                elapsed = time.time() - started
                                _dt_progress_update(
                                    transfer_id, bytes_sent=overall_bytes_sent,
                                    speed_mbps=(overall_bytes_sent / (1024 * 1024)) / elapsed if elapsed > 0 else 0.0,
                                )

                        file_reply = _doc_transfer_recv_exact(sock, 1)
                        if file_reply[0] == 0x01:
                            file_ok = True
                            break
                        else:
                            reason_len = int.from_bytes(_doc_transfer_recv_exact(sock, 2), "big")
                            last_error = _doc_transfer_recv_exact(sock, reason_len).decode("utf-8", errors="replace")
                            if attempt < DOC_TRANSFER_MAX_FILE_ATTEMPTS:
                                time.sleep(DOC_TRANSFER_RETRY_BACKOFF_SEC)
                                continue
                    except DocTransferCancelled:
                        # Not a connection failure -- re-raise past the
                        # generic except below so the outer handler can
                        # give this its own 'Cancelled' status instead of
                        # 'FAILED'.
                        raise
                    except Exception as e:
                        # Socket-level failure (sendall/recv raised): the
                        # connection itself is dead, do not retry -- bubble
                        # up to the outer except like before.
                        last_error = str(e)
                        connection_dead = True
                        break

                if connection_dead:
                    raise ConnectionError(last_error)

                if file_ok:
                    files_sent += 1
                else:
                    files_failed += 1
                _dt_progress_update(transfer_id, files_done=files_sent + files_failed)

                # Part 3 Activity table: one row per file, distinct
                # event_type from the aggregate DOC-TRANSFER summary
                # below so compute_doc_transfer_stats() (which treats
                # each DOC-TRANSFER record as one whole-push attempt)
                # doesn't double-count against these per-file rows.
                file_duration = round(time.time() - file_started, 3)
                file_speed = round((file_size / (1024 * 1024)) / file_duration, 2) if file_duration > 0 and file_size > 0 else None
                if relative_path.startswith(REPORTS_SUBDIR + "/") and os.path.basename(relative_path) == os.path.basename(get_report_path(pid)):
                    file_transfer_type = "Report"
                elif relative_path.startswith(REPORTS_SUBDIR + "/") and os.path.basename(relative_path) == os.path.basename(get_history_path(pid)):
                    file_transfer_type = "History"
                elif relative_path.startswith(ATTACHMENTS_SUBDIR + "/"):
                    file_transfer_type = "Attachment"
                else:
                    file_transfer_type = "Other"
                write_push_log(
                    pid, "", "" if file_ok else last_error,
                    final_status="OK" if file_ok else "FAILED",
                    event_type="DOC-TRANSFER-FILE",
                    destination_name=dest_name,
                    filename=os.path.basename(relative_path),
                    file_size=file_size,
                    duration_sec=file_duration,
                    transfer_speed_mbps=file_speed,
                    retry_count=max(0, attempts_used - 1),
                    transfer_type=file_transfer_type,
                    authenticated=bool(cfg.get("doc_transfer_auth_key")),
                    push_job_id=push_job.get("started_at"),
                )

            sock.sendall(bytes([0x00]))  # continuation: pusher is done sending files
    except DocTransferCancelled:
        files_failed = len(candidates) - files_sent
        last_error = "Cancelled by user"
        cancelled = True
    except Exception as e:
        files_failed = len(candidates) - files_sent
        last_error = str(e)

    transfer_time_sec = round(time.time() - started, 3)
    total_attempted = len(candidates) + len(oversized)
    total_failed = files_failed + len(oversized)
    if oversized and not last_error:
        last_error = (f"{len(oversized)} file(s) exceeded the "
                      f"{max_size_bytes // (1024 * 1024)} MB doc-transfer limit")
    final_status = "CANCELLED" if cancelled else ("OK" if total_failed == 0 and files_sent > 0 else "FAILED")
    _dt_progress_finish(transfer_id, "cancelled" if cancelled else ("done" if final_status == "OK" else "failed"))

    if final_status != "OK":
        set_fields(pid, last_doc_transfer_error=last_error or "Unknown document transfer error")
    else:
        # C.7: "modified since last send" indicator depends on this
        # timestamp -- set alongside the existing last_doc_transfer_error
        # reset, purely additive.
        set_fields(pid, last_doc_transfer_error="",
                  doc_transfer_sent_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    write_push_log(
        pid, "", last_error if final_status != "OK" else "",
        final_status=final_status,
        destination_name=dest_name,
        files_attempted=total_attempted,
        files_sent=files_sent,
        files_failed=total_failed,
        transfer_time_sec=transfer_time_sec,
        event_type="DOC-TRANSFER",
        push_job_id=push_job.get("started_at"),
    )
    write_audit_log("DOC-TRANSFER-OK" if final_status == "OK" else "DOC-TRANSFER-FAILED",
                    f"pid={pid} dest={dest_name} status={final_status}")

    if final_status == "FAILED":
        # D.4: queue for retry of the document leg only -- never touches
        # the DICOM-only offline queue behavior for existing "dicom"-kind
        # (or legacy, kind-less) entries. A user-cancelled transfer is
        # NOT queued for retry -- that would silently override their
        # explicit choice to stop it.
        enqueue_offline(pid, dest_name, last_error or "Unknown document transfer error", kind="documents")
    else:
        dequeue_offline(pid, kind="documents")


def do_resend_patient_documents(pid):
    """A.7: 'Resend documents only' manual action. Looks up the
    last-used destination for this patient by parsing the worklist's
    existing push_target field ("AE@ip:port") back into a full destination
    profile, falling back to the default destination if none is recorded
    or the match fails. Calls push_patient_documents() directly --
    bypassing push_single_patient and the DICOM send loop entirely -- so
    it works even if the patient's DICOM status is already STATUS_SENT
    from a previous full push. Runs the actual transfer on a background
    thread since it performs real network I/O."""
    with data_lock:
        d = patient_data.get(pid, {})
        push_target = d.get("push_target", "")

    cfg = None
    if push_target and "@" in push_target and ":" in push_target:
        try:
            ae_part, hostport = push_target.split("@", 1)
            host_part, port_part = hostport.rsplit(":", 1)
            for dest in load_destinations():
                if (dest.get("ae") == ae_part and dest.get("ip") == host_part
                        and str(dest.get("port")) == port_part):
                    cfg = dest
                    break
        except Exception:
            cfg = None

    if cfg is None:
        cfg = get_default_destination()

    if not cfg:
        err = "No destination available to resend documents to"
        set_fields(pid, last_doc_transfer_error=err)
        write_push_log(pid, "", err, final_status="FAILED", files_attempted=0, files_sent=0, files_failed=0,
                       event_type="DOC-TRANSFER")
        write_audit_log("DOC-TRANSFER-FAILED", f"pid={pid} reason=no-destination")
        return False

    threading.Thread(target=push_patient_documents, args=(pid, cfg), daemon=True).start()
    return True


def is_pending_stale(d):
    """True if a worklist entry is still Pending/Failed and has been sitting

    for longer than STALE_PENDING_HOURS since it was received/imported."""
    if d.get("status") not in (STATUS_PENDING, STATUS_FAILED):
        return False
    t = d.get("time", "")
    if not t:
        return False
    try:
        received_at = datetime.datetime.strptime(t, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return False
    age_hours = (datetime.datetime.now() - received_at).total_seconds() / 3600.0
    return age_hours >= STALE_PENDING_HOURS

# =========================================================
# DISK SPACE MONITORING
# =========================================================

def get_free_disk_gb(path="."):
    try:
        usage = shutil.disk_usage(path)
        return usage.free / (1024 ** 3)
    except Exception:
        log_exception("Failed to check disk usage")
        return None


_low_disk_notify_throttle = [0.0]


def check_disk_space_and_warn():
    """Returns (low, free_gb). Posts a UI event if space is low so the GUI
    can show a one-time-per-check warning without blocking receiver threads."""
    free_gb = get_free_disk_gb(OUTPUT_DIR if os.path.isdir(OUTPUT_DIR) else ".")
    if free_gb is None:
        return False, None
    low = free_gb < LOW_DISK_WARNING_GB
    if low:
        ui_event_queue.put(("low_disk_warning", free_gb))
        now = time.time()
        if now - _low_disk_notify_throttle[0] > 300:  # once per 5 min, matches the popup throttle
            _low_disk_notify_throttle[0] = now
            notify_event("disk_space_low", "Disk Space Low", f"Only {free_gb:.2f} GB remaining.")
    return low, free_gb

# =========================================================
# IMPORT FOLDER (Receiver/Pusher tabs)
# =========================================================

def is_dicom_candidate(path):
    lower = path.lower()
    if lower.endswith(DICOM_EXTENSIONS):
        return True
    # extensionless files: peek for DICM magic at offset 128
    try:
        with open(path, "rb") as f:
            f.seek(128)
            return f.read(4) == b"DICM"
    except Exception:
        return False


def import_folder(folder, progress_cb=None):
    """Recursively scan a folder for DICOM files and ingest them into OUTPUT_DIR
    and the worklist, just like files received over the network."""
    files = []
    for root, _dirs, names in os.walk(folder):
        for name in names:
            full = os.path.join(root, name)
            if is_dicom_candidate(full):
                files.append(full)

    total = len(files)
    done = 0
    imported = 0
    failed = 0

    for path in files:
        try:
            try:
                bandwidth_limiter.throttle(os.path.getsize(path))
            except OSError:
                pass
            ds = pydicom.dcmread(path, force=True)

            pid = str(getattr(ds, "PatientID", "UNKNOWN")) or "UNKNOWN"
            pname = format_dicom_person_name(getattr(ds, "PatientName", ""))
            institution = str(getattr(ds, "InstitutionName", ""))
            modality = str(getattr(ds, "Modality", ""))
            study_uid = str(getattr(ds, "StudyInstanceUID", ""))
            sop_uid = str(getattr(
                ds, "SOPInstanceUID",
                datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
            ))
            sop_class_uid = str(getattr(ds, "SOPClassUID", ""))
            is_document = sop_class_uid in DOCUMENT_SOP_CLASS_UIDS

            dest_folder = get_patient_folder(pid)
            os.makedirs(dest_folder, exist_ok=True)
            dest_path = os.path.join(dest_folder, f"{sop_uid}.dcm")

            is_duplicate = os.path.exists(dest_path)
            if not is_duplicate:
                shutil.copy2(path, dest_path)

            # Same duplicate-safe accounting as handle_store(): re-importing
            # a SOP Instance UID that's already present must not inflate
            # the displayed image/document count or the daily stats.
            upsert_patient(pid, pname, institution, study_uid, modality,
                           source="Imported", is_document=is_document,
                           is_duplicate=is_duplicate)
            if not is_duplicate:
                record_receive_stat(study_uid, sop_class_uid)
            imported += 1

        except Exception as e:
            failed += 1
            write_receiver_log("", "", "", f"IMPORT FAILED ({path}): {e}")
            log_exception(f"Import failed for {path}")

        done += 1
        if progress_cb:
            progress_cb(done, total)

    write_audit_log("IMPORT", f"folder={folder} imported={imported} failed={failed} total={total}")
    check_disk_space_and_warn()
    ui_event_queue.put(("import_done", (imported, failed, total)))

# =========================================================
# DICOM C-STORE RECEIVER
# =========================================================

def handle_store(event):
    ds = None
    pid_for_log = pname_for_log = institution_for_log = ""
    receive_started = time.time()
    try:
        ds = event.dataset
        ds.file_meta = event.file_meta

        pid = str(getattr(ds, "PatientID", "UNKNOWN")) or "UNKNOWN"
        pname = format_dicom_person_name(getattr(ds, "PatientName", ""))
        institution = str(getattr(ds, "InstitutionName", ""))
        modality = str(getattr(ds, "Modality", ""))
        study_uid = str(getattr(ds, "StudyInstanceUID", ""))
        series_uid = str(getattr(ds, "SeriesInstanceUID", ""))
        sop_uid = str(getattr(
            ds, "SOPInstanceUID",
            datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
        ))
        sop_class_uid = str(getattr(ds.file_meta, "MediaStorageSOPClassUID", "") or getattr(ds, "SOPClassUID", ""))
        transfer_syntax = str(getattr(ds.file_meta, "TransferSyntaxUID", ""))
        pid_for_log, pname_for_log, institution_for_log = pid, pname, institution

        # Identify the calling/called AE Titles and the peer socket for the
        # Source column and enterprise diagnostics.
        try:
            calling_ae = event.assoc.requestor.ae_title
            calling_ae = calling_ae.decode().strip() if isinstance(calling_ae, bytes) else str(calling_ae).strip()
        except Exception:
            calling_ae = "UNKNOWN_AE"
        try:
            called_ae = event.assoc.acceptor.ae_title
            called_ae = called_ae.decode().strip() if isinstance(called_ae, bytes) else str(called_ae).strip()
        except Exception:
            called_ae = "UNKNOWN_AE"
        try:
            peer_addr = event.assoc.requestor.address_info
            source_ip = getattr(peer_addr, "address", "") or ""
            source_port = getattr(peer_addr, "port", "") or ""
        except Exception:
            source_ip, source_port = "", ""

        folder = get_patient_folder(pid)
        os.makedirs(folder, exist_ok=True)
        file_path = os.path.join(folder, f"{sop_uid}.dcm")

        is_duplicate = os.path.exists(file_path)
        ds.save_as(file_path, write_like_original=False)

        try:
            file_size = os.path.getsize(file_path)
        except Exception:
            file_size = None
        duration_sec = round(time.time() - receive_started, 3)

        is_document = sop_class_uid in DOCUMENT_SOP_CLASS_UIDS
        upsert_patient(pid, pname, institution, study_uid, modality,
                       source=calling_ae, status=STATUS_RECEIVED, is_document=is_document,
                       is_duplicate=is_duplicate)
        # A duplicate re-transmission of a SOP Instance we already have
        # isn't a "new" image/document arriving -- counting it here would
        # inflate the "Images Received Today" dashboard/report statistic
        # the same way it used to inflate the worklist "#" column.
        if not is_duplicate:
            record_receive_stat(study_uid, sop_class_uid)

        write_audit_log(
            "C-STORE",
            f"from_ae={calling_ae} patient_id={pid} sop_uid={sop_uid} "
            f"duplicate={is_duplicate}"
        )
        write_receiver_log(
            pid, pname, institution, "",
            result="SUCCESS",
            event_type="C-STORE",
            study_uid=study_uid,
            series_uid=series_uid,
            sop_instance_uid=sop_uid,
            sop_class=sop_class_uid,
            modality=modality,
            calling_ae=calling_ae,
            called_ae=called_ae,
            source_ip=source_ip,
            source_port=source_port,
            transfer_syntax=transfer_syntax,
            file_size=file_size,
            duration_sec=duration_sec,
            save_location=file_path,
            duplicate=is_duplicate,
        )

        low, free_gb = check_disk_space_and_warn()
        if low:
            app_logger.warning("Low disk space: %.2f GB free", free_gb or -1)

        if receiver_state.get("autoroute"):
            _schedule_autoroute_push(pid)

        ui_event_queue.put(("toast", ("Image Received", f"{pid} ({modality}) from {calling_ae}")))
        ui_event_queue.put(("recv_progress", (file_size or 0, calling_ae, pid, modality)))

        return 0x0000

    except Exception as e:
        write_receiver_log(
            pid_for_log, pname_for_log, institution_for_log, str(e),
            result="FAILURE",
            event_type="C-STORE",
            duration_sec=round(time.time() - receive_started, 3),
            error_code="0xC210",
            stack_trace=traceback.format_exc(),
        )
        log_exception("handle_store failed")
        write_audit_log("C-STORE-FAILED", str(e))
        return 0xC210


MAX_PRESENTATION_CONTEXTS = 128  # hard DICOM upper bound per association

DICOM_TLS_PORT = 2762  # IANA-registered "dicom-tls" port; a bare protocol-level
                        # abort on exactly this port with no TLS attempted is a
                        # strong signal the remote node requires DICOM-over-TLS.


def cap_sop_list(sops):
    """pynetdicom/DICOM allows at most 128 presentation contexts in a
    single association's A-ASSOCIATE-RQ -- this only actually applies to
    ae.add_requested_context() (the SCU/pusher side proposing a specific
    association), which raises ValueError past 128. ae.add_supported_
    context() (the SCP/receiver side's catalog of what it's willing to
    accept) has no such limit -- it's just a matching pool, verified
    directly against pynetdicom, and any single incoming association is
    still bounded to 128 negotiated contexts regardless of how large this
    pool is. Only call this before add_requested_context(); the receiver
    intentionally does NOT cap its supported-context list (see
    start_receiver_server()) so it can accept every SOP Class in
    sopclass.ini, however many there are."""
    if len(sops) > MAX_PRESENTATION_CONTEXTS:
        write_receiver_log(
            "", "", "",
            f"sopclass.ini lists {len(sops)} SOP Classes; only the first "
            f"{MAX_PRESENTATION_CONTEXTS} can be PROPOSED in a single push association. "
            f"(The receiver itself has no such limit and still accepts all of them.)"
        )
        return sops[:MAX_PRESENTATION_CONTEXTS]
    return sops


# =========================================================
# ASSOCIATION DIAGNOSTICS
# =========================================================
# pynetdicom already knows EXACTLY why a remote node rejected or aborted
# an association (bad Calling/Called AE title, unsupported application
# context, congestion, no acceptable presentation context, etc) -- it
# just logs that detail via its own "pynetdicom" logger and doesn't
# expose it on the Association object. Previously this app discarded
# that detail entirely and just reported "Association rejected/failed"
# for every kind of failure, which made it impossible to tell a config
# problem (e.g. wrong AE title) apart from a network problem. This
# section captures pynetdicom's own diagnostic log lines for the
# duration of a single associate() call and turns them into an
# actionable message.

AE_TITLE_MAX_LEN = 16


def validate_ae_title(title, field_label="AE Title", allow_empty=False):
    """Validate a DICOM AE Title before it ever reaches pynetdicom.
    Returns (ok: bool, cleaned_value_or_error_message: str).
    Catching this here gives a clear, specific error instead of a raw
    pynetdicom ValueError surfacing later at association time.

    §fix: coerce to str BEFORE the emptiness check. `(title or "").strip()`
    only guards against None/"" -- any other truthy non-str value (an int,
    for instance) sails through the `or` unchanged and .strip() then raises
    AttributeError: 'int' object has no attribute 'strip'. Several callers
    pass through values pulled straight from a saved destination dict
    without guaranteeing they're already strings."""
    title = str(title).strip() if title is not None else ""
    if not title:
        if allow_empty:
            return True, ""
        return False, f"{field_label} cannot be empty."
    if len(title) > AE_TITLE_MAX_LEN:
        return False, (f"{field_label} '{title}' is {len(title)} characters long; "
                        f"DICOM AE Titles must be {AE_TITLE_MAX_LEN} characters or fewer.")
    if any(ch in title for ch in (" ", "\t", "\\")):
        return False, (f"{field_label} '{title}' contains spaces or backslashes, which "
                        f"most DICOM SCPs will reject. Use letters, digits, '_' or '-' only.")
    return True, title


def validate_port(port_str, field_label="Port"):
    """Validate a TCP port string (or int). Returns (ok, int_value_or_error_message).

    §fix: same class of bug as validate_ae_title above -- several call
    sites pass `int(dest["port"])` straight through (e.g. dicom_echo()
    calls at the destination-list "Test Connection" buttons), so this must
    accept ints too rather than assuming a string and crashing on
    .strip()."""
    port_str = str(port_str).strip() if port_str is not None else ""
    if not port_str.isdigit():
        return False, f"{field_label} must be numeric."
    port = int(port_str)
    if not (1 <= port <= 65535):
        return False, f"{field_label} must be between 1 and 65535 (got {port})."
    return True, port


class _PynetdicomDiagnosticCapture(logging.Handler):
    """Temporarily attached to the 'pynetdicom' logger around a single
    associate() call so we can recover the exact Result/Source/Reason
    that library already computed for a rejection or abort."""

    def __init__(self):
        super().__init__(level=logging.INFO)
        self.lines = []

    def emit(self, record):
        try:
            self.lines.append(record.getMessage())
        except Exception:
            pass


def _diagnose_captured_lines(lines):
    """Turn captured pynetdicom log lines into a structured diagnosis."""
    joined = " | ".join(lines)
    result_str = reason_str = ""
    for line in lines:
        if line.startswith("Result:"):
            # e.g. "Result: Rejected Permanent, Source: Service User"
            result_str = line
        elif line.startswith("Reason:"):
            reason_str = line.split("Reason:", 1)[1].strip()
    permanent = "Rejected Permanent" in joined
    no_acceptable_context = "No accepted presentation contexts" in joined
    return {
        "reason": reason_str,
        "result_line": result_str,
        "permanent": permanent,
        "no_acceptable_context": no_acceptable_context,
        "raw": joined,
    }


def associate_with_diagnostics(ae, addr, port, remote_ae_title, **assoc_kwargs):
    """Wraps ae.associate() and returns (assoc, diagnosis) where
    diagnosis is a dict describing exactly what happened:
        ok: bool                     -- association established
        retryable: bool              -- worth retrying / retrying later
        message: str                 -- human-readable, actionable detail
    This is what lets the app tell the difference between "wrong AE
    Title configured for this destination" (fix the config, retrying
    won't help) and "server was briefly busy" (retry is reasonable)."""
    dest_preview = f"{remote_ae_title}@{addr}:{port}"

    # Fail fast rather than waiting on a network round trip we already
    # know will fail: port 2762 is the IANA/DICOM-registered 'dicom-tls'
    # port, reserved specifically for DICOM-over-TLS. If we're about to
    # connect there without ever having built a TLS context, every single
    # attempt will end in the same unreasoned protocol-level abort once
    # the remote's TLS listener receives our plaintext A-ASSOCIATE-RQ and
    # drops it. No amount of retrying changes that outcome, so surface the
    # real, fixable problem immediately instead of burning the acse/dimse
    # timeout and the retry/backoff schedule on a foregone conclusion.
    if int(port) == DICOM_TLS_PORT and "tls_args" not in assoc_kwargs:
        return None, {
            "ok": False, "retryable": False,
            "message": (f"Refusing to attempt a plaintext connection to {dest_preview} — "
                        f"port {DICOM_TLS_PORT} is the IANA/DICOM-registered 'dicom-tls' "
                        f"port, reserved for DICOM-over-TLS. Enable TLS for this "
                        f"destination (Settings → TLS: set 'enabled' to true and supply "
                        f"ca_cert/cert/key) before retrying, or confirm with the remote "
                        f"PACS administrator that this destination truly expects plaintext "
                        f"DICOM on a non-standard port."),
        }

    capture = _PynetdicomDiagnosticCapture()
    pynd_logger = logging.getLogger("pynetdicom")
    prev_level = pynd_logger.level
    if pynd_logger.level == logging.NOTSET or pynd_logger.level > logging.INFO:
        pynd_logger.setLevel(logging.INFO)
    pynd_logger.addHandler(capture)
    try:
        try:
            assoc = ae.associate(addr, port, ae_title=remote_ae_title, **assoc_kwargs)
        except ValueError as e:
            # e.g. an AE title that's too long/blank slipped through, or a
            # bad host/port -- surface it directly rather than as a
            # mysterious connection failure.
            return None, {
                "ok": False, "retryable": False,
                "message": f"Configuration error: {e}",
            }
        except (ConnectionRefusedError, TimeoutError, OSError) as e:
            return None, {
                "ok": False, "retryable": True,
                "message": (f"Could not reach {remote_ae_title}@{addr}:{port} — {e}. "
                            f"Check the IP/hostname, port, firewall rules, and that the "
                            f"remote DICOM service is actually running."),
            }
    finally:
        pynd_logger.removeHandler(capture)
        pynd_logger.setLevel(prev_level)

    diag = _diagnose_captured_lines(capture.lines)

    if assoc.is_established:
        # §fix: log exactly which transfer syntax the peer accepted for each
        # SOP Class right when the association comes up. Compared against
        # native_ts_by_sop (built in push_single_patient), this turns "is
        # the peer respecting our proposal order?" into a one-line fact
        # instead of a guess -- if a file's native TS shows up here as
        # accepted, it goes over the wire untouched; if only a baseline TS
        # (Implicit/Explicit VR LE) shows up despite JPEG2000 being offered,
        # the peer is choosing to downgrade regardless of order, and the
        # in-memory transcode fallback in _send_one_file() is doing its job.
        try:
            for cx in assoc.accepted_contexts:
                accepted_ts = str(cx.transfer_syntax[0]) if cx.transfer_syntax else "NONE"
                app_logger.info("Context accepted: SOP=%s -> TS=%s", cx.abstract_syntax, accepted_ts)
        except Exception:
            pass
        return assoc, {"ok": True, "retryable": True, "message": "Association established."}

    dest = f"{remote_ae_title}@{addr}:{port}"

    if assoc.is_rejected:
        reason = diag["reason"] or "no reason given by the remote node"
        hint = ""
        low = reason.lower()
        if "calling ae title" in low:
            hint = (" — the destination has an AE-title whitelist and does not recognise "
                    "OUR Calling AE Title. Check the 'Calling AE Title' field for this "
                    "destination against what the remote PACS administrator configured "
                    "(it's case-sensitive).")
        elif "called ae title" in low:
            hint = (" — the 'Remote AE Title' you entered does not match the AE title the "
                    "destination is actually listening as. Verify the exact, case-sensitive "
                    "AE title with the remote PACS administrator.")
        elif "application context" in low:
            hint = (" — the remote host/port may not be a DICOM Storage SCP at all "
                    "(wrong port, or a non-DICOM service answering on it).")
        elif "congestion" in low or "local limit" in low:
            hint = " — the remote node is busy or at its connection limit; retrying later may help."
        message = f"Association to {dest} was REJECTED: {reason}{hint}"
        return assoc, {"ok": False, "retryable": not diag["permanent"], "message": message}

    if assoc.is_aborted:
        if diag["no_acceptable_context"]:
            proposed = []
            try:
                for ctx in ae.requested_contexts:
                    sop_uid = ctx.abstract_syntax
                    name = getattr(sop_uid, "name", None) or str(sop_uid)
                    proposed.append(f"{name} ({sop_uid})")
            except Exception:
                pass
            proposed_str = "; ".join(proposed) if proposed else "unknown (could not enumerate)"
            message = (f"Association to {dest} was ABORTED: the remote node's A-ASSOCIATE-AC "
                       f"accepted NONE of the proposed presentation contexts. We proposed: "
                       f"{proposed_str}. This means either (a) the destination doesn't "
                       f"support any transfer syntax we offered for these SOP Classes -- "
                       f"fixable by adding the missing Transfer Syntax to sopclass.ini via "
                       f"the in-GUI SOP editor -- or (b) the destination simply doesn't "
                       f"support these SOP Classes at all, which sopclass.ini cannot fix; "
                       f"check the destination's DICOM conformance statement (or ask its "
                       f"administrator) for its supported SOP Class list.")
        else:
            hint = ""
            if int(port) == DICOM_TLS_PORT and "tls_args" not in assoc_kwargs:
                hint = (f" This is the IANA/DICOM-registered port for 'dicom-tls' "
                        f"(DICOM Upper Layer Protocol over TLS). A bare, unreasoned "
                        f"protocol-level abort on this exact port almost always means "
                        f"the remote node is waiting for a TLS handshake and we just "
                        f"sent it a plaintext DICOM association request instead, which "
                        f"it can't parse and simply drops. Enable TLS for this "
                        f"connection (Settings → TLS) and supply the appropriate "
                        f"CA/cert/key, then retry.")
            message = (f"Association to {dest} was ABORTED by the remote node "
                       f"(protocol-level abort, no further reason given).{hint}")
        return assoc, {"ok": False, "retryable": True, "message": message}

    return assoc, {
        "ok": False, "retryable": True,
        "message": f"Association to {dest} failed for an unknown reason (no rejection or abort observed).",
    }


def start_receiver_server(ae_title, port):
    ae = AE(ae_title=ae_title)

    sops = load_extra_sops()  # NOT capped -- see cap_sop_list()'s docstring:
                               # add_supported_context() has no 128 limit,
                               # only add_requested_context() (the pusher's
                               # push_single_patient()) does.
    # "Use all transfer syntaxes": accept anything in sopclass.ini PLUS the
    # full pydicom-known transfer syntax catalog (ALL_TRANSFER_SYNTAXES), so
    # the receiver never rejects a sender purely because its transfer syntax
    # wasn't explicitly listed in sopclass.ini.
    transfer_syntaxes = list(dict.fromkeys(load_transfer_syntaxes() + ALL_TRANSFER_SYNTAXES))

    # Verification (C-ECHO) is a baseline capability every DICOM SCP is
    # expected to support, independent of whatever Storage SOP Classes
    # happen to be configured in sopclass.ini -- an admin editing the
    # Storage SOP Class list should never be able to accidentally make
    # this receiver unpingable. VERIFICATION_SOP_CLASS is defined further
    # down in this file (next to the pusher's echo-test code); Python
    # resolves that at call time, so the forward reference is safe here.
    try:
        ae.add_supported_context(VERIFICATION_SOP_CLASS, transfer_syntaxes)
    except Exception:
        log_exception("Failed to add supported context for Verification SOP Class")

    for sop in sops:
        try:
            ae.add_supported_context(sop, transfer_syntaxes)
        except Exception:
            log_exception(f"Failed to add supported context for {sop}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    receiver_state["server_ae"] = ae
    receiver_state["running"] = True
    ui_event_queue.put(("receiver_status", True))
    write_audit_log("RECEIVER_START", f"ae_title={ae_title} port={port}")
    notify_event("receiver_started", "Receiver Started", f"AE={ae_title} port={port}")

    # Document transfer: only opens a new TCP port when the global
    # doc_transfer_enabled setting is on. Off by default -- a destination
    # saved before this feature, or an admin who never enables it, gets
    # zero new listening ports and zero behavior change.
    if APP_SETTINGS.get("doc_transfer_enabled"):
        doc_stop_event = threading.Event()
        doc_port = APP_SETTINGS.get("doc_transfer_receiver_port")
        doc_thread = threading.Thread(
            target=doc_transfer_accept_loop, args=(doc_port, doc_stop_event), daemon=True)
        receiver_state["doc_server"] = {"thread": doc_thread, "stop_event": doc_stop_event}
        doc_thread.start()
        ui_event_queue.put(("doc_transfer_status", (True, doc_port)))
    else:
        receiver_state["doc_server"] = None
        ui_event_queue.put(("doc_transfer_status", (False, None)))

    tls_cfg = load_tls_config()
    ssl_context = build_ssl_context_for_server(tls_cfg)

    try:
        kwargs = dict(
            evt_handlers=[(evt.EVT_C_STORE, handle_store)],
            block=True,
        )
        if ssl_context is not None:
            kwargs["ssl_context"] = ssl_context
        ae.start_server(("0.0.0.0", int(port)), **kwargs)
    except Exception as e:
        log_exception("Receiver server crashed")
        write_audit_log("RECEIVER_CRASH", str(e))
        notify_event("receiver_stopped", "Receiver Stopped", f"Receiver crashed: {e}")
        ui_event_queue.put(("toast", ("Receiver Stopped", f"Receiver crashed: {e}")))
    finally:
        receiver_state["running"] = False
        ui_event_queue.put(("receiver_status", False))
        write_audit_log("RECEIVER_STOP", f"ae_title={ae_title} port={port}")
        notify_event("receiver_stopped", "Receiver Stopped", f"AE={ae_title} port={port}")
        doc_server = receiver_state.get("doc_server")
        if doc_server:
            try:
                doc_server["stop_event"].set()
            except Exception:
                log_exception("Error stopping document-transfer server")
            receiver_state["doc_server"] = None
            ui_event_queue.put(("doc_transfer_status", (False, None)))


def stop_receiver_server():
    ae = receiver_state.get("server_ae")
    if ae:
        try:
            ae.shutdown()
        except Exception:
            log_exception("Error shutting down receiver AE")
    receiver_state["running"] = False
    ui_event_queue.put(("receiver_status", False))

    doc_server = receiver_state.get("doc_server")
    if doc_server:
        try:
            doc_server["stop_event"].set()
        except Exception:
            log_exception("Error stopping document-transfer server")
        receiver_state["doc_server"] = None
        ui_event_queue.put(("doc_transfer_status", (False, None)))

# =========================================================
# DICOM C-STORE PUSHER
# =========================================================

VERIFICATION_SOP_CLASS = UID("1.2.840.10008.1.1")  # Verification SOP Class (C-ECHO)


def probe_tls_handshake(ip, port, timeout=5):
    """Attempt a bare TLS handshake against ip:port -- no DICOM involved at
    all, no ca_cert/cert/key required. This exists specifically for sites
    where the remote admin can't be reached for certs: it tells you
    definitively whether the problem is 'TLS handshake never completes'
    (network/firewall/TLS-version/SNI issue -- nothing to do with DICOM or
    sopclass.ini) versus 'TLS is fine, something DICOM-specific is wrong'.
    On success it also surfaces the remote's actual certificate subject,
    issuer, and validity dates, which is often enough on its own to decide
    whether to trust it -- without needing anyone to hand you a ca_cert
    file. Returns (ok: bool, message: str)."""
    import socket
    import ssl
    try:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        with socket.create_connection((ip, int(port)), timeout=timeout) as sock:
            with ctx.wrap_socket(sock, server_hostname=ip) as tls_sock:
                cert_bin = tls_sock.getpeercert(binary_form=True)
                version = tls_sock.version()
                cipher = tls_sock.cipher()
        detail = f"TLS handshake succeeded (protocol={version}, cipher={cipher[0] if cipher else '?'})."
        if cert_bin:
            try:
                import datetime
                from cryptography import x509
                from cryptography.hazmat.backends import default_backend
                cert = x509.load_der_x509_certificate(cert_bin, default_backend())
                detail += (f" Remote certificate: subject={cert.subject.rfc4514_string()}, "
                           f"issuer={cert.issuer.rfc4514_string()}, "
                           f"valid {cert.not_valid_before_utc:%Y-%m-%d} to "
                           f"{cert.not_valid_after_utc:%Y-%m-%d}.")
            except Exception:
                detail += (" (Certificate details unavailable -- install the 'cryptography' "
                           "package for subject/issuer/validity parsing.)")
        return True, detail
    except ssl.SSLError as e:
        return False, (f"TLS handshake FAILED against {ip}:{port} — {e}. This is a pure "
                        f"TLS-layer failure with no DICOM involved: check TLS version "
                        f"support, whether the remote actually terminates TLS at this "
                        f"port, or whether it requires a client certificate (mutual TLS) "
                        f"to even begin the handshake — some servers abort before sending "
                        f"their own certificate if they don't see one from the client.")
    except (ConnectionRefusedError, TimeoutError, OSError) as e:
        return False, (f"Could not even open a TCP connection to {ip}:{port} — {e}. "
                        f"This is below the TLS layer entirely: check IP/port/firewall "
                        f"before worrying about certs at all.")


def dicom_echo(remote_ae, remote_ip, remote_port, timeout=5, calling_ae=None, dest_label=None):
    """Perform a DICOM C-ECHO against the configured remote node.
    Returns (ok: bool, message: str).
    `calling_ae` sets our Calling AE Title for servers with AE whitelists.
    IMPORTANT: the blank-field fallback here must stay identical to the one
    used by the real push (DEFAULT_PUSH_CALLING_AE), otherwise this "test
    connection" can succeed under an identity the remote PACS's AE-title
    whitelist recognises while the actual push -- using a different
    identity -- gets its association rejected."""
    try:
        ok, remote_ae_clean = validate_ae_title(remote_ae, "Remote AE Title")
        if not ok:
            return False, remote_ae_clean
        ok, calling_clean = validate_ae_title(calling_ae or DEFAULT_PUSH_CALLING_AE, "Calling AE Title")
        if not ok:
            return False, calling_clean
        ok, port_val = validate_port(remote_port)
        if not ok:
            return False, port_val

        ae = AE(ae_title=calling_clean)
        ae.add_requested_context(VERIFICATION_SOP_CLASS)
        ae.acse_timeout = timeout
        ae.dimse_timeout = timeout
        ae.network_timeout = timeout

        tls_cfg = load_tls_config()
        ssl_context = build_ssl_context_for_client(tls_cfg, port=port_val, dest_label=dest_label or remote_ae_clean)

        assoc_kwargs = {}
        if ssl_context is not None:
            assoc_kwargs["tls_args"] = (ssl_context, None)
            tls_ok, tls_detail = probe_tls_handshake(remote_ip, port_val, timeout=timeout)
            write_audit_log("TLS-PROBE-OK" if tls_ok else "TLS-PROBE-FAILED",
                             f"{remote_ip}:{port_val} {tls_detail}")
            if not tls_ok:
                notify_event("tls_errors", "TLS Error", f"{remote_ip}:{port_val} — {tls_detail}")
                return False, (f"TLS layer check failed before attempting DICOM negotiation: "
                                f"{tls_detail}")

        assoc, diag = associate_with_diagnostics(ae, remote_ip, port_val, remote_ae_clean, **assoc_kwargs)
        if not diag["ok"]:
            write_audit_log("C-ECHO-FAILED", f"{remote_ae_clean}@{remote_ip}:{port_val} {diag['message']}")
            return False, diag["message"]

        status = assoc.send_c_echo()
        assoc.release()

        if status and getattr(status, "Status", None) == 0x0000:
            write_audit_log("C-ECHO-OK", f"{remote_ae_clean}@{remote_ip}:{port_val}")
            return True, f"C-ECHO succeeded against {remote_ae_clean}@{remote_ip}:{port_val}."
        write_audit_log("C-ECHO-FAILED", f"{remote_ae_clean}@{remote_ip}:{port_val} status={status}")
        return False, f"C-ECHO returned non-success status: {status}"

    except Exception as e:
        log_exception("C-ECHO failed")
        write_audit_log("C-ECHO-FAILED", f"{remote_ae}@{remote_ip}:{remote_port} error={e}")
        return False, f"C-ECHO failed: {e}"


# =========================================================
# TRANSFER SYNTAX TRANSCODING (§fix: transfer-syntax overhaul)
# =========================================================
# Preference order used when the peer's accepted transfer syntax for a SOP
# Class differs from the file's native encoding, so we must pick ONE of
# the transfer syntaxes the peer actually accepted. Uncompressed syntaxes
# come first (cheapest, no codec needed, always exactly representable),
# then lossless-compressed, then (last resort, and gated by
# allow_lossy_transcode -- see _pick_target_ts()) lossy-compressed. This
# replaces the previous `next(iter(accepted_ts))`, which picked an
# unordered set element and could silently choose a lossy transfer syntax
# over an available lossless one.
_TS_TARGET_PREFERENCE = [
    "1.2.840.10008.1.2.1",      # Explicit VR Little Endian
    "1.2.840.10008.1.2",        # Implicit VR Little Endian
    "1.2.840.10008.1.2.1.99",   # Deflated Explicit VR Little Endian
    "1.2.840.10008.1.2.2",      # Explicit VR Big Endian (retired, but harmless if offered)
    "1.2.840.10008.1.2.4.90",   # JPEG 2000 Lossless
    "1.2.840.10008.1.2.4.70",   # JPEG Lossless, Non-Hierarchical (Process 14, SV1)
    "1.2.840.10008.1.2.5",      # RLE Lossless
    "1.2.840.10008.1.2.4.201",  # High-Throughput JPEG 2000 Lossless
    "1.2.840.10008.1.2.4.202",  # High-Throughput JPEG 2000 Lossless RPCL
]

# Lossy transfer syntaxes and the DICOM-standard code for
# LossyImageCompressionMethod (PS3.3 C.7.6.16) each one implies, used when
# we actually have to lossy-recompress on the way out. NOTE: this is a
# best-effort mapping for the *method code tag* only -- it is intentionally
# NOT used to decide whether a transfer syntax is lossy (see
# _is_lossy_ts() below). A hand-maintained lossy allowlist misses any
# transfer syntax someone forgets to add, and ALL_TRANSFER_SYNTAXES (used
# as the fallback proposal list) includes ~20 lossy/video transfer syntaxes
# not listed here (JPEG-LS Near-Lossless, non-lossless HTJ2K, every
# MPEG2/MPEG-4/HEVC video syntax) -- if this dict were used as the safety
# gate, allow_lossy_transcode could be silently bypassed for any of them.
LOSSY_TRANSFER_SYNTAXES = {
    "1.2.840.10008.1.2.4.50": "ISO_10918_1",   # JPEG Baseline (Process 1)
    "1.2.840.10008.1.2.4.51": "ISO_10918_1",   # JPEG Extended (Process 2 & 4)
    "1.2.840.10008.1.2.4.91": "ISO_15444_1",   # JPEG 2000
    "1.2.840.10008.1.2.4.203": "ISO_18181_1",  # High-Throughput JPEG 2000
}

# §fix: robust lossy-detection used for the actual allow_lossy_transcode
# safety gate. Rather than hand-maintain a second allowlist that's just as
# likely to miss entries as the one above, this is: compressed AND not one
# of pydicom's own named *Lossless* transfer syntax constants. New/exotic
# transfer syntax UIDs (future DICOM supplements) default to "treat as
# lossy" -- the safe direction to fail in -- rather than defaulting to
# "assume safe" like a missing allowlist entry would.
_NAMED_LOSSLESS_TS = {
    str(_uid_const) for _uid_const in (
        _pydicom_uid.JPEGLossless, _pydicom_uid.JPEGLosslessSV1,
        _pydicom_uid.JPEGLSLossless, _pydicom_uid.JPEG2000Lossless,
        _pydicom_uid.JPEG2000MCLossless, _pydicom_uid.RLELossless,
        _pydicom_uid.HTJ2KLossless, _pydicom_uid.HTJ2KLosslessRPCL,
    )
}


def _is_lossy_ts(ts_uid):
    """True if `ts_uid` (a pydicom UID) represents lossy-compressed pixel
    data. Video transfer syntaxes (MPEG/HEVC) and any other compressed,
    non-named-lossless syntax are treated as lossy since pydicom's
    compress()/decompress() pipeline can't losslessly round-trip them
    through this app's re-encode path anyway."""
    uid = UID(str(ts_uid))
    return uid.is_compressed and str(uid) not in _NAMED_LOSSLESS_TS


def _pick_target_ts(accepted_ts, current_ts):
    """Choose which peer-accepted transfer syntax to send a file as.
    Prefers keeping the current/native TS if the peer accepts it, then
    walks _TS_TARGET_PREFERENCE, then any other lossless TS, and only
    falls back to a lossy TS -- deterministically, sorted -- as an
    absolute last resort."""
    if current_ts in accepted_ts:
        return current_ts
    for candidate in _TS_TARGET_PREFERENCE:
        if candidate in accepted_ts:
            return candidate
    lossless_remaining = {ts for ts in accepted_ts if not _is_lossy_ts(UID(ts))}
    if lossless_remaining:
        return sorted(lossless_remaining)[0]
    return sorted(accepted_ts)[0] if accepted_ts else None


def _transcode_dataset_for_target(ds, target_ts):
    """Re-encode `ds` IN PLACE so its pixel data actually matches
    `target_ts` -- not just its file_meta.TransferSyntaxUID label.
    Handles every combination:
        compressed   -> uncompressed : ds.decompress()
        uncompressed -> compressed   : ds.compress(target_ts)
        compressed A -> compressed B : decompress() then compress(B)
        uncompressed -> uncompressed : just flip VR/endianness flags
    Raises RuntimeError with an actionable message on failure (missing
    codec plugin, lossy transcode not permitted, etc.) instead of ever
    silently sending pixel data that doesn't match the declared transfer
    syntax -- which is what this app did previously: it would set
    file_meta.TransferSyntaxUID to whatever the peer accepted WITHOUT
    ever compressing into it, producing a file that *claims* e.g. JPEG
    2000 encoding while actually containing raw/uncompressed pixel
    bytes. That silently corrupts the image for any peer that trusts
    the declared transfer syntax (i.e. every conformant DICOM SCP)."""
    target_uid = UID(target_ts)
    source_uid = UID(str(ds.file_meta.TransferSyntaxUID))

    if source_uid == target_uid:
        return  # nothing to do

    if source_uid.is_compressed:
        if not DECODE_CODECS_AVAILABLE:
            raise RuntimeError(
                f"file is encoded as {source_uid.name} and no pixel-data decoder is "
                f"installed to read it. Install pylibjpeg + pylibjpeg-libjpeg/"
                f"pylibjpeg-openjpeg/pylibjpeg-rle, or python-gdcm."
            )
        try:
            ds.decompress()
        except Exception as e:
            raise RuntimeError(
                f"could not decompress source pixel data ({source_uid.name}): {e}"
            ) from e

    if target_uid.is_compressed:
        if _is_lossy_ts(target_uid) and not APP_SETTINGS.get("allow_lossy_transcode", False):
            raise RuntimeError(
                f"peer only accepted the LOSSY transfer syntax {target_uid.name} for this "
                f"SOP Class (no lossless or uncompressed option was accepted). Sending "
                f"would require re-encoding this image lossy, which is disabled by default "
                f"-- enable 'Allow lossy transcode' in Settings if this is acceptable for "
                f"your workflow, or configure the destination to accept an uncompressed / "
                f"lossless transfer syntax instead."
            )
        if not ENCODE_CODECS_AVAILABLE:
            raise RuntimeError(
                f"peer only accepted {target_uid.name} for this SOP Class, but no pixel-data "
                f"ENCODER is installed to compress into it (python-gdcm can decode but not "
                f"encode). Install: pip install pylibjpeg pylibjpeg-libjpeg "
                f"pylibjpeg-openjpeg pylibjpeg-rle"
            )
        try:
            # §fix: generate_instance_uid=False -- ds.compress() defaults to
            # TRUE, silently minting a brand new SOPInstanceUID every time a
            # file gets lossy- or lossless-recompressed on the way out. The
            # file actually on disk, this app's push logs, and any
            # dedup/study-tracking keyed by SOPInstanceUID all still refer
            # to the ORIGINAL UID -- so the peer would receive and store an
            # instance under a UID that matches nothing else in the system,
            # breaking traceability between what was sent and what's on
            # disk, with no error or warning.
            ds.compress(target_uid, generate_instance_uid=False)
        except Exception as e:
            raise RuntimeError(f"could not compress pixel data into {target_uid.name}: {e}") from e
        if _is_lossy_ts(target_uid):
            # PS3.3 C.7.6.16: any instance whose pixel data has been through
            # lossy compression MUST declare it, even if it arrived losslessly,
            # and regardless of whether we know the specific method code below.
            ds.LossyImageCompression = "01"
            method = LOSSY_TRANSFER_SYNTAXES.get(target_ts)
            if method:
                existing = list(getattr(ds, "LossyImageCompressionMethod", []) or [])
                if method not in existing:
                    ds.LossyImageCompressionMethod = existing + [method]
    else:
        ds.file_meta.TransferSyntaxUID = target_uid
        ds.is_little_endian = target_uid.is_little_endian
        ds.is_implicit_VR = target_uid.is_implicit_VR
        # §fix (crash + silent-ignore bug): the two lines above alone are
        # NOT enough. pydicom.Dataset tracks the encoding a dataset was
        # ORIGINALLY read from disk with in private _read_implicit /
        # _read_little attributes (exposed via ds.original_encoding), and
        # that recorded value does not change just because we set
        # is_implicit_VR/is_little_endian ourselves. pynetdicom's
        # Association.send_c_store() uses ds.original_encoding -- NOT
        # ds.is_implicit_VR/is_little_endian -- to decide what encoding the
        # dataset is "really" in. Left as-is, sending a file whose target
        # transfer syntax differs from its on-disk encoding either raises
        # "'dataset' is encoded as ... but the file meta has a Transfer
        # Syntax UID of ..." or, for two specific encoding combinations,
        # silently ignores our target_ts entirely and reverts to Implicit
        # VR Little Endian / Explicit VR Big Endian -- meaning this branch
        # was previously a no-op (or a hard crash, see below) on every
        # uncompressed<->uncompressed re-encode, which is the single most
        # common transcode case (e.g. Implicit VR LE file, peer only
        # accepted Explicit VR LE). Clearing these tells pydicom "treat this
        # as freshly constructed" so pynetdicom trusts our explicit
        # is_implicit_VR/is_little_endian instead.
        ds._read_implicit = None
        ds._read_little = None
        # (This branch also previously called the non-existent
        # `target_uid.is_big_endian` -- UID only exposes is_little_endian --
        # which raised AttributeError on every single call into this branch,
        # i.e. on every push where transcoding between two uncompressed
        # transfer syntaxes was needed. That crash, uncaught by the
        # RuntimeError handler in _send_one_file(), is why sends were
        # failing outright rather than falling back cleanly.)


def _accepted_ts_for_sop(assoc, sop_class_uid):
    """Return the set of transfer syntax UIDs the peer actually accepted
    for a given SOP Class on this association (empty set if that SOP
    Class wasn't accepted at all)."""
    accepted = set()
    try:
        for cx in assoc.accepted_contexts:
            if str(cx.abstract_syntax) == str(sop_class_uid):
                if cx.transfer_syntax:
                    accepted.add(str(cx.transfer_syntax[0]))
    except Exception:
        pass
    return accepted


def _send_one_file(assoc, fpath, anonymize, pid):
    """Send a single file over an existing association. Optionally
    anonymizes the dataset in memory before sending (does NOT modify the
    file on disk). Returns (ok: bool, error_message: str).

    §fix: the peer only ever accepts ONE transfer syntax per SOP Class
    context. If the file on disk happens to be encoded in a transfer
    syntax the peer did NOT accept for that SOP Class (e.g. the file is
    JPEG2000 but the peer only accepted Implicit VR LE for that SOP
    Class), pynetdicom's send_c_store() raises
    'No presentation context ... has been accepted by the peer with
    <TS> for the SCU role' and the file is dropped -- even though a
    context for that SOP Class WAS negotiated, just with a different
    transfer syntax. Rather than fail, decompress the pixel data in
    memory and re-encode the dataset using whatever transfer syntax the
    peer DID accept, then send that instead. This never touches the
    file on disk -- only the in-memory copy being sent."""
    try:
        try:
            bandwidth_limiter.throttle(os.path.getsize(fpath))
        except OSError:
            pass
        ds = pydicom.dcmread(fpath, force=True)
        if anonymize:
            ds = anonymize_dataset(ds, pid)

        sop_class_uid = str(getattr(ds, "SOPClassUID", "") or getattr(ds.file_meta, "MediaStorageSOPClassUID", ""))
        current_ts = str(getattr(ds.file_meta, "TransferSyntaxUID", ""))
        accepted_ts = _accepted_ts_for_sop(assoc, sop_class_uid)

        if accepted_ts and current_ts not in accepted_ts:
            target_ts = _pick_target_ts(accepted_ts, current_ts)
            if target_ts is None:
                return False, f"{fpath}: peer accepted this SOP Class with no usable transfer syntax."
            try:
                _transcode_dataset_for_target(ds, target_ts)
            except RuntimeError as transcode_err:
                return False, (f"{fpath}: peer only accepted {sorted(accepted_ts)} for this SOP Class "
                                f"(file is encoded as {current_ts}) and could not be transcoded to "
                                f"{target_ts}: {transcode_err}")

        status = assoc.send_c_store(ds)
        if status and getattr(status, "Status", 1) == 0x0000:
            return True, ""
        return False, f"C-STORE non-success status for {fpath}: {status}"
    except Exception as e:
        return False, f"{fpath}: {e}"

# =========================================================
# PER-PATIENT PUSH LOCK (§3.1 / §3.9 fix)
# =========================================================
# handle_store() used to spawn an independent push_single_patient() call
# on EVERY incoming C-STORE when autoroute is on. A single study is
# typically dozens to hundreds of images arriving in a burst, so that
# meant dozens of overlapping push attempts for the same patient at
# once -- duplicate sends, association storms against the remote PACS,
# and thrashing of set_status()/checkpoint state, since
# push_single_patient() itself also recurses into its own retry-with-
# backoff logic (see the recursive call at the end of
# _push_single_patient_body() below).
#
# _push_pid_locks / _push_pid_locks_guard mirror the doc-transfer
# per-pid lock pattern (_doc_transfer_acquire_pid_lock /
# _doc_transfer_release_pid_lock above) so that at most one push is ever
# active for a given patient ID at a time, no matter which of the four
# call sites triggered it: auto-route (handle_store), the offline queue
# worker, a manual push job (run_push_job / run_push_job_multi), or the
# "Push Selected" flow. A second push for the same pid simply blocks
# until the first one (including all of its own retries) finishes,
# instead of running concurrently with it.
_push_pid_locks = {}
_push_pid_locks_guard = threading.Lock()


def _push_acquire_pid_lock(pid):
    with _push_pid_locks_guard:
        lock = _push_pid_locks.get(pid)
        if lock is None:
            lock = threading.Lock()
            _push_pid_locks[pid] = lock
    lock.acquire()
    return lock


def _push_release_pid_lock(pid, lock):
    lock.release()
    # Opportunistic cleanup: if nobody else is waiting on this lock right
    # now, drop it from the registry so a long-running receiver doesn't
    # accumulate one Lock object per patient ID forever.
    with _push_pid_locks_guard:
        current = _push_pid_locks.get(pid)
        if current is lock and lock.acquire(blocking=False):
            lock.release()
            del _push_pid_locks[pid]


# §3.1: coalesce a burst of incoming images for the same patient into a
# SINGLE auto-route push, fired AUTOROUTE_DEBOUNCE_SEC after the last
# image received rather than once per image. The per-PID push lock above
# is the hard guarantee against overlapping pushes; this debounce is what
# stops a 200-image study from queuing 200 near-simultaneous push
# attempts (each of which would re-scan the whole patient folder) against
# it in the first place.
AUTOROUTE_DEBOUNCE_SEC = 2.0
_autoroute_pending_timers = {}
_autoroute_debounce_lock = threading.Lock()


def _schedule_autoroute_push(pid):
    """Called from handle_store() instead of spawning a push thread
    directly. Resets a per-patient quiet-period timer on every call; only
    the last call within AUTOROUTE_DEBOUNCE_SEC of silence actually fires
    a push. Safe to call from the C-STORE handler thread for every image
    in a burst."""
    def _fire():
        with _autoroute_debounce_lock:
            _autoroute_pending_timers.pop(pid, None)
        threading.Thread(target=push_single_patient, args=(pid,), daemon=True).start()

    with _autoroute_debounce_lock:
        existing = _autoroute_pending_timers.get(pid)
        if existing is not None:
            existing.cancel()
        timer = threading.Timer(AUTOROUTE_DEBOUNCE_SEC, _fire)
        timer.daemon = True
        _autoroute_pending_timers[pid] = timer
        timer.start()


def push_single_patient(pid, on_progress=None, destination=None, anonymize=False):
    """Public entry point for pushing every image of one patient --
    kept as the same name/signature/return shape every existing caller
    (GUI code, offline queue worker, run_push_job, etc.) already uses.
    The actual push logic lives in _push_single_patient_body(); this
    wrapper's only job is the §3.1/§3.9 per-PID lock, so however many
    threads try to push the same patient at once, only one push
    association is ever active for that patient at a time. The lock
    covers _push_single_patient_body()'s own retry-with-backoff
    recursion too, since that recurses directly into
    _push_single_patient_body() rather than back through this wrapper
    (a plain, non-reentrant Lock is used deliberately -- see the
    docstring on the recursive call site)."""
    lock = _push_acquire_pid_lock(pid)
    try:
        return _push_single_patient_body(pid, on_progress=on_progress, destination=destination, anonymize=anonymize)
    finally:
        _push_release_pid_lock(pid, lock)


def _push_single_patient_body(pid, on_progress=None, destination=None, anonymize=False):
    """Push every image of one patient. Resolves the destination via the
    routing rules engine unless one is explicitly supplied. Retries failed
    files with exponential backoff up to MAX_RETRY_ATTEMPTS. Returns
    (sent, total, ok)."""
    cfg = destination or resolve_destination_for_patient(pid)
    if not cfg:
        write_push_log(pid, "", "No push destination configured")
        set_fields(pid, status=STATUS_FAILED, last_error="No push destination configured")
        return 0, 0, False

    # Validate the destination's AE titles/port BEFORE touching the network.
    # A misconfigured field here (too long, blank, stray whitespace) is a
    # permanent problem — no amount of retrying will fix it, so we want a
    # clear, immediate error instead of a cryptic pynetdicom exception
    # buried three retries deep.
    ok, remote_ae_clean = validate_ae_title(cfg.get("ae"), "Remote AE Title")
    if not ok:
        write_push_log(pid, "", remote_ae_clean)
        set_fields(pid, status=STATUS_FAILED, last_error=remote_ae_clean)
        return 0, 0, False

    ok, calling_clean = validate_ae_title(cfg.get("calling_ae") or DEFAULT_PUSH_CALLING_AE, "Calling AE Title")
    if not ok:
        write_push_log(pid, "", calling_clean)
        set_fields(pid, status=STATUS_FAILED, last_error=calling_clean)
        return 0, 0, False

    ok, port_val = validate_port(cfg.get("port"), "Remote Port")
    if not ok:
        write_push_log(pid, "", port_val)
        set_fields(pid, status=STATUS_FAILED, last_error=port_val)
        return 0, 0, False

    folder = get_patient_folder(pid)
    if not os.path.isdir(folder):
        write_push_log(pid, "", "No local files for patient")
        set_fields(pid, status=STATUS_FAILED, last_error="No local files for patient")
        return 0, 0, False

    files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".dcm")]
    total = len(files)
    if total == 0:
        return 0, 0, True

    set_status(pid, STATUS_SENDING)

    tls_cfg = load_tls_config()
    ssl_context = build_ssl_context_for_client(tls_cfg, port=port_val, dest_label=cfg.get("name"))

    ae = AE(ae_title=calling_clean)
    ae.acse_timeout = get_network_timeout_sec()
    ae.dimse_timeout = get_network_timeout_sec()
    ae.network_timeout = get_network_timeout_sec()

    # Build the set of SOP Classes to negotiate.
    #
    # IMPORTANT: propose only what's actually needed for THIS push
    # (the exact SOPClassUIDs found in the files, plus a couple of
    # baseline transfer syntaxes). Some production PACS enforce a much
    # tighter limit on the number of presentation contexts / total
    # A-ASSOCIATE-RQ size than the DICOM hard maximum of 128, and will
    # reject the ENTIRE association outright if it's proposed too many
    # contexts — this is why permissive dev servers (Orthanc, dcm4chee,
    # etc.) accept a push that a stricter vendor PACS rejects: the
    # bloated blanket proposal (every SOP class in sopclass.ini, every
    # push, regardless of what's being sent) tripped their limit. We
    # only fall back to the full static catalog if we couldn't read any
    # SOP Class UID from the files themselves (e.g. unreadable files),
    # so the association still has a chance of covering them.
    file_sop_classes = set()
    file_study_uids = set()
    native_ts_by_sop = defaultdict(list)
    for fpath in files:
        try:
            ds = pydicom.dcmread(fpath, stop_before_pixels=True, force=True)
            sop_uid = str(getattr(ds, "SOPClassUID", "")).strip()
            if sop_uid:
                file_sop_classes.add(sop_uid)
                native_ts = str(getattr(ds.file_meta, "TransferSyntaxUID", "")).strip()
                if native_ts and native_ts not in native_ts_by_sop[sop_uid]:
                    native_ts_by_sop[sop_uid].append(native_ts)
            study_uid_val = str(getattr(ds, "StudyInstanceUID", "")).strip()
            if study_uid_val:
                file_study_uids.add(study_uid_val)
        except Exception:
            pass  # unreadable file — still attempt push, send_c_store will catch it

    if file_sop_classes:
        all_sops = list(file_sop_classes)
    else:
        all_sops = load_extra_sops()
    all_sops = cap_sop_list(all_sops)

    if not all_sops:
        last_error = ("No usable SOP Class found for this patient's files (all files "
                      "unreadable and sopclass.ini has no entries) — nothing to propose.")
        write_push_log(pid, "", last_error)
        set_fields(pid, status=STATUS_FAILED, last_error=last_error)
        return 0, total, False

    # "Use all transfer syntaxes": propose everything sopclass.ini has
    # configured, PLUS the full pydicom-known transfer syntax catalog
    # (ALL_TRANSFER_SYNTAXES -- Implicit/Explicit VR, JPEG family, JPEG-LS,
    # JPEG2000/HTJ2K family, RLE, MPEG/HEVC video, SMPTE ST 2110, etc), not
    # just the 3 uncompressed baseline syntaxes. This maximises the odds
    # that whatever the destination PACS natively supports is on offer, so
    # fewer transfers fall back to in-memory decompress/re-encode.
    configured_ts = load_transfer_syntaxes()
    fallback_ts = list(dict.fromkeys(configured_ts + ALL_TRANSFER_SYNTAXES))  # preserve order, deduplicate

    for sop in all_sops:
        try:
            # §fix: propose each file's own native transfer syntax FIRST for
            # its SOP Class, before the configured/baseline fallback list.
            # Previously only fallback_ts was proposed, so a peer that
            # actually supports e.g. JPEG2000 natively was never even
            # offered it unless JPEG2000 happened to be listed in
            # sopclass.ini -- every such file was forced through the
            # in-memory decompress/re-encode path in _send_one_file(),
            # which both wastes CPU and depends on pylibjpeg/gdcm being
            # installed. Native-first means the common case (peer supports
            # the native encoding) needs zero transcoding at all.
            native_list = [UID(ts) for ts in native_ts_by_sop.get(sop, [])]
            sop_ts_list = list(dict.fromkeys(native_list + fallback_ts))
            ae.add_requested_context(sop, sop_ts_list)
        except Exception:
            log_exception(f"Failed to add requested context for {sop}")

    sent = 0
    last_error = ""
    retryable_failure = True
    # Resume Transfers: skip whatever this destination has already
    # confirmed received in a prior attempt (tracked by SOP Instance UID,
    # which is exactly the file's basename since handle_store saves files
    # as "{sop_uid}.dcm"). If there's no matching checkpoint (first
    # attempt, or a different destination than last time), this is just
    # every file, same as before.
    dest_name_for_checkpoint = cfg.get("name", "")
    already_sent_uids = get_transfer_checkpoint(pid, dest_name_for_checkpoint)
    if already_sent_uids:
        # "D:" prefix keeps DICOM checkpoint identifiers in their own
        # namespace, distinct from document-transfer identifiers (which
        # would use "F:") -- see record_checkpoint_progress call below.
        remaining_files = [f for f in files
                           if f"D:{os.path.splitext(os.path.basename(f))[0]}" not in already_sent_uids]
        sent = total - len(remaining_files)  # credit for what a prior attempt already delivered
        if sent:
            with data_lock:
                push_job["attempted_images"] += sent
                push_job["sent_images"] += sent
            _push_resume_progress[pid] = f"{sent}/{total}"
            app_logger.info("Resuming push for %s: %d/%d already delivered, %d remaining",
                             pid, sent, total, len(remaining_files))
    else:
        remaining_files = list(files)
    assoc_established = False
    push_started_at = time.time()
    association_time_sec = None

    if not remaining_files:
        # Checkpoint shows this destination already has every file --
        # nothing left to resume, so skip the association entirely.
        push_patient_documents(pid, cfg)
        set_fields(
            pid, status=STATUS_SENT,
            sent_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            push_target=f"{remote_ae_clean}@{cfg['ip']}:{port_val}",
            last_error="", retry_count=0,
        )
        clear_transfer_checkpoint(pid)
        write_audit_log("PUSH-OK", f"pid={pid} sent={sent}/{total} dest={remote_ae_clean}@{cfg['ip']}:{port_val} (resumed, nothing left to send)")
        record_push_stat(next(iter(file_study_uids), None), sent, 0)
        _push_resume_progress.pop(pid, None)
        return sent, total, True

    try:
        assoc_kwargs = {}
        if ssl_context is not None:
            assoc_kwargs["tls_args"] = (ssl_context, None)
        _assoc_start = time.time()
        assoc, diag = associate_with_diagnostics(ae, cfg["ip"], port_val, remote_ae_clean, **assoc_kwargs)
        association_time_sec = round(time.time() - _assoc_start, 3)
        if not diag["ok"]:
            last_error = diag["message"]
            retryable_failure = diag["retryable"]
            write_push_log(pid, "", last_error)
            write_audit_log("PUSH-ASSOC-FAILED", f"pid={pid} dest={remote_ae_clean}@{cfg['ip']}:{port_val} {last_error}")
            set_fields(pid, status=STATUS_FAILED, last_error=last_error)
            if not retryable_failure:
                # Permanent/config-level rejection (bad AE title, etc) —
                # retrying with the exact same config will just fail again
                # after wasting the whole exponential-backoff schedule.
                _push_resume_progress.pop(pid, None)
                return 0, total, False
            # Transient issue (busy server, timeout, network blip) — fall
            # through to the shared retry logic below rather than touching
            # the association any further.
        else:
            assoc_established = True
            for fpath in remaining_files:
                if push_job["stop_flag"]:
                    last_error = "Stopped by user"
                    break

                ok, err = _send_one_file(assoc, fpath, anonymize, pid)
                if ok:
                    sent += 1
                    sop_uid_sent = os.path.splitext(os.path.basename(fpath))[0]
                    record_checkpoint_progress(pid, dest_name_for_checkpoint, f"D:{sop_uid_sent}")
                else:
                    last_error = err
                    write_push_log(pid, "", err)

                with data_lock:
                    push_job["attempted_images"] += 1
                    if ok:
                        push_job["sent_images"] += 1
                push_job["current_pid"] = pid
                push_job["current_dest"] = f"{remote_ae_clean}@{cfg['ip']}:{port_val}"
                if on_progress:
                    on_progress(sent, total)
                ui_event_queue.put(("push_progress", None))

            assoc.release()

    except Exception as e:
        if not last_error:
            last_error = str(e)
        write_push_log(pid, "", last_error)
        log_exception(f"Association-level failure pushing patient {pid}")

    failed_count = total - sent

    if sent == total:
        push_patient_documents(pid, cfg)
        set_fields(
            pid,
            status=STATUS_SENT,
            sent_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            push_target=f"{remote_ae_clean}@{cfg['ip']}:{port_val}",
            last_error="",
            retry_count=0,
        )
        write_audit_log("PUSH-OK", f"pid={pid} sent={sent}/{total} dest={remote_ae_clean}@{cfg['ip']}:{port_val}")
        clear_transfer_checkpoint(pid)
        transfer_time_sec = round(time.time() - push_started_at, 3)
        write_push_log(
            pid, "", "",
            final_status="OK",
            destination_name=cfg.get("name", ""),
            destination_ae=remote_ae_clean,
            destination_ip=cfg.get("ip", ""),
            destination_port=port_val,
            calling_ae=calling_clean,
            called_ae=remote_ae_clean,
            images_attempted=total,
            images_sent=sent,
            images_failed=total - sent,
            retry_count=patient_data.get(pid, {}).get("retry_count", 0),
            transfer_time_sec=transfer_time_sec,
            transfer_speed_files_per_sec=round(sent / transfer_time_sec, 3) if transfer_time_sec else None,
            compression_used=False,
            tls_enabled=ssl_context is not None,
            association_time_sec=association_time_sec,
            anonymized=anonymize,
        )
        record_push_stat(next(iter(file_study_uids), None), sent, total - sent)
        _push_resume_progress.pop(pid, None)
        return sent, total, True

    # Partial or total failure: retry the failed remainder with backoff,
    # unless the user explicitly stopped the job or the failure was a
    # permanent/config-level rejection that retrying can't fix.
    retry_count = 0
    with data_lock:
        retry_count = patient_data.get(pid, {}).get("retry_count", 0)

    if push_job["stop_flag"] or not retryable_failure:
        record_checkpoint_interruption(pid, dest_name_for_checkpoint, last_error or f"{failed_count} image(s) failed")
        set_fields(pid, status=STATUS_FAILED, last_error=last_error or f"{failed_count} image(s) failed")
        write_audit_log("PUSH-FAILED", f"pid={pid} sent={sent}/{total} error={last_error}")
        write_push_log(
            pid, "", last_error or f"{failed_count} image(s) failed",
            final_status="FAILED",
            destination_name=cfg.get("name", ""),
            destination_ae=remote_ae_clean,
            destination_ip=cfg.get("ip", ""),
            destination_port=port_val,
            calling_ae=calling_clean,
            called_ae=remote_ae_clean,
            images_attempted=total,
            images_sent=sent,
            images_failed=failed_count,
            retry_count=retry_count,
            transfer_time_sec=round(time.time() - push_started_at, 3),
            tls_enabled=ssl_context is not None,
            association_time_sec=association_time_sec,
            failure_reason=last_error,
            anonymized=anonymize,
        )
        record_push_stat(next(iter(file_study_uids), None), sent, failed_count)
        _push_resume_progress.pop(pid, None)
        return sent, total, False

    if retry_count >= MAX_RETRY_ATTEMPTS:
        # In-process retries (with exponential backoff) are exhausted but
        # this is still a transient/network-level failure -- rather than
        # abandoning the patient as a dead "Failed" entry, hand it off to
        # the persistent Offline Queue. The checkpoint already recorded
        # above means whenever this resumes (destination comes back
        # online, or an admin retries manually), it picks up at file
        # `sent`, not file 0.
        record_checkpoint_interruption(pid, dest_name_for_checkpoint, last_error or f"{failed_count} image(s) failed")
        enqueue_offline(pid, dest_name_for_checkpoint, last_error or f"{failed_count} image(s) failed")
        set_fields(pid, status=STATUS_QUEUED, last_error=last_error or f"{failed_count} image(s) failed")
        write_audit_log("PUSH-QUEUED-OFFLINE",
                        f"pid={pid} sent={sent}/{total} dest={dest_name_for_checkpoint} error={last_error}")
        write_push_log(
            pid, "", last_error or f"{failed_count} image(s) failed",
            final_status="QUEUED",
            destination_name=cfg.get("name", ""),
            destination_ae=remote_ae_clean,
            destination_ip=cfg.get("ip", ""),
            destination_port=port_val,
            calling_ae=calling_clean,
            called_ae=remote_ae_clean,
            images_attempted=total,
            images_sent=sent,
            images_failed=failed_count,
            retry_count=retry_count,
            transfer_time_sec=round(time.time() - push_started_at, 3),
            tls_enabled=ssl_context is not None,
            association_time_sec=association_time_sec,
            failure_reason=last_error,
            anonymized=anonymize,
        )
        record_push_stat(next(iter(file_study_uids), None), sent, failed_count)
        _push_resume_progress.pop(pid, None)
        return sent, total, False

    delay = RETRY_BASE_DELAY_SEC * (2 ** retry_count)
    record_checkpoint_interruption(pid, dest_name_for_checkpoint, last_error or f"{failed_count} image(s) failed")
    set_fields(pid, status=STATUS_RETRYING, last_error=last_error, retry_count=retry_count + 1)
    write_audit_log("PUSH-RETRY-SCHEDULED", f"pid={pid} attempt={retry_count + 1} delay={delay}s")

    for _ in range(int(delay * 10)):
        if push_job["stop_flag"] or app_shutdown_event.is_set():
            set_fields(pid, status=STATUS_FAILED, last_error="Stopped before retry")
            _push_resume_progress.pop(pid, None)
            return sent, total, False
        time.sleep(0.1)

    # Recurse directly into the body, NOT the public push_single_patient()
    # wrapper -- the wrapper holds a non-reentrant per-PID lock (§3.1/§3.9)
    # that this same thread is already holding for the whole retry chain;
    # going back through the wrapper here would self-deadlock.
    return _push_single_patient_body(pid, on_progress=on_progress, destination=destination, anonymize=anonymize)


def run_push_job_multi(patient_ids, destinations_list, anonymize=False, worker_threads=None):
    """2.1 -- fans a push job out across multiple destinations. run_push_job
    already blocks its caller's thread until its own ThreadPoolExecutor
    finishes (see the `with ThreadPoolExecutor(...) as executor:` above),
    so destinations are pushed to one after another here, each one still
    fanning the patient list out across up to DEFAULT_PUSH_WORKER_THREADS
    workers exactly as a normal single-destination job would. Per-destination
    offline-queue-on-failure (enqueue_offline) behavior is untouched --
    it happens inside push_single_patient exactly as it does today."""
    for dest in destinations_list:
        run_push_job(patient_ids, destination=dest, anonymize=anonymize, worker_threads=worker_threads)


def run_push_job(patient_ids, destination=None, anonymize=False, worker_threads=None):
    """Push a list of patients concurrently, updating an overall progress
    counter. `destination`, if given, overrides per-patient routing rules
    for this whole job (used by the "Push Selected to..." flow)."""
    if push_job["running"]:
        return

    if not load_destinations():
        ui_event_queue.put(("push_error", "No push destinations configured. Add one first."))
        return

    with data_lock:
        total_images = sum(patient_data.get(pid, {}).get("count", 0) for pid in patient_ids)

    push_job["running"] = True
    push_job["total_images"] = max(total_images, 1)
    push_job["sent_images"] = 0
    push_job["attempted_images"] = 0
    push_job["stop_flag"] = False
    push_job["started_at"] = time.time()

    ui_event_queue.put(("push_started", None))
    write_audit_log("PUSH-JOB-START", f"patients={len(patient_ids)} images={total_images}")

    for pid in patient_ids:
        set_status(pid, STATUS_PENDING)

    def worker(pid):
        return push_single_patient(pid, destination=destination, anonymize=anonymize)

    threads = worker_threads or get_max_worker_threads()
    threads = max(1, min(threads, MAX_PUSH_WORKER_THREADS))

    failures = 0
    try:
        with ThreadPoolExecutor(max_workers=threads) as executor:
            futures = {executor.submit(worker, pid): pid for pid in patient_ids}
            for fut in as_completed(futures):
                pid = futures[fut]
                try:
                    _sent, _total, ok = fut.result()
                    if not ok:
                        failures += 1
                except Exception as e:
                    write_push_log(pid, "", str(e))
                    log_exception(f"Unhandled exception in push worker for {pid}")
                    set_fields(pid, status=STATUS_FAILED, last_error=str(e))
                    failures += 1
    finally:
        push_job["running"] = False
        ui_event_queue.put(("push_finished", None))
        write_audit_log(
            "PUSH-JOB-END",
            f"patients={len(patient_ids)} failures={failures} "
            f"sent_images={push_job['sent_images']}/{push_job['total_images']}"
        )
        if failures:
            notify_event("push_failed", "Push Job Finished", f"{failures} patient(s) failed. Check Logs tab.")
        else:
            notify_event("push_complete", "Push Job Finished", "All patients pushed successfully.")


def stop_push_job():
    push_job["stop_flag"] = True
    write_audit_log("PUSH-JOB-STOP", "User requested stop")


def reset_patient_status(pid):
    """Reset/unlock a study so it becomes eligible for push again."""
    set_fields(pid, status=STATUS_PENDING, last_error="", retry_count=0)


def get_push_throughput_eta():
    """Returns (images_per_sec, eta_seconds) for the currently running push
    job, or (0, None) if not running / not enough data yet."""
    if not push_job["running"] or not push_job["started_at"]:
        return 0.0, None
    elapsed = max(time.time() - push_job["started_at"], 0.001)
    sent = push_job["sent_images"]
    rate = sent / elapsed
    remaining = max(push_job["total_images"] - push_job["attempted_images"], 0)
    eta = remaining / rate if rate > 0 else None
    return rate, eta

# =========================================================
# DICOM QUERY / RETRIEVE (C-FIND + C-MOVE)
# =========================================================
# Lets the user query a remote PACS (study-level) and pull selected
# studies back to this node via C-MOVE, addressed to our own receiver
# AE Title/port so retrieved studies land straight in the worklist
# (the receiver must be running for C-MOVE retrieves to succeed).

def query_remote_studies(remote_ae, remote_ip, remote_port, patient_id="", patient_name="",
                          study_date="", modality="", timeout=15):
    """Study-level C-FIND. Returns (ok, list_of_result_dicts, message)."""
    results = []
    try:
        ae = AE(ae_title="RAPPS_FIND")
        ae.add_requested_context(StudyRootQueryRetrieveInformationModelFind)
        ae.acse_timeout = timeout
        ae.dimse_timeout = timeout
        ae.network_timeout = timeout

        tls_cfg = load_tls_config()
        ssl_context = build_ssl_context_for_client(tls_cfg, port=int(remote_port))
        assoc_kwargs = {}
        if ssl_context is not None:
            assoc_kwargs["tls_args"] = (ssl_context, None)

        assoc, diag = associate_with_diagnostics(ae, remote_ip, int(remote_port), remote_ae, **assoc_kwargs)
        if not diag["ok"]:
            return False, [], diag["message"]

        query = pydicom.Dataset()
        query.QueryRetrieveLevel = "STUDY"
        query.PatientID = patient_id or ""
        query.PatientName = patient_name or ""
        query.StudyDate = study_date or ""
        query.ModalitiesInStudy = modality or ""
        query.StudyInstanceUID = ""
        query.StudyDescription = ""
        query.NumberOfStudyRelatedInstances = ""

        responses = assoc.send_c_find(query, StudyRootQueryRetrieveInformationModelFind)
        for status, identifier in responses:
            if status and getattr(status, "Status", None) in (0xFF00, 0xFF01) and identifier is not None:
                results.append({
                    "patient_id": str(getattr(identifier, "PatientID", "")),
                    "patient_name": format_dicom_person_name(getattr(identifier, "PatientName", "")),
                    "study_date": str(getattr(identifier, "StudyDate", "")),
                    "study_uid": str(getattr(identifier, "StudyInstanceUID", "")),
                    "description": str(getattr(identifier, "StudyDescription", "")),
                    "modality": str(getattr(identifier, "ModalitiesInStudy", "")),
                    "instances": str(getattr(identifier, "NumberOfStudyRelatedInstances", "")),
                })

        assoc.release()
        write_audit_log("C-FIND", f"{remote_ae}@{remote_ip}:{remote_port} results={len(results)}")
        return True, results, f"{len(results)} stud{'y' if len(results)==1 else 'ies'} found."

    except Exception as e:
        log_exception("C-FIND failed")
        write_audit_log("C-FIND-FAILED", str(e))
        return False, [], f"C-FIND failed: {e}"


def retrieve_study(remote_ae, remote_ip, remote_port, study_uid, our_ae_title, timeout=30):
    """Study-level C-MOVE, addressed back to our own receiver AE so the
    retrieved instances arrive via the normal handle_store path and show
    up in the worklist automatically. The receiver MUST be running."""
    try:
        ae = AE(ae_title="RAPPS_MOVE")
        ae.add_requested_context(StudyRootQueryRetrieveInformationModelMove)
        ae.acse_timeout = timeout
        ae.dimse_timeout = timeout
        ae.network_timeout = timeout

        tls_cfg = load_tls_config()
        ssl_context = build_ssl_context_for_client(tls_cfg, port=int(remote_port))
        assoc_kwargs = {}
        if ssl_context is not None:
            assoc_kwargs["tls_args"] = (ssl_context, None)

        assoc, diag = associate_with_diagnostics(ae, remote_ip, int(remote_port), remote_ae, **assoc_kwargs)
        if not diag["ok"]:
            return False, diag["message"]

        query = pydicom.Dataset()
        query.QueryRetrieveLevel = "STUDY"
        query.StudyInstanceUID = study_uid

        responses = assoc.send_c_move(query, our_ae_title, StudyRootQueryRetrieveInformationModelMove)
        final_status = None
        for status, _identifier in responses:
            final_status = status

        assoc.release()

        if final_status and getattr(final_status, "Status", None) == 0x0000:
            write_audit_log("C-MOVE-OK", f"study={study_uid} from={remote_ae} to={our_ae_title}")
            # D.7 (design note): Query/Retrieve only ever moves DICOM
            # objects. There is no source-PACS document-transfer protocol
            # to piggyback on here -- even if `remote_ae` also happens to
            # be configured as a push destination with
            # doc_transfer_enabled=True pointing back at this instance,
            # there is nothing meaningful to pull: the remote side would
            # need to initiate a push_patient_documents() call itself, not
            # us. Left as a known, clearly-marked gap rather than a
            # speculative protocol addition.
            return True, "Retrieve completed; instances will appear in the worklist as they arrive."
        write_audit_log("C-MOVE-FAILED", f"study={study_uid} status={final_status}")
        return False, f"C-MOVE finished with status: {final_status}"

    except Exception as e:
        log_exception("C-MOVE failed")
        write_audit_log("C-MOVE-FAILED", str(e))
        return False, f"C-MOVE failed: {e}"

# =========================================================
# TRAY
# =========================================================

def create_tray_image():
    logo = _load_logo_pil()
    if logo is not None:
        try:
            icon = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
            fit = logo.copy()
            fit.thumbnail((64, 64))
            offset = ((64 - fit.width) // 2, (64 - fit.height) // 2)
            icon.paste(fit, offset, fit)
            return icon
        except Exception:
            pass
    img = Image.new("RGB", (64, 64), "black")
    draw = ImageDraw.Draw(img)
    draw.ellipse((16, 16, 48, 48), fill="white")
    return img


_tray_icon_ref = {"icon": None}


def graceful_shutdown():
    """Stop the receiver, signal background loops to stop, then exit."""
    app_shutdown_event.set()
    if APP_SETTINGS.get("remember_window_geometry", True):
        try:
            _update_app_setting("last_window_geometry", app.geometry())
        except Exception:
            pass
    try:
        stop_receiver_server()
    except Exception:
        log_exception("Error during receiver shutdown")
    stop_push_job()
    write_audit_log("APP_SHUTDOWN", "Application exiting")
    icon = _tray_icon_ref.get("icon")
    if icon:
        try:
            icon.stop()
        except Exception:
            pass
    os._exit(0)


def on_quit(icon, item):
    graceful_shutdown()


def on_show(icon, item):
    # pystray invokes menu callbacks on its own icon thread, not the Tk
    # main thread -- calling a widget method directly here would be the
    # same class of cross-thread Tkinter bug as any other unmarshaled
    # .configure() call from a background thread.
    app.after(0, app.deiconify)


def minimize_to_tray():
    app.withdraw()
    icon = pystray.Icon(
        "DICOM", create_tray_image(), "R-Apps DICOM Receiver",
        pystray.Menu(
            pystray.MenuItem("Show", on_show),
            pystray.MenuItem("Exit", on_quit),
        ),
    )
    _tray_icon_ref["icon"] = icon
    threading.Thread(target=icon.run, daemon=True).start()

# =========================================================
# DESKTOP TOAST NOTIFICATIONS  (cross-platform best-effort)
# =========================================================

def _show_toast(title, message):
    """Non-blocking desktop notification. Tries plyer first, then
    falls back to a plain tkinter messagebox-free overlay, then silently
    drops it. Never raises — toasts are best-effort. Respects
    Settings > Notifications (enable/disable + duration)."""
    if not toasts_enabled():
        return
    duration_ms = notification_duration_ms()
    try:
        from plyer import notification
        notification.notify(title=title, message=message, app_name="R-Apps DICOM",
                             timeout=max(1, duration_ms // 1000))
        return
    except Exception:
        pass
    # Fallback: tiny floating label in the corner of the app window
    try:
        import tkinter as tk
        # Tk quirk (most visible on Windows): creating ANY Toplevel whose
        # implicit master is a withdrawn root silently re-maps that root,
        # which is exactly why minimizing to tray used to make the whole
        # app pop back open on every single notification. Note whether we
        # were withdrawn beforehand, then re-withdraw right after the
        # toast is created so only the small popup itself is visible.
        was_withdrawn = False
        was_iconic = False
        try:
            _state = app.state()
            was_withdrawn = (_state == "withdrawn")
            was_iconic = (_state == "iconic")
        except Exception:
            pass
        toast = tk.Toplevel(app)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        sw, sh = toast.winfo_screenwidth(), toast.winfo_screenheight()
        toast.geometry(f"340x55+{sw - 360}+16")
        toast.configure(bg=THEME_SURFACE)
        tk.Label(
            toast, text=f"{title}\n{message}", bg=THEME_SURFACE, fg=THEME_TEXT,
            font=(FONT_FAMILY, 10), wraplength=320, justify="left",
            padx=8, pady=6,
        ).pack(fill="both", expand=True)
        toast.after(duration_ms, toast.destroy)
        if was_withdrawn:
            try:
                app.withdraw()
            except Exception:
                pass
        elif was_iconic:
            try:
                app.iconify()
            except Exception:
                pass
    except Exception:
        pass


def show_toast_threadsafe(title, message):
    """Queue a toast that will be shown on the main thread via the event
    pump. Using this from background threads avoids Tk thread-safety issues.
    Respects Settings > Notifications > Critical-only mode by dropping
    non-critical toasts before they're even queued."""
    if not toasts_enabled():
        return
    if APP_SETTINGS.get("critical_only_notifications") and \
            not any(w in title.lower() for w in ("fail", "error", "critical", "offline", "down")):
        return
    ui_event_queue.put(("toast", (title, message)))

# =========================================================
# MAIN GUI
# =========================================================

ctk.set_appearance_mode("dark")  # permanent -- this app has no Light theme
ctk.set_default_color_theme("blue")

# THEME PALETTE -- single source of truth for the GUI polish pass. Cool
# slate/navy palette (clinical-workstation tone) rather than a
# high-contrast "developer dark mode" black. Computed once at import time
# from DARK_THEME_PALETTE (folding in High Contrast Mode if enabled) --
# every widget builder below reads these module-level names.
_ACTIVE_THEME = _compute_theme_palette()
THEME_BG = _ACTIVE_THEME["bg"]
THEME_SURFACE = _ACTIVE_THEME["surface"]
THEME_HEADING_BG = _ACTIVE_THEME["heading_bg"]
THEME_ACCENT = _ACTIVE_THEME["accent"]
THEME_ACCENT_HOVER = _ACTIVE_THEME["accent_hover"]
THEME_TEXT = _ACTIVE_THEME["text"]
THEME_TEXT_MUTED = _ACTIVE_THEME["text_muted"]
THEME_NEUTRAL_BTN = _ACTIVE_THEME["neutral_btn"]
THEME_NEUTRAL_BTN_HOVER = _ACTIVE_THEME["neutral_btn_hover"]
THEME_DANGER = _ACTIVE_THEME["danger"]
THEME_DANGER_HOVER = _ACTIVE_THEME["danger_hover"]
THEME_SUCCESS = _ACTIVE_THEME["success"]
THEME_SUCCESS_HOVER = _ACTIVE_THEME["success_hover"]
THEME_WARNING = _ACTIVE_THEME["warning"]
THEME_WARNING_HOVER = _ACTIVE_THEME["warning_hover"]
THEME_SEGMENTED_HOVER = _ACTIVE_THEME["segmented_hover"]
THEME_ODD_ROW = _ACTIVE_THEME["odd_row"]
FONT_FAMILY = "Segoe UI" if platform.system() == "Windows" else "Helvetica"

APP_VERSION = "3.0.0"

# =========================================================
# DESIGN SYSTEM -- shared primitives (Phase 1)
# =========================================================
# Single source of truth for typography scale and reusable widget
# builders (cards, status badges) so every page/tab styles itself the
# same way. These read the live THEME_* globals at call time (not at
# import time), so anything built with them stays correct across
# refresh_ui_theme() calls (Accessibility settings) -- same pattern the
# treeview styling already uses. Nothing here touches DICOM/business
# logic; it only standardizes how existing pages present it.

FONT_SCALE = {
    "title": 17,      # page/app title
    "subtitle": 12,    # secondary header text
    "section": 14,     # card/section headings
    "body": 12,        # normal text
    "small": 11,       # dense table / caption text
    "micro": 10,       # timestamps, footnotes
    "caption": 13,     # card/panel sub-headings (e.g. "Disk Space", "Receiver Status")
    "section_lg": 15,  # tab-page-level section headers ("PACS Health Monitor", etc.)
    "value_lg": 16,    # stat-tile / popup summary value displays
    "value_xl": 18,    # larger stat-tile value displays
    "kpi": 20,         # dashboard KPI numbers
    "kpi_lg": 26,       # largest dashboard KPI numbers
}


def get_font(scale="body", weight="normal"):
    """Returns a CTkFont at one of the standard scale steps above, scaled
    by the user's Settings > Appearance > Font Size preference (see
    get_font_scale_multiplier()). Every label/button in the app already
    routes through this one function, so the Font Size setting takes
    effect everywhere immediately without touching individual widgets."""
    size = FONT_SCALE.get(scale, FONT_SCALE["body"])
    try:
        size = max(6, round(size * get_font_scale_multiplier()))
    except Exception:
        pass
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)


# =========================================================
# ICON LIBRARY -- Lucide icons (bundled, rendered offline)
# =========================================================
# Every icon in the UI comes from this one function, so stroke width,
# proportions and sizing stay identical everywhere (nav rail, toolbar
# buttons, dialogs). Source glyphs are pre-rendered 64x64 white-on-
# transparent PNGs (see lucide_icons_data.py); at request time we tint
# them to whatever color the caller needs (so the same asset works for
# normal / muted / accent / on-accent-button icon color) and cache the
# resulting CTkImage so repeat calls (e.g. re-theming) are cheap.
_icon_cache = {}


def get_icon(name, size=18, color=None):
    """Returns a ctk.CTkImage for the given Lucide icon name, tinted to
    `color` (defaults to THEME_TEXT). Returns None if the icon library
    could not be loaded (e.g. Pillow missing) -- all call sites must
    tolerate an image of None (CTkButton/CTkLabel handle that fine)."""
    if not LUCIDE_ICONS_AVAILABLE:
        return None
    color = color or THEME_TEXT
    cache_key = (name, size, color)
    cached = _icon_cache.get(cache_key)
    if cached is not None:
        return cached
    b64 = LUCIDE_ICONS_B64.get(name)
    if b64 is None:
        return None
    try:
        raw = _icon_b64.b64decode(b64)
        glyph = Image.open(_icon_io.BytesIO(raw)).convert("RGBA")
        glyph = glyph.resize((size, size), Image.LANCZOS)
        r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
        solid = Image.new("RGBA", glyph.size, (r, g, b, 255))
        # Use the source glyph's alpha (white strokes on transparent) as
        # a mask so the recolored solid shows through only the strokes.
        tinted = Image.composite(solid, Image.new("RGBA", glyph.size, (0, 0, 0, 0)), glyph.split()[3])
        img = ctk.CTkImage(light_image=tinted, dark_image=tinted, size=(size, size))
    except Exception:
        return None
    _icon_cache[cache_key] = img
    return img


def _render_warning_rows(frame, messages, color=None, wraplength=500):
    """Rebuilds `frame` (a plain CTkFrame) with one triangle-alert icon +
    text row per message in `messages`. This is the icon-based
    replacement for labels that used to prefix each warning line with a
    literal '⚠ ' emoji character -- the icon now carries that meaning,
    consistent with every other inline warning in the app (see the
    pyzipper-missing notice on the Export tab for the original pattern
    this was factored out of)."""
    if color is None:
        color = THEME_WARNING
    for child in frame.winfo_children():
        child.destroy()
    for msg in messages:
        ctk.CTkLabel(
            frame, text=msg, image=get_icon("triangle-alert", size=14, color=color),
            compound="left", font=get_font("small"), text_color=color,
            justify="left", wraplength=wraplength, anchor="w",
        ).pack(fill="x", anchor="w", pady=(2, 0))


# status-kind -> (fg, dot) color lookup against the *current* palette.
def _status_kind_colors(kind):
    kind = (kind or "neutral").lower()
    mapping = {
        "success": (THEME_SUCCESS, THEME_SUCCESS),
        "running": (THEME_SUCCESS, THEME_SUCCESS),
        "connected": (THEME_SUCCESS, THEME_SUCCESS),
        "completed": (THEME_SUCCESS, THEME_SUCCESS),
        "danger": (THEME_DANGER, THEME_DANGER),
        "failed": (THEME_DANGER, THEME_DANGER),
        "offline": (THEME_DANGER, THEME_DANGER),
        "stopped": (THEME_DANGER, THEME_DANGER),
        "warning": (THEME_WARNING, THEME_WARNING),
        "retrying": (THEME_WARNING, THEME_WARNING),
        "pending": (THEME_TEXT_MUTED, THEME_TEXT_MUTED),
        "sending": (THEME_ACCENT, THEME_ACCENT),
        "receiving": (THEME_ACCENT, THEME_ACCENT),
        "resuming": ("#17a2b8", "#17a2b8"),
        "neutral": (THEME_TEXT_MUTED, THEME_TEXT_MUTED),
    }
    return mapping.get(kind, mapping["neutral"])


def make_status_badge(parent, text, kind="neutral"):
    """A small pill-style status indicator (dot + label) matching the
    app's status-badge vocabulary (Running/Stopped/Connected/Offline/
    Sending/Receiving/Pending/Failed/Retrying/Completed/...). Returns
    the CTkFrame; call badge.update_status(text, kind) later to change it
    in place without rebuilding."""
    color, dot_color = _status_kind_colors(kind)
    frame = ctk.CTkFrame(parent, fg_color=THEME_HEADING_BG, corner_radius=999, height=28)
    frame.pack_propagate(False)
    dot = ctk.CTkLabel(frame, text="●", font=get_font("small"), text_color=dot_color, width=12)
    dot.pack(side="left", padx=(10, 2), pady=2)
    # MarqueeLabel instead of a plain CTkLabel: long status strings (e.g.
    # "Running (AE_TITLE@11112)") scroll smoothly instead of being clipped
    # by the pill's width.
    lbl = MarqueeLabel(frame, text=text, width=170, height=20,
                        font=get_font("small", "bold"), text_color=color,
                        canvas_bg=THEME_HEADING_BG)
    lbl.pack(side="left", padx=(0, 12), pady=2, fill="x", expand=True)

    def update_status(new_text, new_kind="neutral"):
        c, dc = _status_kind_colors(new_kind)
        dot.configure(text_color=dc)
        lbl.configure(text=new_text, text_color=c)

    frame.update_status = update_status
    return frame


# =========================================================
# MARQUEE LABEL -- auto-scrolling text for fixed-width containers
# =========================================================
# Drop-in stand-in for a fixed-width CTkLabel: exposes the same
# .configure(text=..., text_color=...) surface so existing call sites
# don't need to change, but instead of letting long text clip/overflow
# its box, it automatically scrolls it smoothly left, pauses briefly,
# and loops -- only when the text is actually too wide to fit. Text
# that fits stays perfectly still, exactly like a normal label.
class MarqueeLabel(ctk.CTkFrame):
    # §fix (marquee speed): was _SPEED_PX=3 / _TICK_MS=25 -- 3px every 25ms
    # is 120px/sec, which reads as a fast blur rather than scrolling text
    # you can actually follow. 2px every 40ms is 50px/sec (40 ticks/sec is
    # still plenty smooth for a small label) -- a little under half the
    # old speed, and a comfortable reading pace.
    _SPEED_PX = 2           # pixels moved per animation tick
    _TICK_MS = 40           # animation tick interval
    _PAUSE_MS = 700         # pause at the start of each loop
    _GAP_PX = 40            # gap between the end of the text and its repeat
    _all_instances = []     # weak-ish registry (pruned lazily) for reduced_motion toggling

    def __init__(self, parent, text="", width=140, height=20, font=None,
                 text_color=None, anchor="w", fg_color="transparent", canvas_bg=None, **kwargs):
        super().__init__(parent, width=width, height=height, fg_color=fg_color, **kwargs)
        MarqueeLabel._all_instances.append(self)
        self.pack_propagate(False)
        self.grid_propagate(False)
        self._box_w = width
        self._box_h = height
        self._font = font or get_font("micro")
        self._text_color = text_color or THEME_TEXT_MUTED
        self._anchor = anchor
        self._text = text
        self._after_id = None
        self._offset = 0
        self._text_w = 0

        self._canvas = tk.Canvas(self, width=width, height=height,
                                  highlightthickness=0, bd=0, bg=canvas_bg or THEME_SURFACE)
        self._canvas.pack(fill="both", expand=True)
        self._canvas.bind("<Configure>", lambda _e: self._relayout())
        self._text_id = self._canvas.create_text(
            0, height // 2, text=text, anchor="w", font=self._font, fill=self._text_color)
        self.set_text(text, force=True)

    # -- public, CTkLabel-compatible API -------------------------------
    def configure(self, text=None, text_color=None, font=None, **kwargs):
        if text_color is not None:
            self._text_color = text_color
            self._canvas.itemconfigure(self._text_id, fill=text_color)
        if font is not None:
            self._font = font
            self._canvas.itemconfigure(self._text_id, font=font)
        if text is not None and text != self._text:
            # Only rebuild/restart the scroll when the text actually
            # changed. Callers (status bar, KPI cards, etc.) re-configure()
            # with the *same* value on every periodic refresh tick -- if we
            # reset the animation every time, a mid-scroll label would
            # never get further than a couple of words before snapping back
            # to the start, which looked like "it doesn't work".
            self.set_text(text)
        # Anything else (e.g. background color re-theme) falls through
        # to the underlying CTkFrame.
        if kwargs:
            try:
                super().configure(**kwargs)
            except Exception:
                pass

    def set_bg(self, color):
        try:
            self._canvas.configure(bg=color)
        except Exception:
            pass

    # -- internals ------------------------------------------------------
    def _relayout(self):
        new_w = self._canvas.winfo_width() or self._box_w
        if abs(new_w - self._box_w) <= 1:
            return  # no meaningful size change -- don't restart a live scroll
        self._box_w = new_w
        self.set_text(self._text, force=True)

    def set_text(self, text, force=False):
        if not force and text == self._text and self._after_id is not None:
            return
        self._text = text
        self._canvas.itemconfigure(self._text_id, text=text)
        self._text_w = self._font.measure(text) if hasattr(self._font, "measure") else len(text) * 7
        self._offset = 0
        if self._after_id is not None:
            self.after_cancel(self._after_id)
            self._after_id = None

        if self._text_w <= self._box_w:
            # Fits comfortably: render static, respecting the requested anchor.
            if self._anchor == "center":
                x = max(0, (self._box_w - self._text_w) // 2)
            elif self._anchor == "e":
                x = max(0, self._box_w - self._text_w)
            else:
                x = 0
            self._canvas.coords(self._text_id, x, self._box_h // 2)
        elif APP_SETTINGS.get("reduced_motion"):
            # Accessibility: Reduced Motion Mode is on -- don't scroll long
            # text, just truncate it with an ellipsis and render statically.
            measure = self._font.measure if hasattr(self._font, "measure") else (lambda s: len(s) * 7)
            truncated = self._text
            while truncated and measure(truncated + "…") > self._box_w:
                truncated = truncated[:-1]
            self._canvas.itemconfigure(self._text_id, text=(truncated + "…") if truncated else self._text)
            self._canvas.coords(self._text_id, 0, self._box_h // 2)
        else:
            # Overflows: run a continuous "complete circle" loop -- the text
            # starts entirely off-screen to the right, travels the whole
            # distance across the box, ends up entirely off-screen to the
            # left (a full lap, never just a partial shift), pauses, and
            # repeats indefinitely via modulo arithmetic so the position is
            # never lost or reset by anything except a genuine text change.
            self._canvas.coords(self._text_id, self._box_w, self._box_h // 2)
            self._after_id = self.after(self._PAUSE_MS, self._animate)

    def _animate(self):
        if not self.winfo_exists():
            return
        loop_w = self._box_w + self._text_w  # full lap: enter right edge -> exit left edge
        self._offset = (self._offset + self._SPEED_PX) % loop_w
        x = self._box_w - self._offset
        self._canvas.coords(self._text_id, x, self._box_h // 2)
        # Pause briefly right as a lap completes (offset wraps back near 0),
        # then keep the loop going indefinitely.
        if self._offset < self._SPEED_PX:
            self._after_id = self.after(self._PAUSE_MS, self._animate)
        else:
            self._after_id = self.after(self._TICK_MS, self._animate)

    @classmethod
    def refresh_all_for_motion_setting(cls):
        """Re-evaluates Reduced Motion Mode for every live MarqueeLabel
        immediately -- called right after the setting is toggled, rather
        than waiting for the next time each label's text happens to
        change."""
        still_alive = []
        for lbl in cls._all_instances:
            try:
                if lbl.winfo_exists():
                    still_alive.append(lbl)
                    lbl.set_text(lbl._text, force=True)
            except Exception:
                pass
        cls._all_instances = still_alive


# =========================================================
# BUTTON MARQUEE -- every CTkButton in the app, automatically
# =========================================================
# CTkButton renders its caption with a real tkinter.Label (self._text_label)
# rather than drawing text on its canvas, so instead of touching every one
# of the hundreds of CTkButton(...) call sites throughout the app, this
# patches the CTkButton class itself: any button created anywhere from this
# point on automatically gets a continuously-looping ticker for its text
# whenever the button isn't wide enough for it, and goes back to a plain,
# perfectly static caption the instant it fits (including on resize).
def _button_marquee_available_px(btn):
    try:
        btn.update_idletasks()
        w = btn.winfo_width()
        if w <= 1:
            try:
                w = int(btn._apply_widget_scaling(btn._desired_width))
            except Exception:
                w = 140
        img_w = 0
        img_lbl = getattr(btn, "_image_label", None)
        if img_lbl is not None:
            try:
                img_w = img_lbl.winfo_reqwidth() + 6
            except Exception:
                pass
        return max(w - img_w - 20, 12)
    except Exception:
        return 9999


def _install_button_marquee(btn):
    state = {"after_id": None, "full_text": btn.cget("text") or "", "ticking": False}

    def stop_tick():
        if state["after_id"] is not None:
            try:
                btn.after_cancel(state["after_id"])
            except Exception:
                pass
            state["after_id"] = None
        state["ticking"] = False

    def evaluate(*_a):
        full = state["full_text"]
        lbl = getattr(btn, "_text_label", None)
        if not full or lbl is None:
            stop_tick()
            return
        try:
            fnt = tkfont.Font(font=lbl.cget("font"))
        except Exception:
            fnt = None
        avail = _button_marquee_available_px(btn)
        needed = fnt.measure(full) if fnt else len(full) * 7
        if needed <= avail:
            # Fits: make sure it's showing the plain, full, static caption.
            stop_tick()
            try:
                lbl.configure(text=full)
            except Exception:
                pass
            return
        if state["ticking"]:
            return  # already looping at the right size -- leave it running
        state["ticking"] = True
        sep = "     •     "
        buffer = full + sep
        blen = len(buffer)
        avg_char_px = max((fnt.measure("abcdefghijklmnopqrstuvwxyz") / 26.0) if fnt else 7, 4)
        nchars = max(int(avail / avg_char_px), 4)
        idx = {"i": 0}

        def tick():
            try:
                window = (buffer * 2)[idx["i"]: idx["i"] + nchars]
                lbl.configure(text=window)
            except Exception:
                stop_tick()
                return
            idx["i"] = (idx["i"] + 1) % blen
            # §fix (marquee speed): was 70ms/char (~14.3 chars/sec), which
            # visibly jumped a full glyph-width 14 times a second -- fast
            # enough to feel like a jittery blur rather than readable
            # scrolling text. 150ms/char (~6.7 chars/sec) is a comfortable
            # reading pace, consistent with the MarqueeLabel speed fix.
            state["after_id"] = btn.after(150, tick)

        tick()

    def on_configure_event(_e=None):
        btn.after(60, evaluate)

    _orig_configure = btn.configure

    def _patched_configure(*args, **kwargs):
        cnf = dict(args[0]) if args and isinstance(args[0], dict) else {}
        cnf.update(kwargs)
        if "text" in cnf:
            state["full_text"] = cnf["text"] or ""
        result = _orig_configure(*args, **kwargs)
        if "text" in cnf or "image" in cnf:
            stop_tick()
            btn.after(30, evaluate)
        return result

    btn.configure = _patched_configure
    btn.config = _patched_configure
    btn.bind("<Configure>", on_configure_event, add="+")
    btn.after(80, evaluate)


_BaseCTkButton = ctk.CTkButton


class _MarqueeCTkButton(_BaseCTkButton):
    """Drop-in CTkButton: identical everywhere except its caption
    automatically ticks when the button is too narrow for it. Registered
    as ctk.CTkButton below so every existing CTkButton(...) call site in
    the app (built before or after this point) gets it for free."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        try:
            _install_button_marquee(self)
        except Exception:
            log_exception("Failed to attach marquee behavior to a button")


ctk.CTkButton = _MarqueeCTkButton


def _card_padding():
    """Card header/body padding, tightened when Compact Mode is enabled.
    Single choke point for make_card/make_card_row/add_card_to_row so
    Compact Mode affects every card-based panel in the app identically."""
    if APP_SETTINGS.get("compact_mode"):
        return {"header_top": 8, "header_bottom_notitle": 4, "subtitle_bottom": 4,
                "body_bottom": 10, "row_padx": 4, "card_padx": 3, "card_pady": 3}
    return {"header_top": 14, "header_bottom_notitle": 8, "subtitle_bottom": 8,
            "body_bottom": 16, "row_padx": 6, "card_padx": 5, "card_pady": 5}


def make_equal_share_row(parent, specs, pady=(8, 0), fill="x", anchor=None):
    """Generalized make_equal_button_row: each spec is
    {"factory": callable(row) -> widget, "shrink": bool}. Non-shrink
    widgets (e.g. field labels) always keep their natural size; shrink
    widgets (entries, dropdowns, buttons) share the row's leftover width
    in equal ratios once everything stops fitting -- same adaptive
    pack-until-it-doesn't-fit-then-grid behavior as make_equal_button_row,
    just generalized past plain CTkButton rows (e.g. Routing Rules, where
    the Modality/Institution/Source AE entries need to shrink together
    with the buttons, not just the buttons alone).
    Returns (row_frame, [widgets]) in spec order.
    """
    row = ctk.CTkFrame(parent, fg_color="transparent")
    if fill is not None:
        if anchor:
            row.pack(fill=fill, pady=pady, anchor=anchor)
        else:
            row.pack(fill=fill, pady=pady)

    widgets = [spec["factory"](row) for spec in specs]
    shrink_flags = [spec.get("shrink", False) for spec in specs]
    state = {"mode": None}

    def relayout(_evt=None):
        if not row.winfo_exists():
            return
        row.update_idletasks()
        available = row.winfo_width()
        if available <= 1:
            return
        pad_each = 6
        natural_total = sum(w.winfo_reqwidth() for w in widgets) + pad_each * len(widgets)
        fits = natural_total <= available
        new_mode = "pack" if fits else "grid"
        if new_mode == state["mode"]:
            return
        state["mode"] = new_mode
        for i, w in enumerate(widgets):
            w.grid_forget()
            w.pack_forget()
            row.grid_columnconfigure(i, weight=0, uniform="")
        if fits:
            for w in widgets:
                w.pack(side="left", padx=3, pady=2)
        else:
            for i, w in enumerate(widgets):
                if shrink_flags[i]:
                    row.grid_columnconfigure(i, weight=1, uniform="eqshare")
                    w.grid(row=0, column=i, sticky="ew", padx=3, pady=2)
                else:
                    row.grid_columnconfigure(i, weight=0, minsize=w.winfo_reqwidth())
                    w.grid(row=0, column=i, sticky="w", padx=3, pady=2)

    row.bind("<Configure>", relayout, add="+")
    row.after(10, relayout)
    return row, widgets


def make_equal_button_row(parent, button_specs, pady=(8, 0), fill="x", anchor=None):
    """A row of buttons that behave like normal buttons (their own natural
    size, left to right) as long as they all fit -- and only switch to
    sharing the row's width in equal ratios once they genuinely don't fit,
    so every button shrinks together instead of the trailing ones
    disappearing off the edge.

    button_specs: list of dicts of CTkButton kwargs (text, command, fg_color,
    hover_color, width, etc). Returns (row_frame, [button_widgets]) so
    callers can .configure() later.
    """
    row = ctk.CTkFrame(parent, fg_color="transparent")
    if fill is not None:
        if anchor:
            row.pack(fill=fill, pady=pady, anchor=anchor)
        else:
            row.pack(fill=fill, pady=pady)

    widgets = [ctk.CTkButton(row, **spec) for spec in button_specs]
    state = {"mode": None}

    def relayout(_evt=None):
        if not row.winfo_exists():
            return
        row.update_idletasks()
        available = row.winfo_width()
        if available <= 1:
            return
        pad_each = 6  # padx=3 on both sides
        natural_total = sum(b.winfo_reqwidth() for b in widgets) + pad_each * len(widgets)
        fits = natural_total <= available
        new_mode = "pack" if fits else "grid"
        if new_mode == state["mode"]:
            return
        state["mode"] = new_mode
        for i, b in enumerate(widgets):
            b.grid_forget()
            b.pack_forget()
            row.grid_columnconfigure(i, weight=0, uniform="")
        if fits:
            for b in widgets:
                b.pack(side="left", padx=3, pady=2)
        else:
            for i, b in enumerate(widgets):
                row.grid_columnconfigure(i, weight=1, uniform="eqbtn")
                b.grid(row=0, column=i, sticky="ew", padx=3, pady=2)

    row.bind("<Configure>", relayout, add="+")
    row.after(10, relayout)
    return row, widgets


def make_wrapped_label(parent, text, wraplength, font=None, text_color=None, justify="left", **kwargs):
    """CTkLabel with wraplength doesn't grow its own height to fit the
    wrapped text -- the widget keeps whatever height it was given (or the
    library default), so anything past the first line or two gets clipped
    at the bottom. This estimates the wrapped line count from the font
    metrics and text length and sizes the label to fit, so descriptive
    paragraphs (Settings, tab headers, etc.) render in full."""
    f = font or get_font("small")
    try:
        tk_font = tkfont.Font(font=f)
    except Exception:
        tk_font = tkfont.Font()
    line_height = tk_font.metrics("linespace")
    lines = 1
    current = ""
    for word in text.split():
        candidate = f"{current} {word}".strip()
        if current and tk_font.measure(candidate) > wraplength:
            lines += 1
            current = word
        else:
            current = candidate
    lines += text.count("\n")
    height = max(20, int(lines * (line_height + 4)) + 4)
    return ctk.CTkLabel(parent, text=text, font=f, text_color=text_color,
                         wraplength=wraplength, justify=justify, height=height, **kwargs)


def make_card(parent, title=None, subtitle=None, fg_color=None):
    """Standard enterprise 'card' container: rounded surface, optional
    title/subtitle header, and a `.body` frame for page content to pack
    into. Used as the common building block for Dashboard tiles, Receiver/
    Pusher panels, Reports, Settings groups, etc.

    Filled with a solid tone distinct from the page background rather than
    relying on a thin border for definition -- a 1px outline in a color
    close to the fill reads as a faint, incomplete-looking line rather
    than a solid box."""
    pad = _card_padding()
    card = ctk.CTkFrame(parent, fg_color=fg_color or THEME_HEADING_BG, corner_radius=12)
    if title:
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(pad["header_top"], 0 if subtitle else pad["header_bottom_notitle"]))
        ctk.CTkLabel(header, text=title, font=get_font("section", "bold"),
                     text_color=THEME_TEXT).pack(side="left")
        if subtitle:
            ctk.CTkLabel(card, text=subtitle, font=get_font("micro"),
                         text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=16, pady=(2, pad["subtitle_bottom"]))
    body = ctk.CTkFrame(card, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=16, pady=(0, pad["body_bottom"]))
    card.body = body
    return card


def make_card_row(parent, pady=(0, 0)):
    """Container for a row of equal-width cards, laid out with grid + a
    uniform column group instead of pack(fill='x', expand=True).

    CTkFrame pre-renders its rounded-corner border at creation size; when
    several equal-width siblings are stretched afterward via pack's
    expand=True, that border image doesn't always get regenerated at the
    new size, so some cards render with a side that looks open/incomplete.
    grid() computes final widths in a single pass before first paint, which
    avoids that resize-after-render artifact. Use with add_card_to_row()."""
    row = ctk.CTkFrame(parent, fg_color="transparent")
    row.pack(fill="x", padx=_card_padding()["row_padx"], pady=pady)
    row._card_col = 0
    return row


def add_card_to_row(row, card, padx=None, pady=None):
    """Places `card` as the next equal-width column in `row` (see make_card_row)."""
    pad = _card_padding()
    if padx is None:
        padx = pad["card_padx"]
    if pady is None:
        pady = pad["card_pady"]
    col = row._card_col
    row._card_col += 1
    row.grid_columnconfigure(col, weight=1, uniform="card_row")
    card.grid(row=0, column=col, padx=padx, pady=pady, sticky="nsew")



    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        pass  # Older Windows without shcore, or running under Wine -- degrade gracefully

if TKINTERDND2_AVAILABLE:
    try:
        class _DnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
            """Adds tkdnd's native OS drag-and-drop to a normal CTk root,
            via the documented retrofit pattern (mix in DnDWrapper, call
            TkinterDnD._require on self) rather than requiring the root
            to be created as TkinterDnD.Tk() from the start -- keeps
            this a pure CTk subclass everywhere else in the app."""
            def __init__(self, *a, **kw):
                super().__init__(*a, **kw)
                self.TkdndVersion = TkinterDnD._require(self)
        app = _DnDCTk()
    except Exception:
        TKINTERDND2_AVAILABLE = False
        app = ctk.CTk()
else:
    app = ctk.CTk()
app.title("R-Apps DICOM")
app.geometry("1600x900")
app.minsize(1250, 750)
app.configure(fg_color=THEME_BG)
try:
    apply_ui_scaling()  # Settings > Appearance > UI Scaling, applied before first paint
except Exception:
    pass
try:
    _logo_for_icon = _load_logo_pil()
    if _logo_for_icon is not None:
        app.iconphoto(True, ImageTk.PhotoImage(_logo_for_icon))
except Exception:
    pass
app.withdraw()  # stay hidden until the one-time Admin PIN setup (if needed) completes

# =========================================================
# STARTUP LOADING SCREEN
# =========================================================
# A real progress indicator, not a decorative timer: _splash_step() is
# called at genuine milestones as the UI gets built and, later, as
# startup() loads real data (CSV worklist, config, destinations, the
# worklist trees). The bar only advances when that actual work finishes,
# so if a step is slow (e.g. a large worklist CSV) the bar visibly stalls
# on that step's label instead of lying about progress. Only ever runs
# once, at initial launch -- never again for the life of the process.
_splash_state = {"win": None, "bar": None, "label": None}


def _keep_toplevel_small(win, max_width=520, max_height=640):
    """Constrain a popup/dialog Toplevel so it never grows larger than a
    sensible size, regardless of its contents. Safe to call multiple times
    or on windows that get destroyed before idle tasks run."""
    try:
        win.resizable(False, False)
        win.update_idletasks()
        req_w = win.winfo_reqwidth()
        req_h = win.winfo_reqheight()
        w = min(req_w, max_width) if req_w > 1 else max_width
        h = min(req_h, max_height) if req_h > 1 else max_height
        win.geometry(f"{w}x{h}")
    except Exception:
        pass


def _safe_grab_set(win, attempts=6, delay_ms=40):
    """Robust wrapper around Tk's Toplevel.grab_set().

    grab_set() raises TclError if called before the window manager has
    actually mapped the window -- which can happen right after restoring
    the app from the system tray, or on some window managers/remote
    desktop setups even for a normal dialog open. Every dialog in this
    app used to call win.grab_set() bare, so on an unlucky timing hit the
    exception propagated out of the triggering button's command
    callback, leaving a non-modal dialog behind: the window underneath
    remained clickable, so a stray click could open a second, independent
    copy of the same dialog. For the Admin Login dialog specifically,
    that meant two Toplevels each fighting for the grab -- reported as
    "admin login works at first, then eventually starts malfunctioning"
    over an extended session. Retry a few times a beat apart instead of
    raising; the window stays visible and usable even on the rare case
    where the grab never succeeds."""
    def _try(n):
        try:
            if not win.winfo_exists():
                return
            win.grab_set()
        except Exception:
            if n > 0:
                win.after(delay_ms, lambda: _try(n - 1))
            else:
                log_exception("grab_set() failed after retries -- dialog will stay non-modal")
    _try(attempts)



    """Forces a floating Toplevel/CTkToplevel panel or dialog to open as
    a small windowed popup, never full-screen/maximized.

    On several platforms (most notably Windows), a Toplevel created while
    its master is maximized or in a fullscreen/"zoomed" state can itself
    inherit that state, which made small panels like the Notification
    Center balloon out to fill the whole screen. Explicitly forcing
    'normal' state -- both immediately and again shortly after the window
    manager finishes mapping the window -- keeps these panels at their
    intended size regardless of whether the main app window is
    maximized, fullscreen, or windowed. Defined here (before its first
    caller, the splash screen below) because the splash is built and
    shown immediately at module-load time, earlier than every other
    Toplevel in the app."""
    try:
        win.state("normal")
    except Exception:
        pass

    def _reassert():
        try:
            if win.winfo_exists() and win.state() not in ("normal", "withdrawn", "iconic"):
                win.state("normal")
        except Exception:
            pass

    win.after(10, _reassert)
    win.after(150, _reassert)


def _build_splash_screen():
    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.overrideredirect(True)
    win.attributes("-topmost", True)
    try:
        win.configure(fg_color=THEME_BG)
    except Exception:
        pass
    w, h = 460, 260
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    x, y = (sw - w) // 2, (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    shell = ctk.CTkFrame(win, fg_color=THEME_SURFACE, corner_radius=16,
                          border_width=1, border_color=THEME_HEADING_BG)
    shell.pack(fill="both", expand=True)

    try:
        _logo_ctk = make_logo_ctk_image(size=72)
        if _logo_ctk is not None:
            ctk.CTkLabel(shell, image=_logo_ctk, text="").pack(pady=(36, 8))
    except Exception:
        pass

    ctk.CTkLabel(shell, text="R-Apps DICOM", font=(FONT_FAMILY, 20, "bold"),
                 text_color=THEME_TEXT).pack(pady=(0, 2))
    ctk.CTkLabel(shell, text="Receiver + AutoRouter", font=(FONT_FAMILY, 11),
                 text_color=THEME_TEXT_MUTED).pack(pady=(0, 18))

    bar = ctk.CTkProgressBar(shell, width=340, height=8, corner_radius=4,
                              progress_color=THEME_ACCENT)
    bar.set(0.0)
    bar.pack(pady=(0, 10))

    label = ctk.CTkLabel(shell, text="Starting up…", font=(FONT_FAMILY, 11),
                          text_color=THEME_TEXT_MUTED)
    label.pack()

    win.update()
    _splash_state["win"] = win
    _splash_state["bar"] = bar
    _splash_state["label"] = label


def _splash_step(fraction, text):
    """Advances the splash to `fraction` (0-1) with `text` describing the
    step that just finished. Calls win.update() to force an immediate
    repaint -- necessary because nothing else pumps the Tk event loop
    while the rest of this module runs its synchronous setup code, so
    without this the window would just sit there frozen despite genuine
    progress happening underneath it."""
    win = _splash_state.get("win")
    if win is None or not win.winfo_exists():
        return
    try:
        _splash_state["bar"].set(max(0.0, min(1.0, fraction)))
        _splash_state["label"].configure(text=text)
        win.update()
    except Exception:
        pass


def _close_splash_screen():
    win = _splash_state.get("win")
    if win is not None and win.winfo_exists():
        try:
            win.destroy()
        except Exception:
            pass
    _splash_state["win"] = None


try:
    _build_splash_screen()
    _splash_step(0.04, "Starting up…")
except Exception:
    log_exception("Failed to build splash screen")


def on_window_close():
    """Respects Settings > Startup & Behavior: minimize-to-tray-instead-
    of-exit and confirm-before-close, both of which are configurable and
    take effect immediately (no restart required)."""
    if APP_SETTINGS.get("minimize_to_tray_on_close", False):
        minimize_to_tray()
        return
    if not APP_SETTINGS.get("confirm_before_close", True):
        graceful_shutdown()
        return
    if modern_askyesno("Exit", "Exit the application? The receiver will stop."):
        graceful_shutdown()


app.protocol("WM_DELETE_WINDOW", on_window_close)

# ---- shared treeview containers (assigned below, used in callbacks) ----
rec_tree_ref = {}
push_tree_ref = {}
log_text_ref = {}
log_struct_records_by_iid = {}  # tree item id -> full parsed jsonl record, for the detail panel

# =========================================================
# APP HEADER (branding strip)
# =========================================================

header_bar = ctk.CTkFrame(app, fg_color=THEME_SURFACE, corner_radius=12, height=56)
header_bar.pack(fill="x", padx=14, pady=(14, 8))
header_bar.pack_propagate(False)

_header_logo_img = make_logo_ctk_image(size=28)
if _header_logo_img is not None:
    ctk.CTkLabel(header_bar, text="", image=_header_logo_img).pack(side="left", padx=(18, 0))

header_title_lbl = ctk.CTkLabel(
    header_bar, text="R-Apps DICOM",
    font=get_font("title", "bold"),
    text_color=THEME_TEXT,
)
header_title_lbl.pack(side="left", padx=(8 if _header_logo_img is not None else 18, 8))

header_subtitle_lbl = ctk.CTkLabel(
    header_bar, text="Receiver & AutoRouter",
    font=get_font("small"),
    text_color=THEME_TEXT_MUTED,
)
header_subtitle_lbl.pack(side="left", padx=(0, 8))

header_search_var = ctk.StringVar(value="")
header_search_icon_lbl = ctk.CTkLabel(header_bar, text="", image=get_icon("search", size=15, color=THEME_TEXT_MUTED))
header_search_icon_lbl.pack(side="left", padx=(18, 6))
header_search_entry = ctk.CTkEntry(
    header_bar, width=320, textvariable=header_search_var,
    placeholder_text="Search patients, destinations, routing, settings, logs…",
)
header_search_entry.pack(side="left", padx=(0, 0), fill="x", expand=True)
header_search_entry.bind("<KeyRelease>", lambda _evt: _on_universal_search_keyrelease())
header_search_entry.bind("<FocusOut>", lambda _evt: app.after(150, _maybe_close_universal_search_panel))
header_search_entry.bind("<Escape>", lambda _evt: _close_universal_search_panel())

# ---- Connection status chips: PACS / Encryption / Data store ----
header_chips_frame = ctk.CTkFrame(header_bar, fg_color="transparent")
header_chips_frame.pack(side="right", padx=(6, 10))

header_pacs_chip = make_status_badge(header_chips_frame, "PACS: —", kind="neutral")
header_pacs_chip.pack(side="left", padx=5)
header_enc_chip = make_status_badge(header_chips_frame, "Encryption: —", kind="neutral")
header_enc_chip.pack(side="left", padx=5)


def _bind_hover_tooltip(widget, text_provider, delay_ms=400):
    """Attaches a small themed hover tooltip to `widget` and, recursively,
    every widget nested inside it. Composite widgets like the badges
    make_status_badge() builds are covered edge-to-edge by their own
    children (a dot label + a MarqueeLabel, which itself wraps a canvas),
    so binding only the outer frame would silently never fire -- the
    cursor always enters a child widget first. `text_provider` is called
    fresh on each hover (not once at bind time) so the tooltip always
    reflects current state."""
    state = {"win": None, "show_after_id": None, "hide_after_id": None}

    def _cancel_pending():
        for key in ("show_after_id", "hide_after_id"):
            if state[key] is not None:
                try:
                    widget.after_cancel(state[key])
                except Exception:
                    pass
                state[key] = None

    def _destroy_now():
        win = state["win"]
        state["win"] = None
        if win is not None:
            try:
                if win.winfo_exists():
                    win.destroy()
            except Exception:
                pass

    def _show_now():
        state["show_after_id"] = None
        if not widget.winfo_exists():
            return
        text = text_provider()
        if not text:
            return
        _destroy_now()
        win = ctk.CTkToplevel(app)
        _keep_toplevel_small(win)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(fg_color=THEME_HEADING_BG)
        ctk.CTkLabel(win, text=text, font=get_font("small"), text_color=THEME_TEXT,
                     justify="left").pack(padx=10, pady=6)
        win.update_idletasks()
        w = max(win.winfo_reqwidth(), 10)
        h = max(win.winfo_reqheight(), 10)
        x = widget.winfo_rootx()
        y = widget.winfo_rooty() + widget.winfo_height() + 6
        x, y = _clamp_popover_geometry(x, y, w, h)
        win.geometry(f"{w}x{h}+{x}+{y}")
        state["win"] = win

    def _on_enter(_evt=None):
        _cancel_pending()
        state["show_after_id"] = widget.after(delay_ms, _show_now)

    def _on_leave(_evt=None):
        _cancel_pending()
        state["hide_after_id"] = widget.after(80, _destroy_now)

    def _collect(w):
        acc = [w]
        try:
            for c in w.winfo_children():
                acc.extend(_collect(c))
        except Exception:
            pass
        return acc

    for _w in _collect(widget):
        _w.bind("<Enter>", _on_enter, add="+")
        _w.bind("<Leave>", _on_leave, add="+")


def _pacs_chip_tooltip_text():
    if not destination_health_cache:
        return "No push destinations configured yet."
    lines = []
    for name, rec in destination_health_cache.items():
        state = "online" if rec.get("online") else "unreachable"
        lines.append(f"{name}: {state}")
    return "\n".join(lines) or "No destinations checked yet."


_bind_hover_tooltip(header_pacs_chip, _pacs_chip_tooltip_text)


def refresh_header_status_chips():
    """Keeps the header connection chips in sync with real state:
    PACS reachability comes from the same destination_health_cache the
    PACS Health tab and Dashboard already read; Encryption reflects whether
    the Fernet key file this app uses for rec.enc/push.enc exists.
    No new polling -- this just presents state the app already tracks."""
    try:
        if not destination_health_cache:
            header_pacs_chip.update_status("PACS: No Destinations", "neutral")
        else:
            any_online = any(rec.get("online") for rec in destination_health_cache.values())
            all_online = all(rec.get("online") for rec in destination_health_cache.values())
            if all_online:
                header_pacs_chip.update_status("PACS: Connected", "connected")
            elif any_online:
                header_pacs_chip.update_status("PACS: Degraded", "warning")
            else:
                header_pacs_chip.update_status("PACS: Offline", "offline")
    except Exception:
        log_exception("Header PACS chip refresh failed")

    try:
        if os.path.isfile(KEY_FILE):
            header_enc_chip.update_status("Encryption: Enabled", "connected")
        else:
            header_enc_chip.update_status("Encryption: Not Set Up", "warning")
    except Exception:
        log_exception("Header encryption chip refresh failed")


header_notif_btn = ctk.CTkButton(
    header_bar, text="", width=40, height=28, corner_radius=8,
    image=get_icon("bell", size=16, color=THEME_TEXT), compound="left",
    fg_color="transparent", hover_color=THEME_HEADING_BG, text_color=THEME_TEXT,
    font=get_font("small", "bold"),
)
header_notif_btn.pack(side="right", padx=(0, 6))

notification_panel_state = {"window": None}


def _notification_kind_color(kind):
    return {
        "success": THEME_SUCCESS, "error": THEME_DANGER,
        "warning": THEME_WARNING, "info": THEME_ACCENT,
    }.get(kind, THEME_TEXT_MUTED)


def _refresh_notification_center_ui():
    """Updates the bell badge count, and refreshes the panel's list if
    it's currently open. Called from pump_events (main thread) whenever
    notify_event() fires, and also on manual open/clear."""
    count = len(notification_history)
    header_notif_btn.configure(text=f"{count}" if count else "")

    win = notification_panel_state.get("window")
    if win is None or not win.winfo_exists():
        return
    for child in win.list_frame.winfo_children():
        child.destroy()
    if not notification_history:
        ctk.CTkLabel(win.list_frame, text="No notifications yet.",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(pady=20)
        return
    for rec in reversed(notification_history[-50:]):
        row = ctk.CTkFrame(win.list_frame, fg_color=THEME_HEADING_BG, corner_radius=8)
        row.pack(fill="x", padx=8, pady=4)
        dot = ctk.CTkLabel(row, text="●", font=get_font("caption"),
                            text_color=_notification_kind_color(rec["kind"]), width=18)
        dot.pack(side="left", padx=(8, 0), pady=8)
        text_col = ctk.CTkFrame(row, fg_color="transparent")
        text_col.pack(side="left", fill="x", expand=True, padx=(4, 8), pady=6)
        ctk.CTkLabel(text_col, text=rec["title"], font=get_font("small", "bold"),
                     text_color=THEME_TEXT, anchor="w", justify="left").pack(fill="x")
        make_wrapped_label(text_col, rec["message"], 340, font=get_font("micro"),
                            text_color=THEME_TEXT_MUTED, anchor="w").pack(fill="x")
        ctk.CTkLabel(text_col, text=rec["time"], font=get_font("micro"),
                     text_color=THEME_TEXT_MUTED, anchor="w").pack(fill="x")


def _clear_notification_history():
    notification_history.clear()
    _refresh_notification_center_ui()


NOTIFICATION_PANEL_W = 380
NOTIFICATION_PANEL_H = 460
NOTIFICATION_PANEL_MARGIN = 16
NOTIFICATION_PANEL_AUTO_MINIMIZE_MS = 8000  # auto-withdraws if left untouched


def _notification_panel_top_right_geometry():
    try:
        app.update_idletasks()
        sw = app.winfo_screenwidth()
    except Exception:
        sw = 1920
    x = max(0, sw - NOTIFICATION_PANEL_W - NOTIFICATION_PANEL_MARGIN)
    y = NOTIFICATION_PANEL_MARGIN
    return f"{NOTIFICATION_PANEL_W}x{NOTIFICATION_PANEL_H}+{x}+{y}"


def _toggle_notification_panel():
    win = notification_panel_state.get("window")
    if win is not None and win.winfo_exists():
        if win.state() == "withdrawn":
            win.deiconify()
            win.geometry(_notification_panel_top_right_geometry())
            win.lift()
            _refresh_notification_center_ui()
            _arm_notification_panel_auto_minimize(win)
        else:
            win.withdraw()
        return

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title("Notification Center")
    # Small, fixed, non-resizable popup pinned to the top-right corner --
    # never a full window regardless of the main app's state.
    win.resizable(False, False)
    win.geometry(_notification_panel_top_right_geometry())
    win.attributes("-topmost", True)
    win.configure(fg_color=THEME_SURFACE)
    # Hide (not destroy) on the window's own close button too, so its
    # CTkScrollableFrame -- and the permanent app-wide scroll binding it
    # registers -- is only ever created once for the life of the app.
    win.protocol("WM_DELETE_WINDOW", win.withdraw)
    win.bind("<Enter>", lambda _e: _cancel_notification_panel_auto_minimize(win))
    win.bind("<Leave>", lambda _e: _arm_notification_panel_auto_minimize(win))

    header = ctk.CTkFrame(win, fg_color="transparent")
    header.pack(fill="x", padx=12, pady=(12, 6))
    ctk.CTkLabel(header, text="Notification Center", font=get_font("section", "bold"),
                 text_color=THEME_TEXT).pack(side="left")
    ctk.CTkButton(header, text="Clear All", width=90, fg_color=THEME_NEUTRAL_BTN,
                  hover_color=THEME_NEUTRAL_BTN_HOVER, command=_clear_notification_history).pack(side="right")

    scroll = ctk.CTkScrollableFrame(win, fg_color="transparent")
    scroll.pack(fill="both", expand=True, padx=4, pady=(0, 8))
    win.list_frame = scroll

    notification_panel_state["window"] = win
    _refresh_notification_center_ui()
    _fade_in_window(win)
    _arm_notification_panel_auto_minimize(win)


def _cancel_notification_panel_auto_minimize(win):
    aid = notification_panel_state.pop("auto_min_after_id", None)
    if aid is not None:
        try:
            win.after_cancel(aid)
        except Exception:
            pass


def _arm_notification_panel_auto_minimize(win):
    """Auto-minimizes (withdraws) the notification panel if it's left
    open and un-hovered for a while, so it behaves like a transient popup
    rather than a window someone has to remember to close."""
    _cancel_notification_panel_auto_minimize(win)

    def _do_minimize():
        try:
            if win.winfo_exists() and win.state() != "withdrawn":
                win.withdraw()
        except Exception:
            pass

    try:
        notification_panel_state["auto_min_after_id"] = win.after(
            NOTIFICATION_PANEL_AUTO_MINIMIZE_MS, _do_minimize)
    except Exception:
        pass


header_notif_btn.configure(command=_toggle_notification_panel)

def _show_about_dialog():
    """Minimal, read-only About dialog. This is the only place version /
    build information is shown -- it no longer sits permanently in the
    header or status bar."""
    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title("About")
    win.geometry("360x220")
    win.resizable(False, False)
    win.configure(fg_color=THEME_SURFACE)
    win.transient(app)
    _safe_grab_set(win)

    body = ctk.CTkFrame(win, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=24)

    about_title_row = ctk.CTkFrame(body, fg_color="transparent")
    about_title_row.pack(anchor="w", fill="x")
    _about_logo_img = make_logo_ctk_image(size=26)
    if _about_logo_img is not None:
        ctk.CTkLabel(about_title_row, text="", image=_about_logo_img).pack(side="left", padx=(0, 8))
    ctk.CTkLabel(about_title_row, text="R-Apps DICOM", font=get_font("title", "bold"),
                 text_color=THEME_TEXT).pack(side="left")
    ctk.CTkLabel(body, text="Receiver & AutoRouter", font=get_font("subtitle"),
                 text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 16))
    ctk.CTkLabel(body, text=f"Version {APP_VERSION}", font=get_font("body"),
                 text_color=THEME_TEXT).pack(anchor="w")
    ctk.CTkLabel(body, text=f"Python {platform.python_version()}  ·  {platform.system()}",
                 font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 0))

    ctk.CTkButton(body, text="Close", width=100, height=32, corner_radius=8,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=win.destroy).pack(anchor="e", pady=(20, 0))


header_about_btn = ctk.CTkButton(
    header_bar, text="About", width=64, height=28, corner_radius=8,
    fg_color="transparent", hover_color=THEME_HEADING_BG, text_color=THEME_TEXT_MUTED,
    font=get_font("small"), command=_show_about_dialog,
)
header_about_btn.pack(side="right", padx=(0, 18))

# (Clock intentionally lives only in the bottom status bar -- see
# sb_clock_lbl / refresh_status_bar -- to keep the top bar focused on
# search + connection state, and to avoid showing the same live value
# in two places at once.)

# =========================================================
# ADMIN BAR (always visible, above the tabview)
# =========================================================

_splash_step(0.15, "Building interface…")

admin_bar = ctk.CTkFrame(app, fg_color=THEME_SURFACE, corner_radius=12)
admin_bar.pack(fill="x", padx=14, pady=(0, 8))

admin_bar_status_badge = make_status_badge(admin_bar, "Mode: USER", kind="neutral")
admin_bar_status_badge.pack(side="left", padx=(14, 15), pady=8)

# Back-compat shim: existing code elsewhere configures a text_color on
# `admin_bar_status_lbl` directly (see refresh_ui_theme / role-switch logic).
# Point that name at the badge's inner label so those call sites keep
# working unmodified.
admin_bar_status_dot = admin_bar_status_badge.winfo_children()[0]
admin_bar_status_lbl = admin_bar_status_badge.winfo_children()[1]

admin_login_btn = ctk.CTkButton(admin_bar, text="Admin Login", width=150, height=32,
                                 image=get_icon("lock", size=15, color="#ffffff"), compound="left",
                                 corner_radius=8, fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
                                 font=get_font("body", "bold"))
admin_login_btn.pack(side="left", padx=4, pady=6)

admin_switch_view_btn = ctk.CTkButton(admin_bar, text="Switch to User View", width=180, height=32,
                                       image=get_icon("user", size=15, color=THEME_TEXT), compound="left",
                                       corner_radius=8,
                                       fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                                       font=get_font("body"))

admin_change_pin_btn = ctk.CTkButton(admin_bar, text="Change Admin PIN", width=170, height=32,
                                      image=get_icon("key-round", size=15, color=THEME_TEXT), compound="left",
                                      corner_radius=8,
                                      fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                                      font=get_font("body"))

admin_lock_btn = ctk.CTkButton(admin_bar, text="Lock", width=110, height=32,
                                image=get_icon("lock", size=15, color="#ffffff"), compound="left",
                                corner_radius=8,
                                fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER,
                                font=get_font("body", "bold"))



# =========================================================
# SHARED WORKLIST TREEVIEW BUILDER
# =========================================================

WL_COLUMNS = (
    "sel", "patient_id", "patient_name", "institution", "modality",
    "count", "documents", "report", "history", "source", "status", "time", "sent_time", "push_target", "last_error",
)

WL_HEADINGS = {
    "sel": "\u2610",
    "patient_id": "Patient ID",
    "patient_name": "Patient Name",
    "institution": "Institution",
    "modality": "Mod.",
    "count": "#",
    "documents": "Images",
    "report": "Report",
    "history": "History",
    "source": "Source",
    "status": "Status",
    "time": "Received",
    "sent_time": "Sent",
    "push_target": "Pushed To",
    "last_error": "Last Error",
}

WL_WIDTHS = {
    "sel": 34, "patient_id": 170, "patient_name": 160, "institution": 230,
    "modality": 95, "count": 70, "documents": 175, "report": 175, "history": 175, "source": 190, "status": 170,
    "time": 190, "sent_time": 190, "push_target": 260, "last_error": 340,
}


# Columns that never get auto-fit to their content -- "sel" is a fixed-size
# checkbox glyph, and "patient_name" intentionally stays compact per spec
# (it's the one column allowed to truncate).
WL_NO_AUTOFIT_COLUMNS = {"sel", "patient_name"}


WORKLIST_COLUMN_WIDTHS_FILE = "worklist_column_widths.json"  # not secret -- remembers user-resized column widths per tree
_worklist_column_widths_cache = {"value": None}


def load_worklist_column_widths():
    """{"rec_tree": {"patient_id": 170, ...}, "push_tree": {...}} -- only
    ever contains overrides; any column/tree missing here just falls back
    to WL_WIDTHS as before."""
    if _worklist_column_widths_cache["value"] is not None:
        return _worklist_column_widths_cache["value"]
    data = {}
    try:
        if os.path.exists(WORKLIST_COLUMN_WIDTHS_FILE):
            with open(WORKLIST_COLUMN_WIDTHS_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                data = loaded
    except Exception:
        log_exception("Failed to load worklist_column_widths.json")
    _worklist_column_widths_cache["value"] = data
    return data


def save_worklist_column_widths(data):
    try:
        atomic_write(WORKLIST_COLUMN_WIDTHS_FILE, json.dumps(data, indent=2))
    except Exception:
        log_exception("Failed to save worklist_column_widths.json")
    _worklist_column_widths_cache["value"] = data


WORKLIST_TREE_STYLE = "WorklistWhite.Treeview"  # dedicated style so rec_tree/push_tree
                                                 # stay white regardless of the app's
                                                 # light/dark theme -- see apply_theme()
                                                 # and apply_ui_scaling(), which keep this
                                                 # style's row height/font in sync without
                                                 # ever touching its (fixed) colors.
WORKLIST_TREE_VISIBLE_ROWS = 22  # bigger default worklist -- was the ttk default of 10


def build_worklist_tree(parent, tree_name=None):
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background=THEME_SURFACE, foreground=THEME_TEXT,
                    fieldbackground=THEME_SURFACE, rowheight=get_worklist_row_height(), borderwidth=0,
                    font=(FONT_FAMILY, 11))
    style.configure("Treeview.Heading", background=THEME_HEADING_BG, foreground=THEME_TEXT,
                    font=(FONT_FAMILY, 11, "bold"), relief="flat", borderwidth=0)
    style.map("Treeview.Heading", background=[("active", THEME_SEGMENTED_HOVER)])
    style.map("Treeview", background=[("selected", THEME_ACCENT)],
              foreground=[("selected", "#ffffff")])
    style.configure("Stale.Treeview", foreground=STALE_HIGHLIGHT_COLOR)
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # drop default border frame

    # Receiver/Pusher worklists: normal (dark) theme background, same as
    # every other tree in the app -- just with row/heading text forced to
    # pure white for readability, independent of THEME_TEXT. A separate
    # style name so that guarantee survives even if THEME_TEXT ever
    # changes (e.g. High Contrast Mode recomputing the palette).
    style.configure(WORKLIST_TREE_STYLE, background=THEME_SURFACE, foreground="#ffffff",
                    fieldbackground=THEME_SURFACE, rowheight=get_worklist_row_height() + 6,
                    borderwidth=0, font=(FONT_FAMILY, 12))
    style.configure(f"{WORKLIST_TREE_STYLE}.Heading", background=THEME_HEADING_BG, foreground="#ffffff",
                    font=(FONT_FAMILY, 12, "bold"), relief="flat", borderwidth=0)
    style.map(f"{WORKLIST_TREE_STYLE}.Heading", background=[("active", THEME_SEGMENTED_HOVER)])
    style.map(WORKLIST_TREE_STYLE, background=[("selected", THEME_ACCENT)],
              foreground=[("selected", "#ffffff")])
    style.layout(WORKLIST_TREE_STYLE, [("Treeview.treearea", {"sticky": "nswe"})])

    style.configure("Vertical.TScrollbar", background=THEME_SURFACE, troughcolor=THEME_BG,
                    bordercolor=THEME_SURFACE, arrowcolor=THEME_TEXT_MUTED, width=14)
    style.configure("Horizontal.TScrollbar", background=THEME_SURFACE, troughcolor=THEME_BG,
                    bordercolor=THEME_SURFACE, arrowcolor=THEME_TEXT_MUTED, width=14)

    frame = ctk.CTkFrame(parent, fg_color=THEME_SURFACE, corner_radius=10)
    inner = ctk.CTkFrame(frame, fg_color=THEME_SURFACE, corner_radius=0)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    vsb = ttk.Scrollbar(inner, orient="vertical")
    hsb = ttk.Scrollbar(inner, orient="horizontal")

    tree = ttk.Treeview(
        inner, columns=WL_COLUMNS, show="headings",
        yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        selectmode="extended", style=WORKLIST_TREE_STYLE,
        height=WORKLIST_TREE_VISIBLE_ROWS,
    )
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    saved_widths = load_worklist_column_widths().get(tree_name, {}) if tree_name else {}
    for col in WL_COLUMNS:
        if col == "sel":
            tree.heading(col, text=WL_HEADINGS[col], command=lambda t=tree: _toggle_select_all_visible(t))
            tree.column(col, width=WL_WIDTHS[col], minwidth=WL_WIDTHS[col],
                        stretch=False, anchor="center")
            continue
        tree.heading(col, text=WL_HEADINGS[col],
                     command=lambda c=col: sort_tree(tree, c, False))
        tree.column(col, width=saved_widths.get(col, WL_WIDTHS[col]), minwidth=40, anchor="w")

    if tree_name:
        # 9.4 -- persist user-adjusted column widths. Treeview has no
        # dedicated "column resized" event, but a drag-resize always ends
        # with a ButtonRelease-1 on the tree, so that's the trigger point;
        # cheap to just read+save all current widths at that moment.
        _last_persisted_widths = {"value": dict(saved_widths)}

        def _persist_column_widths(_evt=None):
            current = {col: tree.column(col, "width") for col in WL_COLUMNS}
            if current == _last_persisted_widths["value"]:
                return  # ButtonRelease-1 also fires on ordinary row clicks -- skip the disk write if nothing actually resized
            _last_persisted_widths["value"] = current
            data = load_worklist_column_widths()
            data[tree_name] = current
            save_worklist_column_widths(data)
        tree.bind("<ButtonRelease-1>", _persist_column_widths, add="+")

    tree.tag_configure(STALE_HIGHLIGHT_TAG, foreground=STALE_HIGHLIGHT_COLOR)
    tree.tag_configure(SEARCH_MATCH_HIGHLIGHT_TAG, background=SEARCH_MATCH_HIGHLIGHT_BG)
    tree.tag_configure("even_row", background=THEME_SURFACE)
    tree.tag_configure("odd_row", background=THEME_ODD_ROW)
    # Status-column color coding (STATUS_COLORS existed but was never
    # wired to the tree before -- this is the actual "plain text status ->
    # visual indicator" fix for the app's main table).
    for _status_val, _status_color in STATUS_COLORS.items():
        tree.tag_configure(f"status_{_status_val}", foreground=_status_color)

    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")  # always visible -- never conditionally hidden
    tree.pack(fill="both", expand=True)
    return frame, tree



def sort_tree(tree, col, descending):
    data = [(tree.set(k, col), k) for k in tree.get_children("")]
    data.sort(reverse=descending, key=lambda x: x[0].lower() if isinstance(x[0], str) else x[0])
    for index, (_val, k) in enumerate(data):
        tree.move(k, "", index)
    tree.heading(col, command=lambda: sort_tree(tree, col, not descending))


_tree_render_cache = {}  # id(tree) -> (data_version, search) last rendered

# Gmail-style pagination: only trees registered in _PAGINATED_TREES get
# sliced into pages by populate_tree(); everything else (e.g. the Pusher
# tab's queue tree) keeps showing every matching row exactly as before.
_PAGINATED_TREES = set()
_tree_pagination = {}   # id(tree) -> {"page": 0, "page_size": 50, "total": 0, "filter_sig": None}
_pagination_ui = {}     # id(tree) -> (label, prev_btn, next_btn)
_rec_results_label_hooks = []  # callables refreshed after every rec_tree populate

# Cross-page persistent selection: Treeview selection state is per-render,
# so paginated trees (see _PAGINATED_TREES) lose the selection when the
# rows currently backing it get deleted/reinserted on a page change. This
# tracks "logically selected" PIDs independent of what's actually inserted
# right now, and populate_tree() re-applies it to whichever of those PIDs
# happen to be on the visible page after each rebuild.
_persistent_selection = {}  # id(tree) -> set of PIDs
_selection_count_hooks = {}  # id(tree) -> callable(n) refreshed on selection change


def _get_persistent_selection(tree):
    sel = _persistent_selection.get(id(tree))
    if sel is None:
        sel = set()
        _persistent_selection[id(tree)] = sel
    return sel


def _sync_persistent_selection(tree, _evt=None):
    """Bound to <<TreeviewSelect>>: merges the currently-visible selection
    state into the persistent set (rows selected on this page get added,
    rows on this page that got deselected get removed -- rows on OTHER
    pages are left untouched since Treeview has no opinion on them)."""
    persistent = _get_persistent_selection(tree)
    visible_ids = set(tree.get_children(""))
    visible_selected = set(tree.selection())
    persistent -= (visible_ids - visible_selected)
    persistent |= visible_selected
    hook = _selection_count_hooks.get(id(tree))
    if hook:
        hook(len(persistent))


def _reapply_persistent_selection(tree):
    """Call after populate_tree() re-renders a page: re-selects whichever
    currently-visible rows are in the persistent selection set."""
    persistent = _get_persistent_selection(tree)
    if not persistent:
        return
    visible_ids = set(tree.get_children(""))
    to_select = [pid for pid in visible_ids if pid in persistent]
    if to_select:
        tree.selection_set(to_select)


def _select_all_matching_filter(tree):
    """Adds every PID currently matching the tree's active search+filters
    (not just the current page) into the persistent selection set."""
    matched = _compute_filtered_pids(tree)
    persistent = _get_persistent_selection(tree)
    persistent.update(matched)
    _reapply_persistent_selection(tree)
    hook = _selection_count_hooks.get(id(tree))
    if hook:
        hook(len(persistent))


def _clear_persistent_selection(tree):
    _get_persistent_selection(tree).clear()
    tree.selection_remove(tree.selection())
    hook = _selection_count_hooks.get(id(tree))
    if hook:
        hook(0)


CHECKBOX_CHECKED = "\u2611"    # ☑
CHECKBOX_UNCHECKED = "\u2610"  # ☐


def _autofit_worklist_columns(tree, tree_name, rows_values):
    """Sizes every worklist column (except the fixed checkbox column and
    Patient Name, which stays compact by design) to exactly fit the
    widest text currently visible in it -- header included -- so a
    narrow column (e.g. #) shrinks down as small as its content allows,
    while a wide one (e.g. Last Error) grows to fit. Columns the user has
    manually resized are left exactly as they set them rather than being
    resized out from under them."""
    if not rows_values:
        return
    saved_widths = load_worklist_column_widths().get(tree_name, {}) if tree_name else {}
    try:
        f = tkfont.Font(family=FONT_FAMILY, size=max(7, round(11 * max(0.5, APP_SETTINGS.get("ui_scaling_pct", 100) / 100.0))))
    except Exception:
        return
    for idx, col in enumerate(WL_COLUMNS):
        if col in WL_NO_AUTOFIT_COLUMNS or col in saved_widths:
            continue
        max_w = f.measure(WL_HEADINGS[col])
        for vals in rows_values:
            w = f.measure(str(vals[idx]))
            if w > max_w:
                max_w = w
        # Floor is just the header's own width (+ padding) -- not the old
        # WL_WIDTHS constant, which forced every column to stay at least
        # that wide even when its actual content (e.g. a short "#" count)
        # never needed the room.
        width = min(520, max(40, max_w + 26))
        if tree.column(col, "width") != width:
            tree.column(col, width=width)
    _stretch_worklist_columns_to_fill(tree)


def _stretch_worklist_columns_to_fill(tree):
    """Hiding a column (View > Columns) used to leave dead space on the
    right, because ttk just removes that column's width from the total
    without anyone reclaiming it. This measures the gap between the
    tree's actual viewport and the sum of its visible columns' widths,
    and if there's slack, hands it out in equal shares across the
    visible, resizable columns (not the fixed checkbox column or the
    intentionally-compact Patient Name column) so the table always fills
    the space it's given."""
    tree.update_idletasks()
    avail = tree.winfo_width()
    if avail <= 1:
        return
    vis = globals().get("_column_visibility", {}).get(id(tree), {})
    visible_cols = [c for c in WL_COLUMNS if vis.get(c, True)]
    if not visible_cols:
        return
    stretchable = [c for c in visible_cols if c not in ("sel", "patient_name")]
    if not stretchable:
        return
    total_width = sum(tree.column(c, "width") for c in visible_cols)
    slack = avail - total_width - 4  # small margin so the scrollbar doesn't flicker in/out
    if slack <= 0:
        return
    share = slack // len(stretchable)
    if share <= 0:
        return
    remainder = slack - share * len(stretchable)
    for i, col in enumerate(stretchable):
        extra = share + (1 if i < remainder else 0)
        tree.column(col, width=tree.column(col, "width") + extra)


def _refresh_row_checkboxes(tree):
    """Repaints the leading checkbox glyph for every currently-visible row
    to match the tree's actual selection -- called after any selection
    change (including search re-filtering) so the checkboxes never drift
    out of sync with what's actually selected."""
    sel = set(tree.selection())
    for iid in tree.get_children(""):
        glyph = CHECKBOX_CHECKED if iid in sel else CHECKBOX_UNCHECKED
        if tree.set(iid, "sel") != glyph:
            tree.set(iid, "sel", glyph)


def _on_worklist_checkbox_click(tree, event):
    """Gmail-style: clicking the checkbox cell toggles just that row,
    on top of whatever else is already selected -- unlike a plain click
    elsewhere on the row, which (via the Treeview's default binding)
    replaces the whole selection with just that one row. Returns "break"
    so the default single-row-select binding never runs for this click."""
    if tree.identify_region(event.x, event.y) != "cell":
        return None
    col = _worklist_column_at(tree, event)
    if col != "sel":
        return None
    row_id = tree.identify_row(event.y)
    if not row_id:
        return "break"
    if row_id in set(tree.selection()):
        tree.selection_remove(row_id)
    else:
        tree.selection_add(row_id)
    return "break"


def _toggle_select_all_visible(tree):
    """Header checkbox: selects every row on screen (current search/page),
    or clears them if they're all already selected -- same as Gmail's
    top-of-list checkbox."""
    visible = tree.get_children("")
    if not visible:
        return
    if set(tree.selection()) >= set(visible):
        tree.selection_remove(*visible)
    else:
        tree.selection_add(*visible)


def _invert_tree_selection(tree):
    """Flips selection on every currently-visible row; rows outside the
    persistent selection set become selected and vice versa."""
    all_visible = set(tree.get_children(""))
    persistent = _get_persistent_selection(tree)
    currently = set(tree.selection()) | (persistent & all_visible)
    new_selection = all_visible - currently
    persistent -= all_visible
    persistent |= new_selection
    tree.selection_set(list(new_selection))
    hook = _selection_count_hooks.get(id(tree))
    if hook:
        hook(len(persistent))


def _pagination_state(tree):
    st = _tree_pagination.get(id(tree))
    if st is None:
        st = {"page": 0, "page_size": 50, "total": 0, "filter_sig": None}
        _tree_pagination[id(tree)] = st
    return st


def _update_pagination_controls(tree):
    ui = _pagination_ui.get(id(tree))
    if ui is None:
        return
    label, prev_btn, next_btn = ui
    st = _pagination_state(tree)
    total = st["total"]
    page_size = st["page_size"]
    if total == 0:
        label.configure(text="No results")
        prev_btn.configure(state="disabled")
        next_btn.configure(state="disabled")
        return
    filter_sig = st.get("filter_sig")
    search_active = bool(filter_sig and filter_sig[0])
    if search_active:
        # Search bypasses pagination entirely (see populate_tree) -- every
        # match is on screen at once, so say so instead of a page range.
        label.configure(text=f"Showing all {total} matching result{'s' if total != 1 else ''}")
        prev_btn.configure(state="disabled")
        next_btn.configure(state="disabled")
        return
    start = st["page"] * page_size + 1
    end = min(start + page_size - 1, total)
    label.configure(text=f"Showing {start}-{end} of {total}")
    prev_btn.configure(state="normal" if st["page"] > 0 else "disabled")
    next_btn.configure(state="normal" if end < total else "disabled")


def _change_tree_page(tree, delta):
    st = _pagination_state(tree)
    max_page = max(0, (st["total"] - 1) // st["page_size"]) if st["total"] else 0
    new_page = min(max(st["page"] + delta, 0), max_page)
    if new_page == st["page"]:
        return
    st["page"] = new_page
    if tree is rec_tree:
        search = rec_search_var.get()
    elif tree is push_tree:
        search = push_search_var.get()
    else:
        search = ""
    populate_tree(tree, search, force=True)


def _status_filter_for_tree(tree):
    if tree is rec_tree:
        return rec_status_filter_var.get()
    if tree is push_tree:
        return push_status_filter_var.get()
    return "All"


def _date_filter_for_tree(tree):
    if tree is rec_tree:
        return rec_date_filter_var.get()
    if tree is push_tree:
        return push_date_filter_var.get()
    return "All Time"


def _report_filter_for_tree(tree):
    if tree is rec_tree:
        return rec_report_filter_var.get()
    if tree is push_tree:
        return push_report_filter_var.get()
    return "All"


def _destination_filter_for_tree(tree):
    if tree is push_tree:
        return push_dest_filter_var.get()
    return "All"


def _row_matches_date_filter(d, date_filter):
    if date_filter == "All Time":
        return True
    t = d.get("time", "")
    if not isinstance(t, str) or not t:
        return False
    try:
        received_date = datetime.datetime.strptime(t[:10], "%Y-%m-%d").date()
    except ValueError:
        return True  # unparsable timestamp: don't hide it
    today = datetime.date.today()
    if date_filter == "Today":
        return received_date == today
    if date_filter == "Last 7 Days":
        return (today - received_date).days <= 7
    if date_filter == "Last 30 Days":
        return (today - received_date).days <= 30
    return True


def _row_search_haystack(pid, d):
    """Every field the global search engine is allowed to match against:
    Patient Name, Patient ID, Study UID, Institution, Modality, Date,
    Status, Destination, Report Status -- partial, case-insensitive."""
    report_status = "has report" if os.path.isfile(get_report_path(pid)) else "no report"
    return (
        str(pid), str(d.get("patient_name", "")), str(d.get("institution", "")),
        str(d.get("study_uid", "")), str(d.get("modality", "")), str(d.get("time", "")),
        str(d.get("status", "")), str(d.get("push_target", "")), str(d.get("source", "")),
        report_status,
    )


_all_empty_state_labels = []


def ensure_tree_empty_state(tree, message, icon=""):
    """Shows a centered guidance label over a ttk.Treeview's own parent
    frame when it has zero rows, hides it otherwise. One reusable helper
    for every tree in the app instead of a bespoke placeholder per tab."""
    container = tree.master
    display_text = f"{icon}\n{message}" if icon else message
    lbl = getattr(container, "_empty_state_label", None)
    if lbl is None or not lbl.winfo_exists():
        lbl = ctk.CTkLabel(container, text=display_text, font=get_font("small"),
                            text_color=THEME_TEXT_MUTED, justify="center")
        container._empty_state_label = lbl
        _all_empty_state_labels.append(lbl)
    else:
        lbl.configure(text=display_text)
    if len(tree.get_children()) == 0:
        lbl.place(relx=0.5, rely=0.45, anchor="center")
        lbl.lift()
    else:
        lbl.place_forget()


def _compute_filtered_rows(tree, search=""):
    """Shared filter logic: returns the list of (pid, d, status, stale,
    search_matched) tuples currently matching `tree`'s active search text
    plus its Status/Date/Report/Destination dropdown filters -- the exact
    same rules populate_tree() uses to decide what's visible, refactored
    out so 'Select All Matching Filter' (which must see every match, not
    just the current page) doesn't duplicate the logic."""
    status_filter = _status_filter_for_tree(tree)
    date_filter = _date_filter_for_tree(tree)
    report_filter = _report_filter_for_tree(tree)
    dest_filter = _destination_filter_for_tree(tree)
    search_lower = search.strip().lower()

    with data_lock:
        rows = list(patient_data.items())

    matched = []
    for pid, d in rows:
        search_matched = True
        if search_lower:
            haystack = _row_search_haystack(pid, d)
            search_matched = any(search_lower in field.lower() for field in haystack)
            if not search_matched:
                continue

        status = d.get("status", "")
        stale = is_pending_stale(d)

        if status_filter != "All":
            if status_filter == "Stale":
                if not stale:
                    continue
            elif status != status_filter:
                continue

        if not _row_matches_date_filter(d, date_filter):
            continue

        if report_filter != "All":
            has_report = os.path.isfile(get_report_path(pid))
            if report_filter == "Has Report" and not has_report:
                continue
            if report_filter == "No Report" and has_report:
                continue

        if dest_filter != "All":
            if (d.get("push_target", "") or "(none)") != dest_filter:
                continue

        matched.append((pid, d, status, stale, search_matched))
    return matched


def _compute_filtered_pids(tree):
    """PIDs only, using each tree's own current search text -- convenience
    wrapper for callers (like 'Select All Matching Filter') that don't
    already have the search string on hand."""
    if tree is rec_tree:
        search = rec_search_var.get()
    elif tree is push_tree:
        search = push_search_var.get()
    else:
        search = ""
    return [pid for pid, _d, _s, _st, _sm in _compute_filtered_rows(tree, search)]


def populate_tree(tree, search="", force=False):
    """Refresh a Treeview with current patient_data, filtered by `search`,
    plus this tree's own Status and Date dropdown filters.
    Skips the (expensive, full delete+reinsert) rebuild if patient_data
    hasn't changed and the search text is the same as last time -- this is
    what was making the UI redo full-tree rebuilds every 2s even when
    nothing had happened. Pass force=True to bypass the cache (e.g. right
    after a mode switch where tags/columns might need a fresh coat)."""
    status_filter = _status_filter_for_tree(tree)
    date_filter = _date_filter_for_tree(tree)
    report_filter = _report_filter_for_tree(tree)
    dest_filter = _destination_filter_for_tree(tree)
    # Include a minute-granularity time bucket: is_pending_stale()'s result
    # can flip purely from time passing (no patient_data change), so the
    # cache must still expire periodically -- just not every 2 seconds.
    cache_key = id(tree)
    cache_val = (_patient_data_version[0], search, status_filter, date_filter,
                 report_filter, dest_filter, int(time.time() // 60))
    if not force and _tree_render_cache.get(cache_key) == cache_val:
        return
    _tree_render_cache[cache_key] = cache_val

    search_lower = search.strip().lower()
    selected_before = set(tree.selection())
    tree.delete(*tree.get_children())

    now = datetime.datetime.now()

    matched = _compute_filtered_rows(tree, search)

    total_matched = len(matched)
    is_paginated_tree = id(tree) in _PAGINATED_TREES
    paginated = is_paginated_tree and not search_lower
    if is_paginated_tree:
        st = _pagination_state(tree)
        filter_sig = (search_lower, status_filter, date_filter, report_filter, dest_filter)
        if st["filter_sig"] != filter_sig:
            st["page"] = 0  # any filter/search change starts back at page 1
            st["filter_sig"] = filter_sig
        st["total"] = total_matched
        if paginated:
            max_page = max(0, (total_matched - 1) // st["page_size"]) if total_matched else 0
            st["page"] = min(st["page"], max_page)
            start = st["page"] * st["page_size"]
            page_rows = matched[start:start + st["page_size"]]
        else:
            # Active search: show every match at once rather than only the
            # first page's worth -- pagination only makes sense for
            # unfiltered browsing, not "find this one study".
            st["page"] = 0
            page_rows = matched
    else:
        page_rows = matched

    visible_row_index = 0
    autofit_row_values = []
    for pid, d, status, stale, search_matched in page_rows:
        zebra_tag = "even_row" if visible_row_index % 2 == 0 else "odd_row"
        visible_row_index += 1
        # 2.4 -- while push_single_patient is actively resuming a patient
        # with prior checkpoint progress toward this destination, show a
        # distinct "Resuming (sent/total)" status instead of plain "Sending".
        resume_progress = _push_resume_progress.get(pid) if status == STATUS_SENDING else None
        display_status = f"{STATUS_RESUMING_DISPLAY} ({resume_progress})" if resume_progress else status
        tags = []
        if stale:
            tags.append(STALE_HIGHLIGHT_TAG)
        elif search_lower and search_matched:
            # Only highlight for the free-text search (not the dropdown
            # filters, which already remove non-matches entirely) --
            # and don't fight with the stale-row highlight color.
            tags.append(SEARCH_MATCH_HIGHLIGHT_TAG)
        elif resume_progress:
            tags.append(f"status_{STATUS_RESUMING_DISPLAY}")
        elif status in STATUS_COLORS:
            tags.append(f"status_{status}")
        tags.append(zebra_tag)
        tags = tuple(tags)

        report_cell = "Open Report" if os.path.isfile(get_report_path(pid)) else "Create Report"
        history_cell = "Open History" if os.path.isfile(get_history_path(pid)) else "Create History"

        # C.1 / C.7: visual-only doc-transfer state suffix, Pusher tab only
        # (the Receiver side is the delivery target, not the source, so it
        # has nothing meaningful to show here). Click behavior is
        # unchanged -- this never affects what the cell *does*, only what
        # it displays.
        if tree is push_tree:
            report_path_now = get_report_path(pid)
            history_path_now = get_history_path(pid)
            report_exists = os.path.isfile(report_path_now)
            history_exists = os.path.isfile(history_path_now)
            if report_exists or history_exists:
                last_doc_err = d.get("last_doc_transfer_error", "")
                sent_at = d.get("doc_transfer_sent_at", "")
                if last_doc_err:
                    suffix = " \u2717"  # failed to send
                elif sent_at:
                    sent_epoch = None
                    try:
                        sent_epoch = datetime.datetime.strptime(sent_at, "%Y-%m-%d %H:%M:%S").timestamp()
                    except Exception:
                        pass
                    modified_since = sent_epoch is not None and any(
                        os.path.isfile(p) and os.path.getmtime(p) > sent_epoch
                        for p in (report_path_now, history_path_now)
                    )
                    suffix = " \u270e" if modified_since else " \u2713"  # modified-since-send vs sent
                else:
                    suffix = " \u25cb"  # exists locally but never confirmed delivered
                if report_exists:
                    report_cell += suffix
                if history_exists:
                    history_cell += suffix

        row_values = (
            CHECKBOX_UNCHECKED,
            pid,
            d.get("patient_name", ""),
            d.get("institution", ""),
            d.get("modality", ""),
            d.get("count", 0),
            "Open in Viewer" if d.get("count", 0) else "No Images",
            report_cell,
            history_cell,
            d.get("source", ""),
            display_status,
            d.get("time", ""),
            d.get("sent_time", ""),
            d.get("push_target", ""),
            d.get("last_error", ""),
        )
        autofit_row_values.append(row_values)

        tree.insert(
            "", "end", iid=pid,
            values=row_values,
            tags=tags,
        )

    # Restore selections that still exist. For paginated trees, selection
    # is tracked cross-page in _persistent_selection (see 1.1); for others,
    # fall back to what was selected on this same tree before the rebuild.
    _tree_name_for_autofit = "rec_tree" if tree is rec_tree else ("push_tree" if tree is push_tree else None)
    _autofit_worklist_columns(tree, _tree_name_for_autofit, autofit_row_values)
    if id(tree) in _persistent_selection:
        _reapply_persistent_selection(tree)
    else:
        for pid in selected_before:
            if tree.exists(pid):
                tree.selection_add(pid)
    _refresh_row_checkboxes(tree)

    if paginated:
        _update_pagination_controls(tree)
    if tree is rec_tree:
        for hook in _rec_results_label_hooks:
            hook()

    filters_active = bool(search_lower) or status_filter != "All" or date_filter != "All Time" or \
        report_filter != "All" or dest_filter != "All"
    if filters_active:
        empty_msg = "No studies match the current search/filters.\nTry clearing them to see everything."
    elif tree is push_tree:
        empty_msg = "Nothing queued for transfer yet.\nStudies appear here once received or imported."
    else:
        empty_msg = "No studies received yet.\nWaiting for incoming DICOM data…"
    ensure_tree_empty_state(tree, empty_msg, icon="")

# =========================================================
# BOTTOM STATUS BAR
# =========================================================
# Persistent live-info strip across the bottom of the window. Every value
# reuses state the app already tracks elsewhere (receiver_state,
# patient_data, perf_history, _last_cpu_percent, APP_VERSION) -- no new
# polling loop, just a second place some of it is displayed.

status_bar = ctk.CTkFrame(app, fg_color=THEME_SURFACE, corner_radius=10, height=32)
status_bar.pack(side="bottom", fill="x", padx=14, pady=(0, 10))
status_bar.pack_propagate(False)


def _status_bar_item(text_init="—", width=150):
    # MarqueeLabel is a drop-in stand-in for CTkLabel here: text that fits
    # renders exactly as a static label would; text that doesn't (a long
    # username, destination name, etc.) scrolls instead of getting clipped
    # or forcing the status bar to grow/overflow.
    lbl = MarqueeLabel(status_bar, text=text_init, width=width, height=20,
                        font=get_font("micro"), text_color=THEME_TEXT_MUTED)
    lbl.pack(side="left", padx=12, pady=6)
    ctk.CTkLabel(status_bar, text="│", font=get_font("micro"), text_color=THEME_HEADING_BG).pack(side="left")
    return lbl


sb_receiver_lbl = _status_bar_item("● Receiver: —", width=90)
sb_queue_lbl = _status_bar_item("Queue: —", width=65)
sb_user_lbl = _status_bar_item("User: —", width=100)
sb_cpu_lbl = _status_bar_item("CPU: —", width=55)
sb_ram_lbl = _status_bar_item("RAM: —", width=55)
sb_net_lbl = _status_bar_item("Net: —", width=80)

sb_version_lbl = MarqueeLabel(status_bar, text="", width=80, height=20,
                               font=get_font("micro"), text_color=THEME_TEXT_MUTED, anchor="e")
sb_version_lbl.pack(side="right", padx=12, pady=6)
ctk.CTkLabel(status_bar, text="│", font=get_font("micro"), text_color=THEME_HEADING_BG).pack(side="right")
sb_clock_lbl = MarqueeLabel(status_bar, text="", width=130, height=20,
                             font=get_font("micro"), text_color=THEME_TEXT_MUTED, anchor="e")
sb_clock_lbl.pack(side="right", padx=12, pady=6)


def refresh_status_bar():
    running = receiver_state.get("running")
    sb_receiver_lbl.configure(
        text=f"● Receiver: {'Running' if running else 'Stopped'}",
        text_color=THEME_SUCCESS if running else THEME_DANGER,
    )

    with data_lock:
        queue_size = sum(1 for d in patient_data.values()
                          if d.get("status") in (STATUS_PENDING, STATUS_RECEIVED, STATUS_IMPORTED, STATUS_RETRYING))
    sb_queue_lbl.configure(text=f"Queue: {queue_size}")

    role = current_view.get("role", "user") if "current_view" in globals() else "user"
    username = _audit_username() if role == "admin" else "—"
    sb_user_lbl.configure(text=f"User: {username if role == 'admin' else 'Guest (User mode)'}")

    if PSUTIL_AVAILABLE:
        sb_cpu_lbl.configure(text=f"CPU: {_last_cpu_percent[0]:.0f}%")
        try:
            sb_ram_lbl.configure(text=f"RAM: {psutil.virtual_memory().percent:.0f}%")
        except Exception:
            sb_ram_lbl.configure(text="RAM: N/A")
    else:
        sb_cpu_lbl.configure(text="CPU: N/A")
        sb_ram_lbl.configure(text="RAM: N/A")

    net_hist = perf_history.get("network_kbps")
    if net_hist:
        sb_net_lbl.configure(text=f"Net: {net_hist[-1]:.0f} KB/s")
    else:
        sb_net_lbl.configure(text="Net: —")
    # Clock is refreshed independently by _tick_status_bar_clock (1s cadence).


app.after(300, lambda: refresh_status_bar())  # one-time initial paint; periodic_refresh() covers every tick after this
app.after(300, lambda: refresh_header_status_chips())


def _tick_status_bar_clock():
    # The clock used to live in the top bar and ticked every second; that
    # felt snappier than the 2s periodic_refresh() cadence the rest of the
    # status bar uses, so it gets its own lightweight ticker here rather
    # than being downgraded to the slower cycle.
    try:
        sb_clock_lbl.configure(text=datetime.datetime.now().strftime("%a %d %b %Y  %H:%M:%S"))
    except Exception:
        pass
    app.after(1000, _tick_status_bar_clock)


app.after(200, _tick_status_bar_clock)

# =========================================================
# MAIN TABVIEW
# =========================================================

# =========================================================
# LEFT NAVIGATION SHELL (Phase 2)
# =========================================================
# Wraps the existing CTkTabview in a left nav rail instead of its default
# top segmented-button strip. Every tabview.add()/.delete()/.set()/.get()
# call elsewhere in the file keeps working exactly as before -- tabs are
# still created, named, populated, and torn down exactly as before. This
# only changes how a tab is SELECTED (nav button vs. top strip) and adds
# collapse/expand + an active-page indicator.

content_row = ctk.CTkFrame(app, fg_color="transparent")
content_row.pack(fill="both", expand=True, padx=14, pady=(0, 14))

# Right-hand inspector column lives in the same row as nav rail + tabview,
# but is only ever packed (shown) on demand -- see attach_inspector_panel()
# further down, once Patient/Study data helpers exist. Declared here so the
# nav rail / tabview keep their existing pack order untouched.
INSPECTOR_WIDTH = 320
inspector_state = {"visible": False, "animating": False}

NAV_WIDTH_EXPANDED = 210
NAV_WIDTH_COLLAPSED = 84
nav_state = {"collapsed": False}

nav_frame = ctk.CTkFrame(content_row, fg_color=THEME_SURFACE, corner_radius=12,
                          width=NAV_WIDTH_EXPANDED)
nav_frame.pack(side="left", fill="y", padx=(0, 10))
nav_frame.pack_propagate(False)

nav_collapse_btn = ctk.CTkButton(
    nav_frame, text="", width=28, height=28, corner_radius=8,
    image=get_icon("panel-left-close", size=15, color=THEME_TEXT_MUTED),
    fg_color=THEME_HEADING_BG, hover_color=THEME_NEUTRAL_BTN_HOVER,
    text_color=THEME_TEXT_MUTED,
)
nav_collapse_btn.pack(anchor="e", padx=8, pady=(8, 0))

nav_scroll = ctk.CTkScrollableFrame(nav_frame, fg_color="transparent")
nav_scroll.pack(fill="both", expand=True, padx=6, pady=(4, 6))

nav_buttons = {}       # tab_name -> CTkButton
nav_button_order = []  # registration order (kept for compatibility; display order is section-driven)
nav_section_headers = {}  # section title -> CTkLabel

# Logical grouping for the left nav, addressing the same "categorize
# settings, avoid clutter" goal the spec's SETTINGS section describes --
# this app has no single settings dump to reorganize, it has 13 admin-only
# tabs that were all flat peers. Grouping them is the equivalent move.
# No tab's content, name, or behavior changes -- purely a nav-ordering
# and visual-grouping layer.
NAV_SECTION_ORDER = ["Overview", "Configuration", "Reporting & Logs", "Tools"]
NAV_SECTION_FOR_TAB = {
    "Dashboard": "Overview",
    "Receiver": "Overview",
    "Pusher": "Overview",
    "Admin Dashboard": "Overview",
    "PACS Health": "Overview",
    "Performance": "Overview",
    "Bandwidth": "Configuration",
    "Destinations": "Configuration",
    "Routing Rules": "Configuration",
    "SOP Classes": "Configuration",
    "LDAP / AD": "Configuration",
    "Backup": "Configuration",
    "Export": "Configuration",
    "Reports": "Reporting & Logs",
    "Logs": "Reporting & Logs",
    "Query/Retrieve": "Tools",
    "Settings": "Tools",
}

# Explicit tab -> icon mapping (replaces the old "icon is the tab name's
# first word" convention, which broke for multi-word tab names once the
# emoji prefixes were removed).
NAV_ICON_FOR_TAB = {
    "Dashboard": "house",
    "Receiver": "inbox",
    "Pusher": "send",
    "Admin Dashboard": "shield",
    "PACS Health": "activity",
    "Performance": "gauge",
    "Bandwidth": "wifi",
    "Destinations": "server",
    "Routing Rules": "route",
    "SOP Classes": "file-text",
    "LDAP / AD": "users",
    "Backup": "archive",
    "Export": "download",
    "Reports": "chart-column",
    "Logs": "scroll-text",
    "Query/Retrieve": "search",
    "Settings": "settings-2",
}


def _nav_icon_and_label(tab_name):
    """Returns (lucide_icon_name, label) for a tab. Falls back to a
    generic icon for any tab that isn't in NAV_ICON_FOR_TAB rather than
    guessing from the tab name's text."""
    return NAV_ICON_FOR_TAB.get(tab_name, "layers"), tab_name


def _nav_button_text(tab_name):
    if nav_state["collapsed"]:
        return ""
    _icon, label = _nav_icon_and_label(tab_name)
    return f"  {label}"


def _on_nav_click(tab_name):
    try:
        tabview.set(tab_name)
    except Exception:
        log_exception(f"Nav click failed to switch to {tab_name}")
    _sync_nav_highlight()
    if APP_SETTINGS.get("restore_previous_session"):
        _update_app_setting("_last_active_tab", tab_name)


NAV_BTN_WIDTH_EXPANDED = 180
NAV_BTN_WIDTH_COLLAPSED = 56


def add_nav_button(tab_name):
    """Registers (or re-shows) the nav entry for an existing tabview tab.
    Safe to call more than once for the same name."""
    if tab_name in nav_buttons:
        return
    icon_name, _label = _nav_icon_and_label(tab_name)
    collapsed = nav_state["collapsed"]
    btn = ctk.CTkButton(
        nav_scroll, text=_nav_button_text(tab_name),
        anchor="center" if collapsed else "w",
        width=NAV_BTN_WIDTH_COLLAPSED if collapsed else NAV_BTN_WIDTH_EXPANDED,
        image=get_icon(icon_name, size=17, color=THEME_TEXT_MUTED),
        compound="left",
        fg_color="transparent", hover_color=THEME_HEADING_BG,
        text_color=THEME_TEXT_MUTED, corner_radius=8, height=get_nav_button_height(),
        font=get_font("small", "bold"),
        command=lambda n=tab_name: _on_nav_click(n),
    )
    btn._nav_icon_name = icon_name
    nav_buttons[tab_name] = btn
    nav_button_order.append(tab_name)
    _layout_nav_buttons()
    _sync_nav_highlight()


def remove_nav_button(tab_name):
    btn = nav_buttons.pop(tab_name, None)
    if btn is not None:
        btn.destroy()
    if tab_name in nav_button_order:
        nav_button_order.remove(tab_name)
    _layout_nav_buttons()


def _get_or_make_section_header(section_title):
    header = nav_section_headers.get(section_title)
    if header is None:
        header = ctk.CTkLabel(
            nav_scroll, text=section_title.upper(), anchor="w",
            font=get_font("micro", "bold"), text_color=THEME_TEXT_MUTED,
        )
        nav_section_headers[section_title] = header
    return header


def _layout_nav_buttons():
    """Renders nav_buttons grouped under NAV_SECTION_ORDER instead of flat
    registration order. A section header only appears when at least one
    of its tabs is currently registered (so it disappears cleanly in User
    mode, where only Overview's first 3 items exist)."""
    for header in nav_section_headers.values():
        header.pack_forget()
    for btn in nav_buttons.values():
        btn.pack_forget()

    if nav_state["collapsed"]:
        # Collapsed = icons only; section headers add no value at that
        # width, so skip them and just list every registered tab in a
        # stable, section-ordered sequence.
        for section in NAV_SECTION_ORDER:
            for name in NAV_SECTION_FOR_TAB:
                if NAV_SECTION_FOR_TAB[name] == section and name in nav_buttons:
                    nav_buttons[name].pack(fill="x", pady=2, padx=2)
        return

    for section in NAV_SECTION_ORDER:
        members = [n for n in NAV_SECTION_FOR_TAB if NAV_SECTION_FOR_TAB[n] == section and n in nav_buttons]
        if not members:
            continue
        header = _get_or_make_section_header(section)
        header.pack(fill="x", padx=8, pady=(12, 2))
        # Preserve each tab's natural creation order within its section
        # (dict insertion order of NAV_SECTION_FOR_TAB matches the app's
        # existing tab order) rather than nav_button_order (registration
        # order, which can differ once Admin tabs are added after the
        # always-visible 3).
        for name in NAV_SECTION_FOR_TAB:
            if name in members:
                nav_buttons[name].pack(fill="x", pady=2, padx=2)


def _sync_nav_highlight():
    try:
        current = tabview.get()
    except Exception:
        return
    _update_page_title()
    for name, btn in nav_buttons.items():
        icon_name = getattr(btn, "_nav_icon_name", "layers")
        if name == current:
            btn.configure(fg_color=THEME_ACCENT, text_color="#ffffff",
                           image=get_icon(icon_name, size=17, color="#ffffff"))
        else:
            btn.configure(fg_color="transparent", text_color=THEME_TEXT_MUTED,
                           image=get_icon(icon_name, size=17, color=THEME_TEXT_MUTED))



_nav_anim_state = {"animating": False}


def _animate_nav_width(target_width, step=0, steps=8):
    # Enterprise UI requirement: no animations/transitions anywhere.
    # Snap directly to the target width instead of stepping toward it.
    try:
        nav_frame.configure(width=target_width)
    except Exception:
        return
    _nav_anim_state["animating"] = False


def _toggle_nav_collapse():
    nav_state["collapsed"] = not nav_state["collapsed"]
    collapsed = nav_state["collapsed"]
    target = NAV_WIDTH_COLLAPSED if collapsed else NAV_WIDTH_EXPANDED
    _nav_anim_state["animating"] = True
    _animate_nav_width(target)
    for name, btn in nav_buttons.items():
        btn.configure(
            text=_nav_button_text(name),
            anchor="center" if collapsed else "w",
            width=NAV_BTN_WIDTH_COLLAPSED if collapsed else NAV_BTN_WIDTH_EXPANDED,
        )
    nav_collapse_btn.configure(image=get_icon(
        "panel-left-open" if collapsed else "panel-left-close",
        size=15, color=THEME_TEXT_MUTED))
    _layout_nav_buttons()


nav_collapse_btn.configure(command=_toggle_nav_collapse)

# =========================================================
# COMMAND PALETTE (Ctrl+K) -- Raycast/Spotlight-style quick access
# =========================================================
# Reuses the exact same nav_buttons registry the left nav rail is built
# from, so "every page reachable from the palette" is automatic -- no
# second list to keep in sync. Quick actions are intentionally limited to
# safe, idempotent operations (refresh/navigate/toggle); nothing here
# starts/stops the receiver or touches destinations, to avoid a stray
# keystroke triggering something irreversible.

command_palette_state = {"window": None}


def _command_palette_quick_actions():
    actions = [
        ("Refresh Now", lambda: (populate_tree(rec_tree, rec_search_var.get(), force=True),
                                     populate_tree(push_tree, push_search_var.get(), force=True),
                                     refresh_status_bar(), refresh_header_status_chips())),
        ("Open Notification Center", _toggle_notification_panel),
        ("Toggle Navigation Collapse", _toggle_nav_collapse),
    ]
    return actions


def _command_palette_entries(query):
    """Returns a flat list of (label, run_callable) -- first every nav
    page whose label loosely matches `query` (case-insensitive substring,
    ordered by where the match occurs), then the quick actions."""
    query = (query or "").strip().lower()
    page_matches = []
    for tab_name in nav_buttons.keys():
        _icon, label = _nav_icon_and_label(tab_name)
        haystack = f"{NAV_SECTION_FOR_TAB.get(tab_name, '')} {label}".lower()
        if not query or query in haystack:
            pos = haystack.find(query) if query else 0
            page_matches.append((pos, f"{_icon}  {label}", tab_name))
    page_matches.sort(key=lambda t: t[0])
    entries = [(label, (lambda n=tab_name: _on_nav_click(n))) for _pos, label, tab_name in page_matches]

    for label, fn in _command_palette_quick_actions():
        if not query or query in label.lower():
            entries.append((label, fn))
    return entries


def _close_command_palette():
    win = command_palette_state.get("window")
    if win is not None and win.winfo_exists():
        win.destroy()
    command_palette_state["window"] = None


def _run_command_palette_entry(entry):
    _close_command_palette()
    try:
        entry()
    except Exception:
        log_exception("Command palette action failed")


def _rebuild_command_palette_results(win, query):
    for child in win.results_frame.winfo_children():
        child.destroy()
    entries = _command_palette_entries(query)[:12]
    if not entries:
        ctk.CTkLabel(win.results_frame, text="No matching pages or actions.",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(pady=16)
        return
    for label, fn in entries:
        btn = ctk.CTkButton(
            win.results_frame, text=label, anchor="w", height=34, corner_radius=8,
            fg_color="transparent", hover_color=THEME_HEADING_BG, text_color=THEME_TEXT,
            font=get_font("body"), command=lambda f=fn: _run_command_palette_entry(f),
        )
        btn.pack(fill="x", padx=6, pady=2)


def _fade_in_window(win, steps=8, delay=12, target_alpha=1.0):
    # Enterprise UI requirement: no animations/transitions anywhere.
    # Show the window immediately at full opacity instead of fading in.
    try:
        win.attributes("-alpha", target_alpha)
    except Exception:
        pass


def toggle_command_palette(_evt=None):
    existing = command_palette_state.get("window")
    if existing is not None and existing.winfo_exists():
        _close_command_palette()
        return "break"

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.overrideredirect(True)
    win.attributes("-topmost", True)
    win.configure(fg_color=THEME_SURFACE)

    app.update_idletasks()
    w, h = 560, 420
    x = app.winfo_rootx() + (app.winfo_width() - w) // 2
    y = app.winfo_rooty() + 90
    win.geometry(f"{w}x{h}+{x}+{y}")

    shell = ctk.CTkFrame(win, fg_color=THEME_SURFACE, corner_radius=12,
                          border_width=1, border_color=THEME_HEADING_BG)
    shell.pack(fill="both", expand=True, padx=1, pady=1)

    query_var = ctk.StringVar(value="")
    entry = ctk.CTkEntry(
        shell, textvariable=query_var, height=44, corner_radius=10,
        placeholder_text="Type a page or action…  (Esc to close)",
        font=get_font("section"), fg_color=THEME_HEADING_BG, border_width=0,
    )
    entry.pack(fill="x", padx=14, pady=(14, 8))

    results = ctk.CTkScrollableFrame(shell, fg_color="transparent")
    results.pack(fill="both", expand=True, padx=8, pady=(0, 12))
    win.results_frame = results

    entry.bind("<KeyRelease>", lambda _e: _rebuild_command_palette_results(win, query_var.get()))
    entry.bind("<Escape>", lambda _e: _close_command_palette())
    entry.bind("<Return>", lambda _e: (
        _run_command_palette_entry(_command_palette_entries(query_var.get())[0][1])
        if _command_palette_entries(query_var.get()) else None
    ))
    win.bind("<FocusOut>", lambda _e: app.after(120, lambda: (
        _close_command_palette() if win.winfo_exists() and app.focus_get() is None else None
    )))

    command_palette_state["window"] = win
    _rebuild_command_palette_results(win, "")
    entry.focus_set()
    _fade_in_window(win)
    return "break"


app.bind("<Control-k>", toggle_command_palette)
app.bind("<Control-K>", toggle_command_palette)

content_col = ctk.CTkFrame(content_row, fg_color="transparent")
content_col.pack(side="left", fill="both", expand=True)

page_title_bar = ctk.CTkFrame(content_col, fg_color=THEME_SURFACE, corner_radius=12, height=40)
page_title_bar.pack(fill="x", pady=(0, 4))
page_title_bar.pack_propagate(False)
page_title_icon_lbl = ctk.CTkLabel(page_title_bar, text="")
page_title_icon_lbl.pack(side="left", padx=(16, 8))
page_title_lbl = ctk.CTkLabel(page_title_bar, text="", font=get_font("section", "bold"), text_color=THEME_TEXT)
page_title_lbl.pack(side="left")

tabview = ctk.CTkTabview(
    content_col, anchor="nw",
    fg_color=THEME_SURFACE,
    segmented_button_fg_color=THEME_HEADING_BG,
    segmented_button_selected_color=THEME_ACCENT,
    segmented_button_selected_hover_color=THEME_ACCENT_HOVER,
    segmented_button_unselected_hover_color=THEME_SEGMENTED_HOVER,
    text_color=THEME_TEXT,
    corner_radius=12,
)
tabview.pack(side="left", fill="both", expand=True)


def _update_page_title():
    try:
        current = tabview.get()
    except Exception:
        return
    icon_name, label = _nav_icon_and_label(current)
    page_title_icon_lbl.configure(image=get_icon(icon_name, size=17, color=THEME_TEXT))
    page_title_lbl.configure(text=label or current)


def _hide_tabview_strip():
    for forget in ("grid_forget", "pack_forget", "place_forget"):
        try:
            getattr(tabview._segmented_button, forget)()
        except Exception:
            pass
    # Some CTkTabview versions also reserve a row/column for the strip even
    # once it's unmapped, leaving an empty gap where it used to be.
    try:
        tabview.grid_rowconfigure(0, weight=0, minsize=0)
    except Exception:
        pass


_hide_tabview_strip()

# CTkTabview.add() re-packs its own segmented button internally every time
# a new tab is added (Receiver/Pusher here, and later all the admin-only
# tabs via build_admin_only_tabs()), so a one-time pack_forget() at
# creation isn't enough -- the top strip would silently reappear the
# moment the second tab got added. Wrapping .add() re-hides it every time
# instead, since the left nav rail is the only tab switcher we want.
_tabview_add_original = tabview.add


def _tabview_add_and_stay_hidden(name):
    result = _tabview_add_original(name)
    _hide_tabview_strip()
    return result


tabview.add = _tabview_add_and_stay_hidden

# Receiver + Pusher are visible in BOTH modes, so they're created once, here,
# unchanged. The Admin-only tabs (Dashboard, Query/Retrieve, Destinations,
# Routing Rules, SOP Classes, Logs) are NOT created here -- they're added by
# build_admin_only_tabs() (defined further down, once all their callbacks
# exist) and removed again by tear_down_admin_only_tabs() on mode switch, so
# there is exactly one copy of each tab's construction code, per spec.
# =========================================================
# ENTERPRISE DASHBOARD TAB (landing page -- visible in BOTH modes,
# built once here just like Receiver/Pusher below). Populated by
# refresh_home_dashboard(), defined further down once all the data
# sources (push_job, receiver_state, get_push_throughput_eta, etc.)
# exist -- Python only needs those names to exist by the time the
# function actually RUNS, not by the time it's defined.)
# =========================================================

tab_dashboard_home = tabview.add("Dashboard")
add_nav_button("Dashboard")

_splash_step(0.35, "Loading dashboard…")

dash_home_scroll = ctk.CTkScrollableFrame(tab_dashboard_home, fg_color="transparent")
dash_home_scroll.pack(fill="both", expand=True, padx=4, pady=4)


def _home_section_label(parent, text):
    ctk.CTkLabel(parent, text=text, font=get_font("section", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=10, pady=(18, 6))


def _home_cards_row(parent):
    return make_card_row(parent)


def _home_stat_card(parent, title, value="—", on_click=None):
    card = make_card(parent)
    add_card_to_row(parent, card)
    title_lbl = ctk.CTkLabel(card.body, text=title, font=get_font("micro"),
                 text_color=THEME_TEXT_MUTED)
    title_lbl.pack(anchor="w", pady=(4, 0))
    # width=1 + fill="x": the card's own layout (grid, uniform column) decides
    # the real width; this just tracks it and scrolls only if the value
    # doesn't fit in whatever room the card ends up with.
    val_lbl = MarqueeLabel(card.body, text=value, width=1, height=32,
                            font=get_font("title", "bold"), text_color=THEME_TEXT,
                            canvas_bg=THEME_HEADING_BG)
    val_lbl.pack(anchor="w", fill="x", pady=(2, 0))
    if on_click:
        # 4.1 -- clickable stat cards jump to the relevant tab with
        # filters pre-applied, instead of being read-only numbers. Bind on
        # the card frame + both labels so the whole tile is clickable, not
        # just a sliver of text.
        for widget in (card, card.body, title_lbl, val_lbl):
            try:
                widget.configure(cursor="hand2")
            except Exception:
                pass
            widget.bind("<Button-1>", lambda _e, cb=on_click: cb())
    return val_lbl


def _goto_pusher_with_status_filter(status):
    """4.1 helper: switches to the Pusher tab and pre-applies a status
    filter, reusing the exact same filter var + refresh path the Pusher
    tab's own dropdown uses."""
    _on_nav_click("Pusher")
    push_status_filter_var.set(status)
    refresh_worklists()


def _goto_receiver_tab():
    _on_nav_click("Receiver")


# ---- Receiver stats ----
_home_section_label(dash_home_scroll, "Receiver")
home_recv_row1 = _home_cards_row(dash_home_scroll)
dash_home_recv_status = _home_stat_card(home_recv_row1, "Receiver Status", "● Stopped")
dash_home_recv_ae = _home_stat_card(home_recv_row1, "AE Title")
dash_home_recv_port = _home_stat_card(home_recv_row1, "Listening Port")
dash_home_recv_ip = _home_stat_card(home_recv_row1, "Current IP")
dash_home_recv_assoc = _home_stat_card(home_recv_row1, "Active Associations", "0")

home_recv_row2 = _home_cards_row(dash_home_scroll)
dash_home_studies_recv = _home_stat_card(home_recv_row2, "Studies Received Today", "0", on_click=_goto_receiver_tab)
dash_home_images_recv = _home_stat_card(home_recv_row2, "Images Received Today", "0", on_click=_goto_receiver_tab)
dash_home_docs_recv = _home_stat_card(home_recv_row2, "Documents Received Today", "0", on_click=_goto_receiver_tab)
dash_home_total_patients = _home_stat_card(home_recv_row2, "Total Patients", "0", on_click=_goto_receiver_tab)
dash_home_recv_queue = _home_stat_card(home_recv_row2, "Current Queue Size", "0", on_click=_goto_receiver_tab)

# ---- Pusher stats ----
_home_section_label(dash_home_scroll, "Pusher")
home_push_row1 = _home_cards_row(dash_home_scroll)
dash_home_push_status = _home_stat_card(home_push_row1, "Push Status", "Idle",
                                          on_click=lambda: _on_nav_click("Pusher"))
dash_home_dest_status = _home_stat_card(home_push_row1, "Destination Status", "Unknown",
                                          on_click=lambda: _on_nav_click("Pusher"))
dash_home_images_sent = _home_stat_card(home_push_row1, "Images Sent Today", "0",
                                          on_click=lambda: _goto_pusher_with_status_filter(STATUS_SENT))
dash_home_studies_sent = _home_stat_card(home_push_row1, "Studies Sent Today", "0",
                                          on_click=lambda: _goto_pusher_with_status_filter(STATUS_SENT))
dash_home_failed_transfers = _home_stat_card(home_push_row1, "Failed Transfers", "0",
                                          on_click=lambda: _goto_pusher_with_status_filter(STATUS_FAILED))

home_push_row2 = _home_cards_row(dash_home_scroll)
dash_home_pending_queue = _home_stat_card(home_push_row2, "Pending Queue", "0")
dash_home_retry_queue = _home_stat_card(home_push_row2, "Retry Queue", "0")
dash_home_throughput_img = _home_stat_card(home_push_row2, "Throughput (img/sec)", "0.0")
dash_home_throughput_mb = _home_stat_card(home_push_row2, "Throughput (MB/sec)", "0.0")

# ---- System stats ----
_home_section_label(dash_home_scroll, "System")
home_sys_row1 = _home_cards_row(dash_home_scroll)
dash_home_cpu = _home_stat_card(home_sys_row1, "CPU Usage", "—")
dash_home_ram = _home_stat_card(home_sys_row1, "RAM Usage", "—")
dash_home_disk = _home_stat_card(home_sys_row1, "Disk Usage", "—")
dash_home_disk_free = _home_stat_card(home_sys_row1, "Remaining Disk Space", "—")
dash_home_uptime = _home_stat_card(home_sys_row1, "Application Uptime", "—")

home_sys_row2 = _home_cards_row(dash_home_scroll)
dash_home_net_up = _home_stat_card(home_sys_row2, "Network Upload Speed", "—")
dash_home_net_down = _home_stat_card(home_sys_row2, "Network Download Speed", "—")

if not PSUTIL_AVAILABLE:
    ctk.CTkLabel(
        dash_home_scroll,
        text="psutil is not installed -- CPU/RAM/network speed metrics will show as N/A. "
             "Install it with: pip install psutil",
        image=get_icon("triangle-alert", size=14, color=THEME_WARNING), compound="left",
        font=get_font("small"), text_color=THEME_WARNING,
    ).pack(anchor="w", padx=16, pady=(0, 4))

# ---- Live graphs ----
_home_section_label(dash_home_scroll, "Live Graphs (auto-refreshing)")
home_graphs_row = ctk.CTkScrollableFrame(dash_home_scroll, fg_color="transparent",
                                          orientation="horizontal", height=200)
# NOTE: home_graphs_row.pack() is called further below, AFTER
# dash_graph_range_row is packed. CTkScrollableFrame.pack() doesn't pack
# the CTkScrollableFrame object itself -- it redirects to an internal
# wrapper frame -- so Tk never sees a packed slave with home_graphs_row's
# own widget path. That means `dash_graph_range_row.pack(before=home_graphs_row)`
# always raised "window ... isn't packed", since there was nothing under
# that exact path for Tk's pack manager to insert before. Packing the two
# rows in the same top-to-bottom order they should visually appear in --
# range selector first, then the graphs row -- gets the identical layout
# without relying on before=/after= against a CTkScrollableFrame.

dash_graph_range_row = ctk.CTkFrame(dash_home_scroll, fg_color="transparent")
dash_graph_range_row.pack(fill="x", padx=6)
ctk.CTkLabel(dash_graph_range_row, text="Trend range:", font=get_font("small"),
             text_color=THEME_TEXT_MUTED).pack(side="left", padx=(4, 6))
dash_graph_range_var = ctk.StringVar(value="Live")
# "Live" keeps the existing ~2-minute in-session dash_history view (also
# the only option for Queue Size / Network Usage, which have no daily
# equivalent); 7d/30d switch studies/failed-transfers sparklines to the
# persisted daily_stats_history.json summary instead (4.2).
ctk.CTkSegmentedButton(dash_graph_range_row, values=["Live", "7d", "30d"],
                       variable=dash_graph_range_var,
                       command=lambda _v: refresh_admin_dashboard()).pack(side="left")

home_graphs_row.pack(fill="x", padx=6, pady=(0, 12))

DASH_GRAPH_COLORS = {
    "studies_received": "#2f8eff",
    "studies_sent": "#2ecc71",
    "failed_transfers": "#f04747",
    "queue_size": "#f1c40f",
    "network_kbps": "#9b59b6",
    "docs_delivered": "#1abc9c",
    "doc_failure_rate": "#e74c3c",
}
dash_graph_canvases = {}


def _make_sparkline(parent, key, title):
    card = make_card(parent)
    card.pack(side="left", padx=5, pady=5, fill="both", expand=True)
    ctk.CTkLabel(card.body, text=title, font=get_font("micro"),
                 text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 0))
    canvas = ctk.CTkCanvas(card.body, width=180, height=110, bg=THEME_SURFACE, highlightthickness=0)
    canvas.pack(pady=(4, 2))
    dash_graph_canvases[key] = canvas


_make_sparkline(home_graphs_row, "studies_received", "Studies Received")
_make_sparkline(home_graphs_row, "studies_sent", "Studies Sent")
_make_sparkline(home_graphs_row, "failed_transfers", "Failed Transfers")
_make_sparkline(home_graphs_row, "queue_size", "Queue Size")
_make_sparkline(home_graphs_row, "network_kbps", "Network Usage (KB/s)")
# B.5: docs-delivered-per-day and doc-transfer failure-rate, last 7/30 days.
_make_sparkline(home_graphs_row, "docs_delivered", "Docs Delivered/Day")
_make_sparkline(home_graphs_row, "doc_failure_rate", "Doc Failure Rate %")


def _draw_sparkline(canvas, values, color):
    """Lightweight dependency-free line chart drawn straight onto a
    tkinter Canvas -- avoids pulling in matplotlib just for small
    auto-refreshing trend lines. Skipped entirely when Settings >
    Performance > Disable Graphs is on, since repeatedly redrawing these
    on every refresh tick is one of the more avoidable CPU costs on
    slower machines."""
    try:
        canvas.delete("all")
        w = int(canvas.cget("width"))
        h = int(canvas.cget("height"))
        if APP_SETTINGS.get("disable_graphs"):
            canvas.create_text(w // 2, h // 2, text="Graphs disabled",
                                fill=THEME_TEXT_MUTED, font=("TkDefaultFont", 9), anchor="center")
            return
        pad = 6
        if not values:
            return
        vmax = max(values) if max(values) > 0 else 1
        n = len(values)
        if n == 1:
            values = [values[0], values[0]]
            n = 2
        step_x = (w - 2 * pad) / max(n - 1, 1)
        points = []
        for i, v in enumerate(values):
            x = pad + i * step_x
            y = h - pad - ((v / vmax) * (h - 2 * pad))
            points.extend([x, y])
        if len(points) >= 4:
            canvas.create_line(*points, fill=color, width=2, smooth=True)
        canvas.create_text(w - pad, pad, text=f"{values[-1]:.1f}" if isinstance(values[-1], float) else str(values[-1]),
                            fill=color, anchor="ne", font=("TkDefaultFont", 9, "bold"))
    except Exception:
        log_exception("Failed to draw dashboard sparkline")


tab_receiver = tabview.add("Receiver")
add_nav_button("Receiver")
tab_pusher = tabview.add("Pusher")
add_nav_button("Pusher")

# Same pattern as the Receiver tab: the whole Pusher page -- monitoring
# strip, config, offline queue, and worklist -- lives inside one
# CTkScrollableFrame so it all scrolls together as a single page.
tab_pusher_scroll = ctk.CTkScrollableFrame(tab_pusher, fg_color="transparent")
tab_pusher_scroll.pack(fill="both", expand=True)
tabview.set("Dashboard")
if APP_SETTINGS.get("restore_previous_session"):
    _last_tab = APP_SETTINGS.get("_last_active_tab")
    if _last_tab in ("Dashboard", "Receiver", "Pusher"):
        try:
            tabview.set(_last_tab)
        except Exception:
            log_exception("Failed to restore previous session tab")
_sync_nav_highlight()

# The entire Receiver tab -- monitoring cards, config/buttons, search bar,
# and the patient list -- lives inside one CTkScrollableFrame, so the whole
# page scrolls as a single unit rather than relying on the inner tree's own
# scrollbar. Every widget below that used to parent directly off
# `tab_receiver` now parents off `tab_receiver_scroll` instead.
tab_receiver_scroll = ctk.CTkScrollableFrame(tab_receiver, fg_color="transparent")
tab_receiver_scroll.pack(fill="both", expand=True)

# ---- Pusher monitoring strip ----
# Purely a presentation layer over state that already exists: push_job
# (running/sent/attempted/total), get_offline_queue_summary() (retry
# queue), destination_health_cache (destination status, populated by the
# existing PACS Health / Dashboard background checker), and daily_stats
# (today's counters). No new tracking, no new background thread.

push_monitor_row = make_card_row(tab_pusher_scroll, pady=(10, 0))
push_monitor_row.pack_configure(padx=10)

push_monitor_status_card = make_card(push_monitor_row)
add_card_to_row(push_monitor_row, push_monitor_status_card)
ctk.CTkLabel(push_monitor_status_card.body, text="Push Status",
             font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 4))
push_monitor_status_badge = make_status_badge(push_monitor_status_card.body, "Idle", kind="pending")
push_monitor_status_badge.pack(anchor="w")

push_monitor_dest_card = make_card(push_monitor_row)
add_card_to_row(push_monitor_row, push_monitor_dest_card)
ctk.CTkLabel(push_monitor_dest_card.body, text="Destination Status",
             font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 4))
push_monitor_dest_badge = make_status_badge(push_monitor_dest_card.body, "Unknown", kind="pending")
push_monitor_dest_badge.pack(anchor="w")


def _push_monitor_stat(title, value="—"):
    card = make_card(push_monitor_row)
    add_card_to_row(push_monitor_row, card)
    ctk.CTkLabel(card.body, text=title, font=get_font("micro"),
                 text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 0))
    val_lbl = ctk.CTkLabel(card.body, text=value, font=get_font("title", "bold"),
                            text_color=THEME_TEXT)
    val_lbl.pack(anchor="w", pady=(2, 0))
    return val_lbl


push_monitor_pending_lbl = _push_monitor_stat("Pending Queue", "0")
push_monitor_retry_lbl = _push_monitor_stat("Retry Queue", "0")
push_monitor_failed_today_lbl = _push_monitor_stat("Failed Transfers Today", "0")
push_monitor_sent_today_lbl = _push_monitor_stat("Images Sent Today", "0")
push_monitor_eta_lbl = _push_monitor_stat("Throughput / ETA", "—")


_push_config_auto_collapse_done = {"value": False}


def refresh_pusher_monitoring():
    if push_job.get("running"):
        if _push_resume_progress:
            # At least one patient in this job is resuming from a prior
            # checkpoint -- surface that distinctly rather than the
            # generic "Sending" (2.4). Picks whichever is furthest along
            # if several are resuming in parallel worker threads.
            some_pid, some_progress = next(iter(_push_resume_progress.items()))
            push_monitor_status_badge.update_status(f"Resuming ({some_progress})", "resuming")
        else:
            push_monitor_status_badge.update_status("Sending", "sending")
        if not _push_config_collapsed["value"] and not _push_config_auto_collapse_done["value"]:
            # Same idea as the Receiver tab: once a push is actually underway,
            # free up the vertical space for the worklist below. One-shot --
            # won't re-collapse if the user manually reopens it.
            _push_config_auto_collapse_done["value"] = True
            _toggle_push_config_panel()
    else:
        push_monitor_status_badge.update_status("Idle", "pending")

    dest = get_default_destination()
    if dest and dest["name"] in destination_health_cache:
        health = destination_health_cache[dest["name"]]
        if health["online"]:
            push_monitor_dest_badge.update_status(f"Online ({dest['name']})", "connected")
        else:
            push_monitor_dest_badge.update_status(f"Offline ({dest['name']})", "offline")
    elif dest:
        push_monitor_dest_badge.update_status(f"Checking… ({dest['name']})", "pending")
    else:
        push_monitor_dest_badge.update_status("No destination configured", "pending")

    with data_lock:
        pending_count = sum(1 for d in patient_data.values() if d.get("status") == STATUS_PENDING)
        retry_count_total = sum(1 for d in patient_data.values() if d.get("status") == STATUS_RETRYING)
    offline_queue_size, _oldest, _next_retry, _interval = get_offline_queue_summary()
    push_monitor_pending_lbl.configure(text=str(pending_count))
    push_monitor_retry_lbl.configure(text=str(retry_count_total + offline_queue_size))

    with _dashboard_stats_lock:
        failed_today = daily_stats["failed_transfers"]
        images_sent_today = daily_stats["images_sent"]
    push_monitor_failed_today_lbl.configure(text=str(failed_today))
    push_monitor_sent_today_lbl.configure(text=str(images_sent_today))

    img_rate, eta_seconds = get_push_throughput_eta()
    if push_job.get("running") and img_rate > 0:
        eta_txt = f"ETA {int(eta_seconds)}s" if eta_seconds is not None else "ETA —"
        push_monitor_eta_lbl.configure(text=f"{img_rate:.2f} img/s · {eta_txt}")
    else:
        push_monitor_eta_lbl.configure(text="—")


# ---- Offline Queue panel ----
# Shown at the very top of the Pusher tab since this is where queued
# push failures naturally belong. Populated by refresh_offline_queue_ui()
# (called from periodic_refresh, cheap -- reads offline_queue.json state
# already cached in memory, no network I/O).
offlineq_header = ctk.CTkFrame(tab_pusher_scroll, fg_color=THEME_HEADING_BG, corner_radius=10)
offlineq_header.pack(fill="x", padx=10, pady=(10, 0))

_offlineq_collapsed = {"value": False}


def _toggle_offlineq_panel():
    _offlineq_collapsed["value"] = not _offlineq_collapsed["value"]
    if _offlineq_collapsed["value"]:
        offlineq_outer.pack_forget()
        offlineq_toggle_btn.configure(text="▸  Offline Queue  (click to expand)")
    else:
        offlineq_outer.pack(fill="x", padx=10, pady=(6, 0), after=offlineq_header)
        offlineq_toggle_btn.configure(text="▾  Offline Queue  (click to collapse)")


offlineq_toggle_btn = ctk.CTkButton(
    offlineq_header, text="▾  Offline Queue  (click to collapse)",
    font=get_font("body", "bold"), anchor="w", corner_radius=8,
    fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER, text_color="#ffffff",
    command=lambda: _toggle_offlineq_panel())
offlineq_toggle_btn.pack(fill="x", padx=8, pady=6)

offlineq_outer = ctk.CTkFrame(tab_pusher_scroll)
offlineq_outer.pack(fill="x", padx=10, pady=(6, 0))

offlineq_top = ctk.CTkFrame(offlineq_outer, fg_color="transparent")
offlineq_top.pack(fill="x", padx=8, pady=(8, 4))
ctk.CTkLabel(offlineq_top, text="Offline Queue", font=get_font("section", "bold")).pack(side="left")
offlineq_retry_all_btn = ctk.CTkButton(offlineq_top, text="Retry Due Items Now", width=180)
offlineq_retry_all_btn.pack(side="right", padx=4)

offlineq_threshold_var = ctk.StringVar(
    value=str(load_notifications_config()["thresholds"].get("offline_queue_size", 20)))


def _save_offlineq_threshold(*_a):
    try:
        value = int(offlineq_threshold_var.get())
    except (TypeError, ValueError):
        return
    cfg = load_notifications_config()
    if cfg["thresholds"].get("offline_queue_size") != value:
        cfg["thresholds"]["offline_queue_size"] = value
        save_notifications_config(cfg)


offlineq_threshold_entry = ctk.CTkEntry(offlineq_top, width=50, textvariable=offlineq_threshold_var)
offlineq_threshold_entry.pack(side="right", padx=(4, 10))
ctk.CTkLabel(offlineq_top, text="Alert at queue size:", font=get_font("small"),
             text_color=THEME_TEXT_MUTED).pack(side="right", padx=(10, 4))
offlineq_threshold_entry.bind("<FocusOut>", _save_offlineq_threshold)
offlineq_threshold_entry.bind("<Return>", _save_offlineq_threshold)

offlineq_stats_row = make_card_row(offlineq_outer)
offlineq_stats_row.pack_configure(padx=8, pady=(0, 6))


def _offlineq_stat(parent, title):
    card = ctk.CTkFrame(parent, corner_radius=12, fg_color=THEME_HEADING_BG)
    add_card_to_row(parent, card, padx=4, pady=0)
    ctk.CTkLabel(card, text=title, font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(pady=(6, 0))
    val_lbl = ctk.CTkLabel(card, text="0", font=get_font("value_lg", "bold"))
    val_lbl.pack(pady=(0, 6))
    return val_lbl


offlineq_size_lbl = _offlineq_stat(offlineq_stats_row, "Queue Size")
offlineq_oldest_lbl = _offlineq_stat(offlineq_stats_row, "Oldest Queue Item")
offlineq_next_retry_lbl = _offlineq_stat(offlineq_stats_row, "Estimated Retry Time")
offlineq_interval_lbl = _offlineq_stat(offlineq_stats_row, "Current Retry Interval")

offlineq_cols = ("pid", "destination", "queued_at", "attempts", "next_retry", "last_error")
offlineq_headings = {
    "pid": "Patient ID", "destination": "Destination", "queued_at": "Queued At",
    "attempts": "Attempts", "next_retry": "Next Retry", "last_error": "Last Error",
}
offlineq_widths = {
    "pid": 110, "destination": 130, "queued_at": 150, "attempts": 70,
    "next_retry": 150, "last_error": 260,
}
offlineq_tree_frame = ctk.CTkFrame(offlineq_outer, fg_color=THEME_SURFACE)
offlineq_tree_frame.pack(fill="x", padx=8, pady=(0, 8))
offlineq_tree = ttk.Treeview(offlineq_tree_frame, columns=offlineq_cols, show="headings", height=5)
for col in offlineq_cols:
    offlineq_tree.heading(col, text=offlineq_headings[col])
    offlineq_tree.column(col, width=offlineq_widths[col], anchor="w")
offlineq_tree.pack(fill="x")
ctk.CTkLabel(offlineq_outer, text="(Double-click a row to retry that item immediately)",
             font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=10, pady=(0, 6))


ADMIN_ONLY_TAB_NAMES = [
    "Admin Dashboard", "PACS Health", "Bandwidth", "Export", "Reports",
    "Performance", "Backup", "LDAP / AD", "Query/Retrieve", "Destinations",
    "Routing Rules", "SOP Classes", "Logs", "Settings",
]

# =========================================================
# RECEIVER TAB
# =========================================================

_splash_step(0.50, "Loading receiver worklist…")

rec_monitor_row = make_card_row(tab_receiver_scroll, pady=(10, 0))
rec_monitor_row.pack_configure(padx=10)

rec_monitor_status_card = make_card(rec_monitor_row)
add_card_to_row(rec_monitor_row, rec_monitor_status_card)
ctk.CTkLabel(rec_monitor_status_card.body, text="Operational Status",
             font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 4))
rec_monitor_status_badge = make_status_badge(rec_monitor_status_card.body, "Stopped", kind="stopped")
rec_monitor_status_badge.pack(anchor="w")


def _rec_monitor_stat(title, value="—"):
    card = make_card(rec_monitor_row)
    add_card_to_row(rec_monitor_row, card)
    ctk.CTkLabel(card.body, text=title, font=get_font("micro"),
                 text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 0))
    val_lbl = ctk.CTkLabel(card.body, text=value, font=get_font("title", "bold"),
                            text_color=THEME_TEXT)
    val_lbl.pack(anchor="w", pady=(2, 0))
    return val_lbl


rec_monitor_assoc_lbl = _rec_monitor_stat("Active Associations", "0")
rec_monitor_studies_today_lbl = _rec_monitor_stat("Studies Received Today", "0")
rec_monitor_images_today_lbl = _rec_monitor_stat("Images Received Today", "0")
rec_monitor_queue_lbl = _rec_monitor_stat("Current Queue Size", "0")


_receiver_stop_was_user_initiated = {"value": False}
_receiver_prev_running = {"value": None}


def refresh_receiver_monitoring():
    """Keeps the Receiver tab's own monitoring strip in sync. Reuses the
    exact same state the Dashboard already reads (receiver_state,
    daily_stats, patient_data) -- no new tracking added."""
    running = receiver_state.get("running")

    # 1.3 -- detect a running->stopped transition nobody asked for (e.g. a
    # port conflict or crash mid-listen) and make it loud instead of just
    # quietly flipping the status badge to "Stopped".
    prev_running = _receiver_prev_running["value"]
    if prev_running is True and running is False and not _receiver_stop_was_user_initiated["value"]:
        notify_event("receiver_stopped", "Receiver Stopped Unexpectedly",
                     "The DICOM receiver stopped without being told to. Check the port/config and restart it.")
        show_toast_threadsafe("Receiver Stopped Unexpectedly",
                               "It stopped on its own -- check for a port conflict or crash and restart it.")
        if _rec_config_collapsed["value"]:
            _toggle_rec_config_panel()
        _rec_config_auto_collapse_done["value"] = False  # allow it to auto-collapse again once restarted
        if APP_SETTINGS.get("auto_restart_receiver", False):
            try:
                do_start_receiver()
                show_toast_threadsafe("Receiver Auto-Restarted",
                                       "Settings > Startup & Behavior > Auto-restart receiver brought it back up.")
            except Exception:
                log_exception("Auto-restart of receiver failed")
    if running:
        _receiver_stop_was_user_initiated["value"] = False
    _receiver_prev_running["value"] = running

    rec_monitor_status_badge.update_status("Running" if running else "Stopped",
                                            "running" if running else "stopped")
    if running and not _rec_config_collapsed["value"] and not _rec_config_auto_collapse_done["value"]:
        # Config panel isn't needed once the receiver is actually up and
        # accepting studies -- free the vertical space for the patient list.
        # One-shot: won't re-collapse if the user manually reopens it.
        _rec_config_auto_collapse_done["value"] = True
        _toggle_rec_config_panel()
    active_assoc = 0
    try:
        server_ae = receiver_state.get("server_ae")
        if server_ae is not None:
            active_assoc = len(getattr(server_ae, "active_associations", []) or [])
    except Exception:
        active_assoc = 0
    rec_monitor_assoc_lbl.configure(text=str(active_assoc))
    with _dashboard_stats_lock:
        studies_today = len(daily_stats["studies_received_uids"])
        images_today = daily_stats["images_received"]
    rec_monitor_studies_today_lbl.configure(text=str(studies_today))
    rec_monitor_images_today_lbl.configure(text=str(images_today))
    with data_lock:
        queue_size = sum(1 for d in patient_data.values()
                          if d.get("status") in (STATUS_PENDING, STATUS_RECEIVED, STATUS_IMPORTED))
    rec_monitor_queue_lbl.configure(text=str(queue_size))


rec_config_header = ctk.CTkFrame(tab_receiver_scroll, fg_color=THEME_HEADING_BG, corner_radius=10)
rec_config_header.pack(fill="x", padx=10, pady=(10, 0))

_rec_config_collapsed = {"value": False}
_rec_config_auto_collapse_done = {"value": False}


def _toggle_rec_config_panel():
    _rec_config_collapsed["value"] = not _rec_config_collapsed["value"]
    if _rec_config_collapsed["value"]:
        rec_top.pack_forget()
        rec_config_toggle_btn.configure(text="▸  Receiver Configuration  (click to expand)")
    else:
        rec_top.pack(fill="x", padx=10, pady=(10, 22), after=rec_config_header)
        rec_config_toggle_btn.configure(text="▾  Receiver Configuration  (click to collapse)")


rec_config_toggle_btn = ctk.CTkButton(
    rec_config_header, text="▾  Receiver Configuration  (click to collapse)",
    font=get_font("body", "bold"), anchor="w", corner_radius=8,
    fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER, text_color="#ffffff",
    command=lambda: _toggle_rec_config_panel())
rec_config_toggle_btn.pack(fill="x", padx=8, pady=6)

rec_top = ctk.CTkFrame(tab_receiver_scroll, fg_color="transparent")
rec_top.pack(fill="x", padx=10, pady=(10, 22))

cfg_frame = ctk.CTkFrame(rec_top)
cfg_frame.pack(side="left", padx=(0, 15))

ae_label = ctk.CTkLabel(cfg_frame, text="AE Title")
ae_label.grid(row=0, column=0, padx=8, pady=8, sticky="w")
ae_entry = ctk.CTkEntry(cfg_frame, width=180)
ae_entry.grid(row=0, column=1, padx=8, pady=8)

port_label = ctk.CTkLabel(cfg_frame, text="Port")
port_label.grid(row=1, column=0, padx=8, pady=8, sticky="w")
port_entry = ctk.CTkEntry(cfg_frame, width=180)
port_entry.grid(row=1, column=1, padx=8, pady=8)

# 9.3 -- live inline validation as the user types, instead of only on
# Save (do_save_receiver already runs the same validators; this just
# surfaces the same result earlier, with a colored border).
_rec_cfg_field_default_border = {"ae": ae_entry.cget("border_color"), "port": port_entry.cget("border_color")}


def _live_validate_receiver_field(key):
    entry = ae_entry if key == "ae" else port_entry
    value = entry.get().strip()
    if not value:
        entry.configure(border_color=_rec_cfg_field_default_border[key], border_width=1)
        return
    if key == "ae":
        ok, _msg = validate_ae_title(value, "Receiver AE Title")
    else:
        ok, _msg = validate_port(value, "Port")
    entry.configure(
        border_color=THEME_DANGER if not ok else _rec_cfg_field_default_border[key],
        border_width=2 if not ok else 1)


ae_entry.bind("<KeyRelease>", lambda _evt: _live_validate_receiver_field("ae"))
port_entry.bind("<KeyRelease>", lambda _evt: _live_validate_receiver_field("port"))

autoroute_var = ctk.BooleanVar(value=True)
autoroute_checkbox = ctk.CTkCheckBox(
    cfg_frame, text="Enable AutoRoute (push on arrival)",
    variable=autoroute_var,
    command=lambda: receiver_state.__setitem__("autoroute", autoroute_var.get()))
autoroute_checkbox.grid(row=2, column=0, columnspan=2, padx=8, pady=8, sticky="w")

rec_status_label = ctk.CTkLabel(cfg_frame, text="● Stopped", text_color=THEME_DANGER,
                                 font=get_font("body", "bold"))
rec_status_label.grid(row=3, column=0, columnspan=2, padx=8, pady=4, sticky="w")

# C.5: always-visible (even at sites that never touch doc transfer --
# it just reads "○ Doc transfer: off"), updated via the same
# ui_event_queue pattern as rec_status_label rather than polling.
rec_doc_transfer_status_label = ctk.CTkLabel(cfg_frame, text="○ Doc transfer: off", text_color=THEME_TEXT_MUTED,
                                             font=get_font("small"))
rec_doc_transfer_status_label.grid(row=4, column=0, columnspan=2, padx=8, pady=(0, 4), sticky="w")

# Widgets in this row that are Admin-only per spec: AE Title / Port fields,
# Start/Stop Receiver Server button, Autoroute checkbox. Hidden (not merely
# disabled) in User mode via apply_receiver_mode_visibility() below.
_receiver_admin_only_grid_widgets = [ae_label, ae_entry, port_label, port_entry, autoroute_checkbox]

btn_col = ctk.CTkFrame(rec_top, fg_color="transparent")
btn_col.pack(side="left", padx=15)

receiver_save_btn = ctk.CTkButton(btn_col, text="Save Receiver Config", width=200)
receiver_save_btn.pack(pady=4)

receiver_edit_btn = ctk.CTkButton(btn_col, text="Edit Receiver Config", width=200,
                                   image=get_icon("settings-2", size=15, color=THEME_TEXT), compound="left",
                                   fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
receiver_edit_btn.pack(pady=4)

start_btn = ctk.CTkButton(btn_col, text="Start Receiver", width=200,
                           image=get_icon("play", size=15, color="#ffffff"), compound="left",
                           fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
start_btn.pack(pady=4)

stop_btn = ctk.CTkButton(btn_col, text="Stop Receiver", width=200,
                          image=get_icon("square", size=15, color="#ffffff"), compound="left",
                          fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER)
stop_btn.pack(pady=4)

# Buttons that configure/control the receiver server itself -- Admin-only.
_receiver_admin_only_pack_widgets = [receiver_save_btn, receiver_edit_btn, start_btn, stop_btn]

tray_btn = ctk.CTkButton(btn_col, text="Minimize to Tray", width=200,
                          image=get_icon("folder-open", size=15, color=THEME_TEXT), compound="left",
                          fg_color=THEME_HEADING_BG, hover_color=THEME_NEUTRAL_BTN_HOVER)
tray_btn.pack(pady=4)

rec_open_viewer_btn = ctk.CTkButton(
    btn_col, text="Open in Viewer", width=200,
    fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
    command=lambda: _open_selected_in_viewer(rec_tree))
rec_open_viewer_btn.pack(pady=4)

import_col = ctk.CTkFrame(rec_top, fg_color="transparent")
import_col.pack(side="left", padx=15)

ctk.CTkLabel(import_col, text="Import DICOM Folder:",
             font=get_font("body", "bold")).pack(anchor="w")
import_btn = ctk.CTkButton(import_col, text="Import Folder...", width=200)
import_btn.pack(pady=6)
import_progress = ctk.CTkProgressBar(import_col, width=200)
import_progress.set(0)
import_progress.pack(pady=4)
import_progress_label = ctk.CTkLabel(import_col, text="No import in progress",
                                      font=get_font("small"))
import_progress_label.pack()

# Live "receiving" progress card, always visible below the Receiver
# Configuration box (not inside rec_top) so collapsing "▾ Receiver
# Configuration" (_toggle_rec_config_panel, which only ever
# pack_forget()'s rec_top) can never hide it mid-receive. DICOM C-STORE
# doesn't announce an upfront total file count for an association, so
# this runs as an indeterminate/animated bar while a receive is active
# and shows a live running count + data size; it snaps to a brief
# "Done" state when the sending association releases.
rec_progress_card = ctk.CTkFrame(tab_receiver_scroll, fg_color=THEME_SURFACE, corner_radius=10)
rec_progress_card.pack(fill="x", padx=10, pady=(0, 16), after=rec_top)

rec_progress_inner = ctk.CTkFrame(rec_progress_card, fg_color="transparent")
rec_progress_inner.pack(fill="x", padx=16, pady=12)

ctk.CTkLabel(rec_progress_inner, text="Live Receive Progress",
             font=get_font("body", "bold")).pack(anchor="w")
rec_live_progress = ctk.CTkProgressBar(rec_progress_inner, height=22, corner_radius=8,
                                        progress_color=THEME_ACCENT, mode="determinate")
rec_live_progress.set(0)
rec_live_progress.pack(pady=(6, 4), fill="x")
rec_live_progress_status_label = ctk.CTkLabel(rec_progress_inner, text="",
                                               font=get_font("small"), text_color=THEME_TEXT_MUTED)
rec_live_progress_status_label.pack(anchor="w")
rec_live_progress_label = ctk.CTkLabel(rec_progress_inner, text="Idle — waiting for incoming studies",
                                        font=get_font("body", "bold"))
rec_live_progress_label.pack(anchor="w")
rec_live_throughput_label = ctk.CTkLabel(rec_progress_inner, text="",
                                          font=get_font("small"), text_color=THEME_TEXT_MUTED)
rec_live_throughput_label.pack(anchor="w")

rec_search_frame = ctk.CTkFrame(tab_receiver_scroll, fg_color="transparent")
rec_search_frame.pack(fill="x", padx=10, pady=(14, 5))
ctk.CTkLabel(rec_search_frame, text="", image=get_icon("search", size=15, color=THEME_TEXT_MUTED)).pack(side="left", padx=(0, 5))
rec_search_var = ctk.StringVar()
rec_search_entry = ctk.CTkEntry(rec_search_frame, width=350, textvariable=rec_search_var,
                                 placeholder_text="Search Patient ID/Name, Study UID, Institution, Modality, Status, Destination…")
rec_search_entry.pack(side="left")

WL_STATUS_FILTER_OPTIONS = ["All", STATUS_RECEIVED, STATUS_IMPORTED, STATUS_PENDING,
                            STATUS_SENDING, STATUS_SENT, STATUS_FAILED, STATUS_RETRYING,
                            STATUS_QUEUED, "Stale"]
WL_DATE_FILTER_OPTIONS = ["All Time", "Today", "Last 7 Days", "Last 30 Days"]

ctk.CTkLabel(rec_search_frame, text="Status").pack(side="left", padx=(12, 5))
rec_status_filter_var = ctk.StringVar(value="All")
rec_status_filter_menu = ctk.CTkOptionMenu(rec_search_frame, values=WL_STATUS_FILTER_OPTIONS,
                                            variable=rec_status_filter_var, width=130)
rec_status_filter_menu.pack(side="left")

ctk.CTkLabel(rec_search_frame, text="Date").pack(side="left", padx=(12, 5))
rec_date_filter_var = ctk.StringVar(value="All Time")
rec_date_filter_menu = ctk.CTkOptionMenu(rec_search_frame, values=WL_DATE_FILTER_OPTIONS,
                                          variable=rec_date_filter_var, width=130)
rec_date_filter_menu.pack(side="left")

WL_REPORT_FILTER_OPTIONS = ["All", "Has Report", "No Report"]
ctk.CTkLabel(rec_search_frame, text="Report").pack(side="left", padx=(12, 5))
rec_report_filter_var = ctk.StringVar(value="All")
rec_report_filter_menu = ctk.CTkOptionMenu(rec_search_frame, values=WL_REPORT_FILTER_OPTIONS,
                                            variable=rec_report_filter_var, width=120)
rec_report_filter_menu.pack(side="left")

rec_refresh_btn = ctk.CTkButton(rec_search_frame, text="Refresh", width=110, image=get_icon("refresh-cw", size=14, color=THEME_TEXT), compound="left")
rec_refresh_btn.pack(side="left", padx=10)

# Selection/action buttons live in their own row below the filters, laid
# out with equal-share columns -- so on a narrow window every button
# shrinks together instead of the later ones being clipped out of view.
rec_btn_row, (rec_select_failed_btn, rec_select_pending_btn, rec_select_all_btn,
              rec_select_all_matching_btn, rec_invert_selection_btn,
              rec_clear_selection_btn, rec_export_btn) = make_equal_button_row(
    tab_receiver_scroll,
    [
        dict(text="Select all Failed", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
        dict(text="Select all Pending", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
        dict(text="Select All", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: rec_tree.selection_set(rec_tree.get_children(""))),
        dict(text="Select All Matching Filter", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: _select_all_matching_filter(rec_tree)),
        dict(text="Invert Selection", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: _invert_tree_selection(rec_tree)),
        dict(text="Clear Selection", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: _clear_persistent_selection(rec_tree)),
        dict(text="Export View to CSV", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
    ],
    pady=(0, 5), fill="x",
)
rec_btn_row.pack_configure(padx=10)

rec_badge_label = ctk.CTkLabel(tab_receiver_scroll, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
rec_badge_label.pack(anchor="w", padx=12, pady=(0, 4))

rec_wl_frame, rec_tree = build_worklist_tree(tab_receiver_scroll, tree_name="rec_tree")
rec_wl_frame.pack(fill="both", expand=True, padx=10, pady=(6, 4))
# One big scrollable box holding every matching row (no pagination) -- the
# Treeview keeps its own vertical scrollbar (see build_worklist_tree), so a
# tall fixed height here just sets how much is visible before you scroll,
# it does not limit how many rows actually get inserted.
rec_tree.configure(height=50)

rec_pagination_bar = ctk.CTkFrame(tab_receiver_scroll, fg_color="transparent")
rec_pagination_bar.pack(fill="x", padx=10, pady=(0, 12))
rec_pagination_label = ctk.CTkLabel(rec_pagination_bar, text="", font=get_font("small"),
                                     text_color=THEME_TEXT_MUTED)
rec_pagination_label.pack(side="left")
rec_tree_ref["tree"] = rec_tree


def _update_rec_results_label():
    """Simple 'N results' readout for the un-paginated Receiver worklist
    (replaces the old Next/Prev page controls -- every matching row is
    already in the box, just scroll)."""
    total = len(rec_tree.get_children(""))
    rec_pagination_label.configure(text=f"{total} result{'s' if total != 1 else ''}")


_rec_results_label_hooks.append(_update_rec_results_label)



# =========================================================
# PUSHER TAB
# =========================================================

_splash_step(0.65, "Loading pusher worklist…")

push_config_header = ctk.CTkFrame(tab_pusher_scroll, fg_color=THEME_HEADING_BG, corner_radius=10)
push_config_header.pack(fill="x", padx=10, pady=(10, 0))

_push_config_collapsed = {"value": False}


def _toggle_push_config_panel():
    _push_config_collapsed["value"] = not _push_config_collapsed["value"]
    if _push_config_collapsed["value"]:
        push_top.pack_forget()
        push_config_toggle_btn.configure(text="▸  Pusher Configuration  (click to expand)")
    else:
        push_top.pack(fill="x", padx=10, pady=(10, 22), after=push_config_header)
        push_config_toggle_btn.configure(text="▾  Pusher Configuration  (click to collapse)")


push_config_toggle_btn = ctk.CTkButton(
    push_config_header, text="▾  Pusher Configuration  (click to collapse)",
    font=get_font("body", "bold"), anchor="w", corner_radius=8,
    fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER, text_color="#ffffff",
    command=lambda: _toggle_push_config_panel())
push_config_toggle_btn.pack(fill="x", padx=8, pady=6)

push_top = ctk.CTkFrame(tab_pusher_scroll, fg_color="transparent")
push_top.pack(fill="x", padx=10, pady=(10, 22))

push_dest_frame = ctk.CTkFrame(push_top)
push_dest_frame.pack(side="left", padx=(0, 15))

ctk.CTkLabel(push_dest_frame, text="Active Destination:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
push_dest_var = ctk.StringVar(value="(none configured)")
push_dest_menu = ctk.CTkOptionMenu(push_dest_frame, variable=push_dest_var, values=["(none configured)"], width=200)
push_dest_menu.grid(row=0, column=1, padx=8, pady=6)

anon_var = ctk.BooleanVar(value=False)
ctk.CTkCheckBox(push_dest_frame, text="Anonymize before push",
                variable=anon_var).grid(row=1, column=0, columnspan=2, padx=8, pady=4, sticky="w")

# 2.1 -- Multi-destination push: an alternate mode next to the existing
# single-destination dropdown, gated behind this toggle. When on, Push
# Selected / Push ALL Pending fan out sequentially across every checked
# destination instead of using push_dest_var.
push_multi_dest_toggle_var = ctk.BooleanVar(value=False)
push_multi_dest_checkboxes = {}  # dest name -> (CTkCheckBox, BooleanVar)


def _toggle_multi_dest_mode():
    if push_multi_dest_toggle_var.get():
        push_dest_menu.configure(state="disabled")
        push_multi_dest_list_frame.grid(row=3, column=0, columnspan=2, padx=8, pady=(0, 6), sticky="ew")
    else:
        push_dest_menu.configure(state="normal")
        push_multi_dest_list_frame.grid_forget()


ctk.CTkCheckBox(push_dest_frame, text="Push to multiple destinations",
                variable=push_multi_dest_toggle_var,
                command=_toggle_multi_dest_mode).grid(row=2, column=0, columnspan=2, padx=8, pady=(0, 2), sticky="w")

push_multi_dest_list_frame = ctk.CTkFrame(push_dest_frame, fg_color=THEME_HEADING_BG, corner_radius=8)
# Not gridded yet -- _toggle_multi_dest_mode() grids it in row 3 once the
# toggle above is checked.


def _refresh_multi_dest_checkboxes():
    """Rebuilds the multi-destination checkbox list to match
    load_destinations(), preserving which ones were already checked.
    Called from refresh_destinations_ui() so it always matches the
    single-select dropdown's options."""
    dests = load_destinations()
    existing = {name: var.get() for name, (_cb, var) in push_multi_dest_checkboxes.items()}
    for child in push_multi_dest_list_frame.winfo_children():
        child.destroy()
    push_multi_dest_checkboxes.clear()
    if not dests:
        ctk.CTkLabel(push_multi_dest_list_frame, text="No destinations configured",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=8, pady=6)
        return
    for d in dests:
        name = d.get("name", "")
        var = ctk.BooleanVar(value=existing.get(name, False))
        cb = ctk.CTkCheckBox(push_multi_dest_list_frame, text=name, variable=var)
        cb.pack(anchor="w", padx=8, pady=2)
        push_multi_dest_checkboxes[name] = (cb, var)


def _get_selected_multi_destinations():
    dests_by_name = {d.get("name", ""): d for d in load_destinations()}
    return [dests_by_name[name] for name, (_cb, var) in push_multi_dest_checkboxes.items()
            if var.get() and name in dests_by_name]

push_btn_col = ctk.CTkFrame(push_top, fg_color="transparent")
push_btn_col.pack(side="left", padx=15)

push_all_btn = ctk.CTkButton(push_btn_col, text="Push Selected", width=200,
                              fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
push_all_btn.pack(pady=4)
push_everything_btn = ctk.CTkButton(push_btn_col, text="Push ALL Pending", width=200)
push_everything_btn.pack(pady=4)
push_stop_btn = ctk.CTkButton(push_btn_col, text="■  Stop Push", width=200,
                               fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER)
push_stop_btn.pack(pady=4)
push_open_viewer_btn = ctk.CTkButton(
    push_btn_col, text="Open in Viewer", width=200,
    fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
    command=lambda: _open_selected_in_viewer(push_tree))
push_open_viewer_btn.pack(pady=4)
echo_btn = ctk.CTkButton(push_btn_col, text="C-ECHO Active Dest.", width=200)
echo_btn.pack(pady=4)

# Overall Push Progress lives in its own always-visible card BELOW the
# Pusher Configuration box (not inside push_top) specifically so that
# collapsing "▾ Pusher Configuration" (_toggle_push_config_panel, which
# only ever pack_forget()'s push_top) can never hide it mid-push.
push_progress_card = ctk.CTkFrame(tab_pusher_scroll, fg_color=THEME_SURFACE, corner_radius=10)
push_progress_card.pack(fill="x", padx=10, pady=(0, 16), after=push_top)

push_progress_inner = ctk.CTkFrame(push_progress_card, fg_color="transparent")
push_progress_inner.pack(fill="x", padx=16, pady=12)

ctk.CTkLabel(push_progress_inner, text="Overall Push Progress",
             font=get_font("body", "bold")).pack(anchor="w")
overall_progress = ctk.CTkProgressBar(push_progress_inner, height=22, corner_radius=8,
                                       progress_color=THEME_ACCENT)
overall_progress.set(0)
overall_progress.pack(pady=(6, 4), fill="x")
push_progress_status_label = ctk.CTkLabel(push_progress_inner, text="",
                                           font=get_font("small"), text_color=THEME_TEXT_MUTED)
push_progress_status_label.pack(anchor="w")
overall_progress_label = ctk.CTkLabel(push_progress_inner, text="0 / 0 images sent",
                                       font=get_font("body", "bold"))
overall_progress_label.pack(anchor="w")
throughput_label = ctk.CTkLabel(push_progress_inner, text="",
                                 font=get_font("small"), text_color=THEME_TEXT_MUTED)
throughput_label.pack(anchor="w")

push_wl_heading_row = ctk.CTkFrame(tab_pusher_scroll, fg_color="transparent")
push_wl_heading_row.pack(fill="x", padx=10, pady=(4, 2))
ctk.CTkLabel(push_wl_heading_row, text="Push Worklist",
             font=get_font("section_lg", "bold")).pack(side="left")
ctk.CTkLabel(push_wl_heading_row,
             text="Search, filter, and multi-select studies below, then use Push Selected / Push ALL Pending above.",
             font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(side="left", padx=(10, 0))

push_search_frame = ctk.CTkFrame(tab_pusher_scroll, fg_color="transparent")
push_search_frame.pack(fill="x", padx=10, pady=(0, 5))
ctk.CTkLabel(push_search_frame, text="", image=get_icon("search", size=15, color=THEME_TEXT_MUTED)).pack(side="left", padx=(0, 5))
push_search_var = ctk.StringVar()
push_search_entry = ctk.CTkEntry(push_search_frame, width=350, textvariable=push_search_var,
                                  placeholder_text="Search Patient ID/Name, Study UID, Institution, Modality, Status, Destination…")
push_search_entry.pack(side="left")

ctk.CTkLabel(push_search_frame, text="Status").pack(side="left", padx=(12, 5))
push_status_filter_var = ctk.StringVar(value="All")
push_status_filter_menu = ctk.CTkOptionMenu(push_search_frame, values=WL_STATUS_FILTER_OPTIONS,
                                             variable=push_status_filter_var, width=130)
push_status_filter_menu.pack(side="left")

ctk.CTkLabel(push_search_frame, text="Date").pack(side="left", padx=(12, 5))
push_date_filter_var = ctk.StringVar(value="All Time")
push_date_filter_menu = ctk.CTkOptionMenu(push_search_frame, values=WL_DATE_FILTER_OPTIONS,
                                           variable=push_date_filter_var, width=130)
push_date_filter_menu.pack(side="left")

ctk.CTkLabel(push_search_frame, text="Report").pack(side="left", padx=(12, 5))
push_report_filter_var = ctk.StringVar(value="All")
push_report_filter_menu = ctk.CTkOptionMenu(push_search_frame, values=WL_REPORT_FILTER_OPTIONS,
                                             variable=push_report_filter_var, width=120)
push_report_filter_menu.pack(side="left")

ctk.CTkLabel(push_search_frame, text="Destination").pack(side="left", padx=(12, 5))
push_dest_filter_var = ctk.StringVar(value="All")
push_dest_filter_menu = ctk.CTkOptionMenu(push_search_frame, values=["All"],
                                           variable=push_dest_filter_var, width=160)
push_dest_filter_menu.pack(side="left")

push_refresh_btn = ctk.CTkButton(push_search_frame, text="Refresh", width=110, image=get_icon("refresh-cw", size=14, color=THEME_TEXT), compound="left")
push_refresh_btn.pack(side="left", padx=10)
push_view_options_btn = ctk.CTkButton(push_search_frame, text="View", width=90,
                                       image=get_icon("settings-2", size=14, color=THEME_TEXT), compound="left",
                                       fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
push_view_options_btn.pack(side="right", padx=4)

# Selection/action buttons in their own equal-share row -- see the
# matching comment on the Receiver tab's rec_btn_row for why.
push_btn_row, (push_select_failed_btn, push_select_pending_btn, push_select_all_btn,
               push_select_all_matching_btn, push_invert_selection_btn,
               push_clear_selection_btn, push_export_btn) = make_equal_button_row(
    tab_pusher_scroll,
    [
        dict(text="Select all Failed", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
        dict(text="Select all Pending", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
        dict(text="Select All", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: push_tree.selection_set(push_tree.get_children(""))),
        dict(text="Select All Matching Filter", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: _select_all_matching_filter(push_tree)),
        dict(text="Invert Selection", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: _invert_tree_selection(push_tree)),
        dict(text="Clear Selection", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
             command=lambda: _clear_persistent_selection(push_tree)),
        dict(text="Export View to CSV", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
    ],
    pady=(0, 5), fill="x",
)
push_btn_row.pack_configure(padx=10)
ctk.CTkLabel(tab_pusher_scroll, text="(Right-click a row for more actions)",
             font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=12, pady=(0, 2))

push_badge_label = ctk.CTkLabel(tab_pusher_scroll, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
push_badge_label.pack(anchor="w", padx=12, pady=(0, 4))

push_wl_frame, push_tree = build_worklist_tree(tab_pusher_scroll, tree_name="push_tree")
push_wl_frame.pack(fill="both", expand=True, padx=10, pady=(6, 4))
# Same reasoning as the Receiver tab: bounded to one page's worth of rows
# (see pagination below) rather than expand=True, which doesn't size
# predictably inside a CTkScrollableFrame -- the whole tab (this box
# included) scrolls as a single page instead.
push_tree.configure(height=50)
_PAGINATED_TREES.add(id(push_tree))
_get_persistent_selection(push_tree)  # register for cross-page selection tracking (1.1)
push_tree.bind("<<TreeviewSelect>>", lambda e: _sync_persistent_selection(push_tree, e), add="+")
push_tree_ref["tree"] = push_tree

push_pagination_bar = ctk.CTkFrame(tab_pusher_scroll, fg_color="transparent")
push_pagination_bar.pack(fill="x", padx=10, pady=(0, 12))
push_pagination_label = ctk.CTkLabel(push_pagination_bar, text="", font=get_font("small"),
                                      text_color=THEME_TEXT_MUTED)
push_pagination_label.pack(side="left")
push_pagination_next_btn = ctk.CTkButton(push_pagination_bar, text="Next ›", width=90,
                                          fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                                          command=lambda: _change_tree_page(push_tree, 1))
push_pagination_next_btn.pack(side="right", padx=(6, 0))
push_pagination_prev_btn = ctk.CTkButton(push_pagination_bar, text="‹ Prev", width=90,
                                          fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                                          command=lambda: _change_tree_page(push_tree, -1))
push_pagination_prev_btn.pack(side="right")
_pagination_ui[id(push_tree)] = (push_pagination_label, push_pagination_prev_btn, push_pagination_next_btn)

_splash_step(0.78, "Finishing interface…")

# =========================================================
# ADMIN-ONLY TABS (Dashboard, Query/Retrieve, Destinations, Routing Rules,
# SOP Classes, Logs). These do not exist in User mode. To satisfy "wrap the
# tabview.add() calls in a check against the current mode, and rebuild
# rather than duplicate tab-construction code", all of their construction +
# command-wiring lives in this ONE function, called once at startup (if
# launching as Admin) and again any time the operator unlocks Admin after
# having been in User mode. tear_down_admin_only_tabs() (below) removes them
# again on the way back to User mode.
# =========================================================

# =========================================================
# SETTINGS TAB -- action handlers (Priority 10)
# =========================================================
# Kept together, above build_admin_only_tabs(), since every one of these
# is wired to a Settings-tab control built inside that function. Follows
# the same "background thread for I/O, app.after(0, ...) back to the
# main thread for UI updates" pattern used by do_manual_backup() etc.

_settings_diag_process = psutil.Process(os.getpid()) if PSUTIL_AVAILABLE else None


def refresh_settings_diagnostics():
    """Updates the live Performance stats + Diagnostics readouts on the
    Settings tab. Cheap -- only runs while the Settings tab is actually
    selected (see the tabview.get() == "Settings" gate in periodic_refresh)."""
    if "settings_perf_stat_lbls" not in globals():
        return
    try:
        interval_ms = get_refresh_interval_ms()
        hz = 1000.0 / interval_ms if interval_ms > 0 else 0.0
        settings_perf_stat_lbls["refresh_freq"].configure(text=f"{hz:.1f} Hz  ({interval_ms} ms)")
        settings_perf_stat_lbls["avg_update_ms"].configure(text=f"{_ui_perf_stats['avg_update_ms']:.1f} ms")
        cpu_txt = mem_txt = "N/A (psutil not installed)"
        if PSUTIL_AVAILABLE:
            try:
                cpu_txt = f"{psutil.cpu_percent(interval=None):.0f}%"
                mem_txt = f"{psutil.virtual_memory().percent:.0f}%"
            except Exception:
                pass
        settings_perf_stat_lbls["cpu"].configure(text=cpu_txt)
        settings_perf_stat_lbls["memory"].configure(text=mem_txt)
    except Exception:
        log_exception("Failed to refresh Settings performance stats")

    try:
        settings_diag_lbls["version"].configure(text=APP_VERSION)
        settings_diag_lbls["python"].configure(text=platform.python_version())
        settings_diag_lbls["os"].configure(text=f"{platform.system()} {platform.release()}")
        settings_diag_lbls["uptime"].configure(text=_fmt_uptime(time.time() - APP_START_TIME))
        if PSUTIL_AVAILABLE and _settings_diag_process is not None:
            try:
                rss_mb = _settings_diag_process.memory_info().rss / (1024 * 1024)
                settings_diag_lbls["memory_usage"].configure(text=f"{rss_mb:.1f} MB")
            except Exception:
                settings_diag_lbls["memory_usage"].configure(text="N/A")
        else:
            settings_diag_lbls["memory_usage"].configure(text="N/A (psutil not installed)")
        settings_diag_lbls["workers"].configure(text=str(threading.active_count()))
        queue_size = get_offline_queue_summary()[0]
        settings_diag_lbls["queue_size"].configure(text=str(queue_size))
        settings_diag_lbls["receiver_status"].configure(
            text="Running" if receiver_state.get("running") else "Stopped")
    except Exception:
        log_exception("Failed to refresh Settings diagnostics")

    # D.2 / C.6: passive misconfiguration warnings.
    if "settings_warnings_frame" in globals():
        try:
            warnings = []

            tls_cfg = load_tls_config()
            if tls_cfg.get("enabled") and not tls_cfg.get("ca_cert"):
                warnings.append(
                    "TLS is enabled but no CA certificate is configured: the connection "
                    "is encrypted but the remote server's identity is not verified.")

            global_doc_transfer = APP_SETTINGS.get("doc_transfer_enabled", False)
            dests = load_destinations()
            any_dest_doc_transfer = any(d.get("doc_transfer_enabled") for d in dests)
            if global_doc_transfer and not any_dest_doc_transfer:
                warnings.append(
                    "Document transfer is enabled globally, but no destination has it "
                    "enabled -- nothing will actually be sent.")
            elif any_dest_doc_transfer and not global_doc_transfer:
                warnings.append(
                    "A destination has document transfer enabled, but the global "
                    "receiver setting above is off -- this receiver won't accept documents.")

            _render_warning_rows(settings_warnings_frame, warnings)
        except Exception:
            log_exception("Failed to refresh Settings misconfiguration warnings")


def do_settings_set(key, value, status_text=None):
    """Generic 'persist one setting immediately' handler shared by every
    control on the Settings tab."""
    _update_app_setting(key, value)
    if "settings_status_lbl" in globals():
        settings_status_lbl.configure(
            text=status_text or "Saved.", text_color=THEME_SUCCESS)


def do_settings_accessibility_change(key, value, label):
    """Persists an Accessibility setting (High Contrast Mode / Larger
    Click Targets) and immediately re-applies it to every already-built
    widget via refresh_ui_theme() -- no restart required."""
    _update_app_setting(key, value)
    try:
        refresh_ui_theme()
    except Exception:
        log_exception(f"Failed to live-refresh UI after changing {key}")
    if "settings_status_lbl" in globals():
        settings_status_lbl.configure(
            text=f"{label} {'enabled' if value else 'disabled'} and applied.",
            text_color=THEME_SUCCESS)


def do_settings_reduced_motion_change(value):
    """Persists Reduced Motion Mode and immediately forces every live
    MarqueeLabel (status bar, KPI tiles) to re-evaluate it -- no restart
    and no waiting for the text to next change."""
    _update_app_setting("reduced_motion", value)
    try:
        MarqueeLabel.refresh_all_for_motion_setting()
    except Exception:
        log_exception("Failed to live-refresh marquee labels after changing reduced_motion")
    if "settings_status_lbl" in globals():
        settings_status_lbl.configure(
            text=f"Reduced Motion Mode {'enabled' if value else 'disabled'} and applied.",
            text_color=THEME_SUCCESS)


def do_settings_performance_mode_change(mode_name):
    apply_performance_mode(mode_name)
    settings_refresh_rate_var.set(f"{APP_SETTINGS['refresh_rate_ms']} ms")
    settings_max_workers_var.set(str(APP_SETTINGS["max_worker_threads"]))
    settings_max_pushes_var.set(str(APP_SETTINGS["max_simultaneous_pushes"]))
    settings_status_lbl.configure(
        text=f"Performance Mode set to {mode_name}. Refresh rate, worker threads, "
             f"and simultaneous pushes updated to match (you can still fine-tune each below).",
        text_color=THEME_SUCCESS)


def do_settings_refresh_rate_change(choice_text):
    try:
        ms = int(choice_text.split()[0])
    except Exception:
        ms = REFRESH_INTERVAL_MS
    do_settings_set("refresh_rate_ms", ms, f"UI refresh rate set to {ms} ms. Takes effect on the next tick.")


def do_settings_ui_scaling_change(choice_text):
    try:
        pct = int(choice_text.replace("%", "").strip())
    except Exception:
        pct = 100
    _update_app_setting("ui_scaling_pct", pct)
    apply_ui_scaling()
    settings_status_lbl.configure(text=f"UI scaling set to {pct}%.", text_color=THEME_SUCCESS)


def do_settings_font_scale_change(choice_text):
    try:
        pct = int(choice_text.replace("%", "").strip())
    except Exception:
        pct = 100
    _update_app_setting("font_scale_pct", pct)
    settings_status_lbl.configure(
        text=f"Font size set to {pct}%. Already-open labels update the next time they redraw; "
             f"for a fully consistent pass across every tab, reopen the app.",
        text_color=THEME_SUCCESS)


def do_settings_logging_level_change(choice_text):
    _update_app_setting("logging_level", choice_text)
    apply_logging_level(choice_text)
    settings_status_lbl.configure(text=f"Logging level set to {choice_text}.", text_color=THEME_SUCCESS)


def do_settings_clear_cache():
    """Clears Python's __pycache__ (the only real 'temp cache' this app
    generates) plus stale *.tmp files it may have left behind."""
    if not modern_confirm("Clear Cache", "Remove temporary cache files? This is safe and does not touch your data."):
        return
    removed = 0
    try:
        for root, dirs, files in os.walk(APP_DIR):
            if "__pycache__" in dirs:
                shutil.rmtree(os.path.join(root, "__pycache__"), ignore_errors=True)
                removed += 1
            for fn in files:
                if fn.endswith(".tmp"):
                    try:
                        os.remove(os.path.join(root, fn))
                        removed += 1
                    except Exception:
                        pass
        settings_status_lbl.configure(text=f"Cache cleared ({removed} item(s) removed).", text_color=THEME_SUCCESS)
        write_audit_log("SETTINGS-CACHE-CLEARED", f"items_removed={removed}")
    except Exception as e:
        log_exception("Failed to clear cache")
        settings_status_lbl.configure(text=f"Failed to clear cache: {e}", text_color=THEME_DANGER)


def do_settings_remove_archived_logs():
    paths = _log_archive_paths()
    if not paths:
        settings_status_lbl.configure(text="No archived logs to remove.", text_color=THEME_TEXT_MUTED)
        return
    if not modern_confirm(
        "Remove Archived Logs",
        f"Permanently delete {len(paths)} archived log file(s)? This cannot be undone.",
        danger=True,
    ):
        return
    removed = 0
    for p in paths:
        try:
            os.remove(p)
            removed += 1
        except Exception:
            pass
    settings_status_lbl.configure(text=f"Removed {removed} of {len(paths)} archived log file(s).",
                                  text_color=THEME_SUCCESS)
    write_audit_log("SETTINGS-ARCHIVES-REMOVED", f"count={removed}")
    try:
        refresh_archive_list()
    except Exception:
        pass


def do_settings_open_config_dir():
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(APP_DIR)
        elif system == "Darwin":
            subprocess.Popen(["open", APP_DIR])
        else:
            subprocess.Popen(["xdg-open", APP_DIR])
        settings_status_lbl.configure(text=f"Opened {APP_DIR}", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Failed to open config directory")
        settings_status_lbl.configure(text=f"Could not open the config directory: {e}", text_color=THEME_DANGER)


def do_settings_export():
    default_name = f"rapps_settings_{datetime.date.today().isoformat()}.json"
    out_path = filedialog.asksaveasfilename(
        title="Export App Settings", defaultextension=".json",
        initialfile=default_name, filetypes=[("JSON settings file", "*.json")])
    if not out_path:
        return
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(APP_SETTINGS, f, indent=2)
        write_audit_log("SETTINGS-EXPORTED", f"path={out_path}")
        settings_status_lbl.configure(text=f"Settings exported to {out_path}", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Failed to export app settings")
        settings_status_lbl.configure(text=f"Export failed: {e}", text_color=THEME_DANGER)


def do_settings_import():
    in_path = filedialog.askopenfilename(
        title="Import App Settings", filetypes=[("JSON settings file", "*.json"), ("All files", "*.*")])
    if not in_path:
        return
    try:
        with open(in_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if not isinstance(loaded, dict):
            raise ValueError("File does not contain a settings object.")
    except Exception as e:
        log_exception("Failed to read imported app settings")
        settings_status_lbl.configure(text=f"Could not read that file: {e}", text_color=THEME_DANGER)
        return
    if not modern_confirm("Import Settings", "This will overwrite your current Settings tab preferences. Continue?"):
        return
    merged = dict(DEFAULT_APP_SETTINGS)
    merged.update(loaded)
    APP_SETTINGS.clear()
    APP_SETTINGS.update(merged)
    save_app_settings(APP_SETTINGS)
    apply_ui_scaling()
    apply_logging_level()
    try:
        refresh_ui_theme()
        MarqueeLabel.refresh_all_for_motion_setting()
    except Exception:
        log_exception("Failed to live-refresh UI after importing settings")
    write_audit_log("SETTINGS-IMPORTED", f"path={in_path}")
    settings_status_lbl.configure(
        text="Settings imported and applied. Compact Mode takes effect after "
             "reopening the app; everything else is already applied.", text_color=THEME_SUCCESS)
    _populate_settings_controls_from_app_settings()


def do_settings_restore_defaults():
    if not modern_confirm(
        "Restore Factory Defaults", "Reset every Settings tab preference to its factory default? "
        "This does not affect Destinations, Routing Rules, or other Configuration tabs.",
        danger=True,
    ):
        return
    APP_SETTINGS.clear()
    APP_SETTINGS.update(dict(DEFAULT_APP_SETTINGS))
    save_app_settings(APP_SETTINGS)
    apply_ui_scaling()
    apply_logging_level()
    try:
        refresh_ui_theme()
        MarqueeLabel.refresh_all_for_motion_setting()
    except Exception:
        log_exception("Failed to live-refresh UI after restoring factory defaults")
    write_audit_log("SETTINGS-FACTORY-RESET", "")
    settings_status_lbl.configure(text="Settings restored to factory defaults.", text_color=THEME_SUCCESS)
    _populate_settings_controls_from_app_settings()


def do_settings_manual_backup():
    default_name = f"rapps_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    dest_path = filedialog.asksaveasfilename(
        title="Backup Every Configuration File As…", defaultextension=".zip",
        filetypes=[("ZIP archive", "*.zip")], initialfile=default_name)
    if not dest_path:
        return
    settings_status_lbl.configure(text="Backing up…", text_color=THEME_TEXT_MUTED)

    def run():
        try:
            record = run_backup_job(dest_path, triggered_by="settings-manual")
            def on_done():
                if record["validation_ok"]:
                    settings_status_lbl.configure(text=record["validation_message"], text_color=THEME_SUCCESS)
                else:
                    settings_status_lbl.configure(
                        text=f"Backup completed but validation failed: {record['validation_message']}",
                        text_color=THEME_WARNING)
            app.after(0, on_done)
        except Exception as e:
            log_exception("Settings-tab manual backup failed")
            err_msg = str(e)
            def on_fail():
                settings_status_lbl.configure(text=f"Backup failed: {err_msg}", text_color=THEME_DANGER)
            app.after(0, on_fail)

    threading.Thread(target=run, daemon=True).start()


def do_settings_verify_integrity():
    """Checks every file in the backup manifest for existence and, for
    the JSON ones, that they actually parse -- reporting anything
    missing or corrupted rather than silently failing later."""
    missing, corrupt, ok = [], [], 0
    for category, paths in _raw_backup_manifest().items():
        for p in paths:
            if not os.path.isfile(p):
                missing.append(p)
                continue
            if p.endswith(".json"):
                try:
                    with open(p, "r", encoding="utf-8") as f:
                        json.load(f)
                    ok += 1
                except Exception:
                    corrupt.append(p)
            else:
                ok += 1  # .enc/.log files aren't JSON-parseable here; existence is the check
    lines = [f"{ok} file(s) OK."]
    if missing:
        lines.append(f"{len(missing)} missing: " + ", ".join(missing[:6]) + (" …" if len(missing) > 6 else ""))
    if corrupt:
        lines.append(f"{len(corrupt)} corrupted (invalid JSON): " + ", ".join(corrupt[:6]))
    text = " ".join(lines)
    color = THEME_SUCCESS if not missing and not corrupt else (THEME_WARNING if not corrupt else THEME_DANGER)
    settings_status_lbl.configure(text=text, text_color=color)
    modern_showinfo("Configuration Integrity", text)


_SETTINGS_SHORTCUTS = [
    ("Ctrl+F", "Focus universal search"),
    ("Escape", "Close the open panel/overlay"),
    ("Click a worklist row", "Select a patient"),
    ("Double-click a worklist row", "Open patient details"),
]


def do_settings_show_shortcuts():
    lines = "\n".join(f"{k}   —   {v}" for k, v in _SETTINGS_SHORTCUTS)
    modern_showinfo("Keyboard Shortcuts", lines)


def do_settings_toggle_experimental(enabled):
    if enabled and not modern_confirm(
        "Experimental Features",
        "Experimental features are still being validated and may be unstable or change without notice. "
        "Enable them anyway?",
        danger=True,
    ):
        settings_experimental_var.set(False)
        return
    do_settings_set("experimental_features_enabled", enabled,
                    "Experimental features enabled." if enabled else "Experimental features disabled.")


def do_settings_save_doc_transfer_port(value):
    ok, port_or_err = validate_port(value, "Document Transfer Port")
    if not ok:
        if "settings_status_lbl" in globals():
            settings_status_lbl.configure(text=port_or_err, text_color=THEME_DANGER)
        return
    do_settings_set("doc_transfer_receiver_port", str(port_or_err),
                    f"Document transfer port set to {port_or_err}.")


def do_settings_save_doc_transfer_auth_key(value):
    key = (value or "").strip()
    do_settings_set(
        "doc_transfer_receiver_auth_key", key,
        "Document transfer authentication key set -- senders must now present it."
        if key else "Document transfer authentication key cleared -- transfers are unauthenticated again.")
    _refresh_doc_transfer_auth_warning()


def _refresh_doc_transfer_auth_warning():
    """§3.6 fix: an empty doc_transfer_receiver_auth_key means every
    connection to the doc-transfer listener is accepted unauthenticated
    -- intentional, for backward compatibility (see
    _handle_doc_transfer_connection), but previously had no proactive
    indicator anywhere telling an operator that's the current state.
    Shows/clears a warning banner under the Settings-tab field for it.
    Safe to call before the Settings tab has been built (checks
    globals() first) and after Save, Import, or Restore Defaults."""
    if "doc_transfer_auth_warning_label" not in globals():
        return
    try:
        if str(APP_SETTINGS.get("doc_transfer_receiver_auth_key", "") or ""):
            doc_transfer_auth_warning_label.configure(text="")
        else:
            doc_transfer_auth_warning_label.configure(
                text="⚠ No document transfer auth key is set -- this receiver currently "
                     "accepts document transfers from anyone on the network.")
    except Exception:
        log_exception("Failed to refresh doc-transfer auth warning banner")


def _populate_settings_controls_from_app_settings():
    """Re-syncs every Settings-tab control's displayed value from
    APP_SETTINGS -- used after Import/Restore Defaults so the UI reflects
    what was just loaded without needing to rebuild the tab."""
    if "settings_refresh_rate_var" not in globals():
        return
    try:
        settings_refresh_rate_var.set(f"{APP_SETTINGS['refresh_rate_ms']} ms")
        settings_perf_mode_var.set(APP_SETTINGS["performance_mode"])
        settings_ui_scale_var.set(f"{APP_SETTINGS['ui_scaling_pct']}%")
        settings_font_scale_var.set(f"{APP_SETTINGS['font_scale_pct']}%")
        settings_compact_var.set(APP_SETTINGS["compact_mode"])
        settings_remember_geom_var.set(APP_SETTINGS["remember_window_geometry"])
        settings_launch_max_var.set(APP_SETTINGS["launch_maximized"])
        settings_toast_var.set(APP_SETTINGS["toast_notifications_enabled"])
        settings_sound_var.set(APP_SETTINGS["notification_sounds_enabled"])
        settings_notif_duration_var.set(f"{APP_SETTINGS['notification_duration_sec']} sec")
        settings_critical_only_var.set(APP_SETTINGS["critical_only_notifications"])
        settings_launch_startup_var.set(APP_SETTINGS["launch_on_system_startup"])
        settings_auto_start_recv_var.set(APP_SETTINGS["auto_start_receiver_on_launch"])
        settings_auto_restart_var.set(APP_SETTINGS["auto_restart_receiver"])
        settings_restore_session_var.set(APP_SETTINGS["restore_previous_session"])
        settings_minimize_tray_var.set(APP_SETTINGS["minimize_to_tray_on_close"])
        settings_confirm_close_var.set(APP_SETTINGS["confirm_before_close"])
        settings_check_updates_var.set(APP_SETTINGS["auto_check_for_updates"])
        settings_log_level_var.set(APP_SETTINGS["logging_level"])
        settings_hc_var.set(APP_SETTINGS["high_contrast_mode"])
        settings_click_targets_var.set(APP_SETTINGS["larger_click_targets"])
        settings_reduced_motion_var.set(APP_SETTINGS["reduced_motion"])
        settings_max_workers_var.set(str(APP_SETTINGS["max_worker_threads"]))
        settings_max_pushes_var.set(str(APP_SETTINGS["max_simultaneous_pushes"]))
        settings_net_timeout_var.set(f"{APP_SETTINGS['network_timeout_sec']} sec")
        settings_cache_limit_var.set(f"{APP_SETTINGS['cache_size_limit_mb']} MB")
        settings_experimental_var.set(APP_SETTINGS["experimental_features_enabled"])
        settings_doc_transfer_var.set(APP_SETTINGS.get("doc_transfer_enabled", False))
        settings_doc_transfer_port_var.set(str(APP_SETTINGS.get("doc_transfer_receiver_port", "11244")))
        if "settings_doc_transfer_chunk_var" in globals():
            settings_doc_transfer_chunk_var.set(f"{APP_SETTINGS.get('doc_transfer_chunk_size_kb', 64)} KB")
        if "settings_doc_transfer_timeout_var" in globals():
            settings_doc_transfer_timeout_var.set(
                f"{APP_SETTINGS.get('doc_transfer_timeout_sec', DOC_TRANSFER_CONNECT_TIMEOUT_SEC)} sec")
        if "settings_admin_idle_timeout_var" in globals():
            _idle_min = APP_SETTINGS.get("admin_idle_timeout_min", 15)
            settings_admin_idle_timeout_var.set(f"{_idle_min} min" if _idle_min else "Disabled")
    except Exception:
        log_exception("Failed to re-sync Settings tab controls")


def do_settings_launch_on_startup_change(enabled):
    """Best-effort: Windows via the Run registry key; other OSes don't
    have one universal equivalent, so the preference is persisted but
    flagged as not automatically actionable there."""
    system = platform.system()
    if system == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                  r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            if enabled:
                exe = sys.executable
                script = os.path.abspath(sys.argv[0])
                winreg.SetValueEx(key, "RAppsDICOM", 0, winreg.REG_SZ, f'"{exe}" "{script}"')
            else:
                try:
                    winreg.DeleteValue(key, "RAppsDICOM")
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
            do_settings_set("launch_on_system_startup", enabled,
                            "Launch on system startup " + ("enabled." if enabled else "disabled."))
        except Exception as e:
            log_exception("Failed to set launch-on-startup registry key")
            do_settings_set("launch_on_system_startup", enabled,
                            f"Preference saved, but the startup entry could not be written: {e}")
    else:
        do_settings_set("launch_on_system_startup", enabled,
                        "Preference saved. Automatic startup registration isn't available on this OS -- "
                        "add the app to your desktop environment's startup apps manually.")


def build_admin_only_tabs():
    global tab_dashboard, tab_qr, tab_destinations, tab_routing, tab_sop, tab_logs
    global qr_top, qr_cfg_frame, qr_ae_entry, qr_ip_entry, qr_port_entry
    global qr_filter_frame, qr_filter_entries, qr_btn_col, qr_find_btn, qr_retrieve_btn
    global qr_status_label, qr_results_cols, qr_results_headings, qr_results_widths
    global qr_tree_frame, qr_vsb, qr_hsb, qr_tree
    global dest_outer, dest_list_frame, dest_listbox_var, dest_listbox
    global dest_select_var, dest_names_var, dest_optmenu, dest_form_frame
    global dest_fields, dest_default_var, dest_btn_frame, dest_add_btn, dest_del_btn
    global dest_field_default_border
    global dest_trust_auth_badge
    global dest_doc_transfer_enabled_var, dest_doc_transfer_use_dicom_host_var
    global dest_echo_btn, dest_status_lbl
    global routing_outer, routing_list_frame, routing_cols, routing_headings, routing_widths
    global routing_vsb, routing_tree, routing_form_frame, routing_fields
    global routing_dest_var, routing_dest_menu, routing_add_btn, routing_del_btn
    global routing_up_btn, routing_down_btn, routing_test_btn
    global sop_outer, sop_split, sop_left, sop_classes_box, sop_right, sop_ts_box
    global sop_btn_row, sop_load_btn, sop_save_btn, sop_status_lbl
    global sop_search_var, sop_search_entry
    global logs_top, log_file_var, log_selector, log_refresh_btn
    global log_archive_now_btn, log_export_btn, log_export_filtered_btn
    global logs_row2, log_retention_var, log_retention_menu
    global log_view_mode_var, log_view_mode_menu, log_archive_var, log_archive_menu
    global log_status_lbl
    global log_tail_var, logs_text
    global log_copy_btn, logs_row1b, log_search_var, log_search_entry
    global log_severity_var, log_severity_menu, log_match_count_lbl
    global log_display_mode_var, log_display_mode_menu
    global logs_row3, log_struct_dest_var, log_struct_dest_menu
    global log_struct_range_var, log_struct_range_menu, log_open_related_btn
    global logs_body, log_raw_container, log_structured_container
    global log_struct_tree, log_detail_box, log_struct_records_by_iid
    global log_struct_cols, log_struct_headings, log_struct_widths
    global dash_received_lbl, dash_pushed_lbl, dash_failed_lbl, dash_docs_pending_lbl
    global dash_disk_bar, dash_disk_lbl, dash_stale_lbl, dash_stale_jump_btn
    global dash_receiver_status_badge, dash_failures_tree, dash_dest_health_tree
    global dash_export_btn, dash_refresh_btn
    global health_outer, health_top, health_run_all_btn, health_status_lbl
    global health_cols, health_headings, health_widths, health_tree_frame, health_tree
    global health_summary_online_badge, health_summary_degraded_badge
    global health_summary_offline_badge, health_summary_unchecked_badge
    global bw_outer, bw_preset_var, bw_preset_menu, bw_custom_entry, bw_save_btn, bw_status_lbl
    global bw_current_card, bw_active_limit_lbl
    global export_outer, export_scope_var, export_pid_var, export_pid_menu
    global export_study_var, export_study_menu, export_series_var, export_series_menu
    global export_files_tree, export_password_var, export_password_entry
    global export_reset_defaults_btn
    global config_bundle_export_btn, config_bundle_import_btn, config_bundle_status_lbl
    global export_files_frame, export_patient_row
    global export_reports_var, export_logs_var, export_metadata_var, export_dicomdir_var
    global export_progress_bar, export_status_lbl, export_start_btn, export_scope_panels
    global reports_outer, reports_type_var, reports_custom_row
    global reports_from_entry, reports_to_entry, reports_generate_btn
    global reports_print_btn, reports_email_btn, reports_status_lbl
    global reports_preview_btn, reports_meta_card, reports_meta_lbl
    global report_sched_enabled_var, report_sched_freq_var, report_sched_hour_var
    global report_sched_save_btn, report_sched_status_lbl
    global perf_outer, perf_cards, perf_graph_canvases, perf_export_btn, perf_status_lbl
    global backup_outer, backup_manual_btn, backup_restore_btn, backup_status_lbl, backup_last_badge
    global backup_next_run_lbl
    global backup_sched_enabled_var, backup_sched_freq_var, backup_sched_hour_var, backup_sched_save_btn
    global backup_history_tree
    global ldap_outer, ldap_enabled_var, ldap_server_entry, ldap_ssl_var, ldap_domain_entry
    global ldap_bind_dn_entry, ldap_bind_pw_entry, ldap_user_base_entry, ldap_user_filter_entry
    global ldap_bind_template_entry, ldap_group_base_entry, ldap_status_lbl
    global ldap_group_map_frame, ldap_group_map_entries, ldap_save_btn, ldap_import_btn
    global ldap_roster_tree
    global settings_outer, settings_refresh_rate_var, settings_perf_mode_var
    global settings_perf_stat_lbls, settings_ui_scale_var, settings_font_scale_var
    global settings_compact_var, settings_remember_geom_var, settings_launch_max_var
    global settings_toast_var, settings_sound_var, settings_notif_duration_var, settings_critical_only_var
    global settings_launch_startup_var, settings_auto_start_recv_var, settings_auto_restart_var, settings_restore_session_var
    global settings_minimize_tray_var, settings_confirm_close_var, settings_check_updates_var
    global settings_log_level_var, settings_diag_lbls
    global settings_hc_var, settings_click_targets_var, settings_reduced_motion_var
    global settings_max_workers_var, settings_max_pushes_var, settings_net_timeout_var
    global settings_cache_limit_var, settings_experimental_var
    global settings_doc_transfer_var, settings_doc_transfer_port_var
    global settings_doc_transfer_chunk_var, settings_doc_transfer_timeout_var
    global settings_admin_idle_timeout_var
    global settings_status_lbl

    # ---- Admin Dashboard tab (first tab in Admin mode) ----
    tab_dashboard = tabview.add("Admin Dashboard")

    dash_outer = ctk.CTkFrame(tab_dashboard, fg_color="transparent")
    dash_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    dash_top_row = make_card_row(dash_outer, pady=(0, 10))

    def _stat_card(parent, title):
        card = ctk.CTkFrame(parent, corner_radius=12, fg_color=THEME_HEADING_BG)
        add_card_to_row(parent, card, padx=6, pady=0)
        ctk.CTkLabel(card, text=title, font=get_font("body")).pack(pady=(10, 0))
        val_lbl = MarqueeLabel(card, text="0", width=1, height=32,
                                font=get_font("kpi_lg", "bold"), text_color=THEME_TEXT,
                                canvas_bg=THEME_HEADING_BG, anchor="center")
        val_lbl.pack(fill="x", padx=10, pady=(0, 10))
        return val_lbl

    dash_received_lbl = _stat_card(dash_top_row, "Received Today")
    dash_pushed_lbl = _stat_card(dash_top_row, "Pushed Successfully Today")
    dash_failed_lbl = _stat_card(dash_top_row, "Failed Today")
    # B.4: patients where DICOM delivery succeeded but the Report/History
    # documents either failed to transfer or were never confirmed
    # delivered at all -- distinct from dash_failed_lbl, which is purely
    # DICOM-status-driven.
    dash_docs_pending_lbl = _stat_card(dash_top_row, "Documents Pending")

    dash_mid_row = ctk.CTkFrame(dash_outer, fg_color="transparent")
    dash_mid_row.pack(fill="x", pady=(0, 10))

    disk_card = make_card(dash_mid_row)
    disk_card.pack(side="left", padx=6, fill="both", expand=True)
    ctk.CTkLabel(disk_card.body, text="Disk Space", font=get_font("caption", "bold")).pack(anchor="w", pady=(4, 2))
    dash_disk_bar = ctk.CTkProgressBar(disk_card.body, width=220)
    dash_disk_bar.set(0)
    dash_disk_bar.pack(pady=4, fill="x")
    dash_disk_lbl = ctk.CTkLabel(disk_card.body, text="—", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    dash_disk_lbl.pack(anchor="w")

    stale_card = make_card(dash_mid_row)
    stale_card.pack(side="left", padx=6, fill="both", expand=True)
    ctk.CTkLabel(stale_card.body, text="Stale Pending Studies", font=get_font("caption", "bold")).pack(anchor="w", pady=(4, 2))
    dash_stale_lbl = ctk.CTkLabel(stale_card.body, text="0", font=get_font("kpi", "bold"), text_color=STALE_HIGHLIGHT_COLOR)
    dash_stale_lbl.pack(anchor="w")
    dash_stale_jump_btn = ctk.CTkButton(stale_card.body, text="Jump to Receiver, filtered", width=230)
    dash_stale_jump_btn.pack(anchor="w", pady=8)

    recv_card = make_card(dash_mid_row)
    recv_card.pack(side="left", padx=6, fill="both", expand=True)
    ctk.CTkLabel(recv_card.body, text="Receiver Status", font=get_font("caption", "bold")).pack(anchor="w", pady=(4, 2))
    dash_receiver_status_badge = make_status_badge(recv_card.body, "Stopped", kind="stopped")
    dash_receiver_status_badge.pack(anchor="w")

    dash_bottom_row = ctk.CTkFrame(dash_outer, fg_color="transparent")
    dash_bottom_row.pack(fill="both", expand=True)

    fail_card = make_card(dash_bottom_row)
    fail_card.pack(side="left", fill="both", expand=True, padx=(0, 6))
    ctk.CTkLabel(fail_card.body, text="Recent Failures (click a row for full error)",
                 font=get_font("caption", "bold")).pack(anchor="w", pady=(0, 4))
    dash_fail_cols = ("patient_id", "patient_name", "time", "last_error")
    dash_fail_headings = {"patient_id": "Patient ID", "patient_name": "Patient Name",
                           "time": "Received", "last_error": "Last Error"}
    dash_fail_widths = {"patient_id": 110, "patient_name": 140, "time": 140, "last_error": 260}
    dash_fail_tree_frame = ctk.CTkFrame(fail_card.body, fg_color=THEME_SURFACE)
    dash_fail_tree_frame.pack(fill="both", expand=True)
    dash_failures_tree = ttk.Treeview(dash_fail_tree_frame, columns=dash_fail_cols, show="headings", height=8)
    for col in dash_fail_cols:
        dash_failures_tree.heading(col, text=dash_fail_headings[col],
                                    command=lambda c=col: sort_tree(dash_failures_tree, c, False))
        dash_failures_tree.column(col, width=dash_fail_widths[col], anchor="w")
    dash_failures_tree.pack(fill="both", expand=True)

    def _on_dash_failure_click(event):
        row_id = dash_failures_tree.identify_row(event.y)
        if not row_id:
            return
        with data_lock:
            info = patient_data.get(row_id, {})
        show_full_error_dialog(row_id, info.get("last_error", "(no error message on file)"))

    dash_failures_tree.bind("<Double-1>", _on_dash_failure_click)

    health_card = make_card(dash_bottom_row)
    health_card.pack(side="left", fill="both", expand=True, padx=(6, 0))
    ctk.CTkLabel(health_card.body, text="Per-Destination Push Health",
                 font=get_font("caption", "bold")).pack(anchor="w", pady=(0, 4))
    dash_health_cols = ("destination", "last_echo", "push_ok", "push_failed")
    dash_health_headings = {"destination": "Destination", "last_echo": "Last C-ECHO",
                             "push_ok": "Sent", "push_failed": "Failed"}
    dash_health_widths = {"destination": 150, "last_echo": 160, "push_ok": 70, "push_failed": 70}
    dash_health_tree_frame = ctk.CTkFrame(health_card.body, fg_color=THEME_SURFACE)
    dash_health_tree_frame.pack(fill="both", expand=True)
    dash_dest_health_tree = ttk.Treeview(dash_health_tree_frame, columns=dash_health_cols, show="headings", height=8)
    for col in dash_health_cols:
        dash_dest_health_tree.heading(col, text=dash_health_headings[col],
                                       command=lambda c=col: sort_tree(dash_dest_health_tree, c, False))
        dash_dest_health_tree.column(col, width=dash_health_widths[col], anchor="w")
    dash_dest_health_tree.pack(fill="both", expand=True)

    def _on_dash_dest_echo(event):
        """Double-click a destination row to run an on-demand C-ECHO test.
        Intentionally on-click, not automatic -- the spec asks us not to add
        new background polling. Runs off the UI thread like the existing
        Destinations-tab C-ECHO test does."""
        row_id = dash_dest_health_tree.identify_row(event.y)
        if not row_id:
            return
        dest = get_destination_by_name(row_id)
        if not dest:
            return
        dash_dest_health_tree.set(row_id, "last_echo", "Testing...")

        def run():
            ok, msg = dicom_echo(dest["ae"], dest["ip"], int(dest["port"]),
                                  calling_ae=dest.get("calling_ae") or DEFAULT_PUSH_CALLING_AE,
                                  dest_label=dest.get("name"))

            def on_ui():
                try:
                    dash_dest_health_tree.set(row_id, "last_echo", ("OK" if ok else f"Failed: {msg}")[:40])
                except Exception:
                    pass  # row/tab may have been torn down (e.g. mode switch) mid-test
            app.after(0, on_ui)

        threading.Thread(target=run, daemon=True).start()

    dash_dest_health_tree.bind("<Double-1>", _on_dash_dest_echo)
    ctk.CTkLabel(health_card, text="(Double-click a row to run an on-demand C-ECHO test)",
                 font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=10, pady=(0, 8))

    dash_btn_row = ctk.CTkFrame(dash_outer, fg_color="transparent")
    dash_btn_row.pack(fill="x", pady=(6, 0))
    dash_refresh_btn = ctk.CTkButton(dash_btn_row, text="Refresh Dashboard", width=180, image=get_icon("refresh-cw", size=14, color=THEME_TEXT), compound="left")
    dash_refresh_btn.pack(side="left", padx=4)
    dash_export_btn = ctk.CTkButton(dash_btn_row, text="Export audit log to CSV", width=200)
    dash_export_btn.pack(side="left", padx=4)

    dash_refresh_btn.configure(command=refresh_admin_dashboard)
    dash_export_btn.configure(command=do_export_audit_log_csv)
    dash_stale_jump_btn.configure(command=jump_to_receiver_stale_filtered)

    # ---- PACS Health Monitor tab (continuous, scheduled C-ECHO to every
    # configured destination -- distinct from the Admin Dashboard's
    # on-demand double-click test above). Runs on its own background
    # timer (see pacs_health_monitor_tick / start_pacs_health_monitor_thread)
    # so it keeps working even while this tab isn't visible, and shares
    # the same destination_health_cache the landing-page Dashboard reads. ----
    tab_health = tabview.add("PACS Health")

    health_outer = ctk.CTkFrame(tab_health, fg_color="transparent")
    health_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    health_top = ctk.CTkFrame(health_outer, fg_color="transparent")
    health_top.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(health_top, text="PACS Health Monitor", font=get_font("section_lg", "bold")).pack(side="left")
    health_run_all_btn = ctk.CTkButton(health_top, text="Check All Now", width=150)
    health_run_all_btn.pack(side="right", padx=4)
    health_status_lbl = ctk.CTkLabel(health_top, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    health_status_lbl.pack(side="right", padx=10)

    health_summary_row = make_card_row(health_outer, pady=(0, 10))

    def _health_summary_card(title, kind):
        card = make_card(health_summary_row)
        add_card_to_row(health_summary_row, card)
        ctk.CTkLabel(card.body, text=title, font=get_font("micro"),
                     text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 4))
        badge = make_status_badge(card.body, "0", kind=kind)
        badge.pack(anchor="w")
        return badge

    health_summary_online_badge = _health_summary_card("Online", "connected")
    health_summary_degraded_badge = _health_summary_card("Degraded", "warning")
    health_summary_offline_badge = _health_summary_card("Offline", "offline")
    health_summary_unchecked_badge = _health_summary_card("Not Yet Checked", "pending")

    health_cols = ("destination", "status", "response_time", "avg_latency",
                   "last_success", "last_failure", "consecutive_failures", "uptime_pct")
    health_headings = {
        "destination": "Destination", "status": "Status", "response_time": "Response Time",
        "avg_latency": "Avg Latency", "last_success": "Last Success", "last_failure": "Last Failure",
        "consecutive_failures": "Consecutive Failures", "uptime_pct": "Uptime %",
    }
    health_widths = {
        "destination": 160, "status": 110, "response_time": 110, "avg_latency": 110,
        "last_success": 150, "last_failure": 150, "consecutive_failures": 140, "uptime_pct": 90,
    }
    health_tree_frame = ctk.CTkFrame(health_outer, fg_color=THEME_SURFACE)
    health_tree_frame.pack(fill="both", expand=True, padx=2, pady=(0, 8))
    health_tree = ttk.Treeview(health_tree_frame, columns=health_cols, show="headings", height=14)
    for col in health_cols:
        health_tree.heading(col, text=health_headings[col],
                             command=lambda c=col: sort_tree(health_tree, c, False))
        health_tree.column(col, width=health_widths[col], anchor="w")
    health_tree.pack(fill="both", expand=True)
    health_tree.tag_configure("status_online", foreground=THEME_SUCCESS)
    health_tree.tag_configure("status_offline", foreground=THEME_DANGER)
    health_tree.tag_configure("status_degraded", foreground="#f1c40f")

    make_wrapped_label(
        health_outer,
        "Automatically runs a C-ECHO against every configured destination every "
        f"{PACS_HEALTH_CHECK_INTERVAL_SEC}s. Green = online, Yellow = degraded "
        "(1-2 recent failures), Red = offline (3+ consecutive failures).",
        900, font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", padx=4, pady=(0, 6))

    health_run_all_btn.configure(command=do_run_all_health_checks_now)

    # ---- Bandwidth Limiter tab ----
    tab_bandwidth = tabview.add("Bandwidth")
    bw_outer = ctk.CTkFrame(tab_bandwidth, fg_color="transparent")
    bw_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(bw_outer, text="Bandwidth Limiter", font=get_font("section_lg", "bold")).pack(anchor="w")
    make_wrapped_label(
        bw_outer,
        "Throttles DICOM Push, Import, and Export so this application never saturates "
        "the network link. Applies as a shared, long-run average rate across all "
        "active transfers -- it never freezes the UI while waiting.",
        700, font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(4, 14))

    bw_card = make_card(bw_outer)
    bw_card.pack(fill="x")
    bw_preset_frame = bw_card.body
    bw_preset_frame.grid_columnconfigure(5, weight=1)

    ctk.CTkLabel(bw_preset_frame, text="Limit:", font=get_font("body")).grid(
        row=0, column=0, padx=(0, 8), pady=12, sticky="w")
    bw_preset_var = ctk.StringVar(value=load_bandwidth_config().get("preset", "Unlimited"))
    bw_preset_menu = ctk.CTkOptionMenu(
        bw_preset_frame, variable=bw_preset_var, width=160,
        values=["Unlimited", "1 Mbps", "5 Mbps", "10 Mbps", "25 Mbps", "50 Mbps", "100 Mbps", "Custom"],
    )
    bw_preset_menu.grid(row=0, column=1, padx=8, pady=12, sticky="w")

    ctk.CTkLabel(bw_preset_frame, text="Custom (Mbps):", font=get_font("body")).grid(
        row=0, column=2, padx=(20, 8), pady=12, sticky="w")
    bw_custom_entry = ctk.CTkEntry(bw_preset_frame, width=100)
    bw_custom_entry.insert(0, str(load_bandwidth_config().get("custom_mbps", 10)))
    bw_custom_entry.grid(row=0, column=3, padx=8, pady=12, sticky="w")

    bw_save_btn = ctk.CTkButton(bw_preset_frame, text="Save", width=100,
                                 image=get_icon("check-check", size=14, color="#ffffff"), compound="left")
    bw_save_btn.grid(row=0, column=4, padx=(20, 0), pady=12)

    bw_current_card = make_card(bw_outer)
    bw_current_card.pack(fill="x", pady=(10, 0))
    ctk.CTkLabel(bw_current_card.body, text="Active Limit", font=get_font("micro"),
                 text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 2))
    bw_active_limit_lbl = ctk.CTkLabel(bw_current_card.body, text="—", font=get_font("title", "bold"),
                                        text_color=THEME_TEXT)
    bw_active_limit_lbl.pack(anchor="w")

    bw_status_lbl = ctk.CTkLabel(bw_outer, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    bw_status_lbl.pack(anchor="w", pady=(8, 0))

    def _bw_sync_custom_entry_state(*_args):
        bw_custom_entry.configure(state="normal" if bw_preset_var.get() == "Custom" else "disabled")

    def _bw_refresh_active_limit_display():
        effective = get_effective_bandwidth_mbps()
        bw_active_limit_lbl.configure(text="Unlimited" if effective <= 0 else f"{effective:g} Mbps")

    bw_preset_var.trace_add("write", _bw_sync_custom_entry_state)
    _bw_sync_custom_entry_state()
    _bw_refresh_active_limit_display()

    bw_save_btn.configure(command=do_save_bandwidth_config)

    # ---- Export tab ----
    tab_export = tabview.add("Export")
    export_outer = ctk.CTkScrollableFrame(tab_export, fg_color="transparent")
    export_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(export_outer, text="Built-in ZIP Export", font=get_font("section_lg", "bold")).pack(anchor="w")
    make_wrapped_label(
        export_outer,
        "Export an entire patient, a single study or series, a hand-picked list of files, "
        "or the whole worklist -- as a plain ZIP or a password-protected (AES-256) ZIP, "
        "optionally bundling the report, a filtered log excerpt, a metadata summary, and/or "
        "a real DICOMDIR.",
        760, font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(4, 14))

    _export_tpl = load_export_template()
    export_scope_var = ctk.StringVar(value=_export_tpl.get("scope", "Entire Patient"))
    export_scope_row = ctk.CTkFrame(export_outer, fg_color="transparent")
    export_scope_row.pack(fill="x", pady=(0, 10))
    ctk.CTkLabel(export_scope_row, text="Export:", font=get_font("body")).pack(side="left", padx=(0, 8))
    export_scope_menu = ctk.CTkSegmentedButton(
        export_scope_row, variable=export_scope_var,
        values=["Entire Patient", "Study", "Series", "Selected Files", "Entire Worklist"],
    )
    export_scope_menu.pack(side="left")

    # ---- Patient / Study / Series pickers (shown/hidden per scope) ----
    export_patient_row = ctk.CTkFrame(export_outer)
    export_patient_row.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(export_patient_row, text="Patient ID:", font=get_font("body")).grid(
        row=0, column=0, padx=(12, 8), pady=10, sticky="w")
    export_pid_var = ctk.StringVar(value="")
    export_pid_menu = ctk.CTkOptionMenu(export_patient_row, variable=export_pid_var, width=200, values=["(no patients)"])
    export_pid_menu.grid(row=0, column=1, padx=8, pady=10, sticky="w")

    ctk.CTkLabel(export_patient_row, text="Study:", font=get_font("body")).grid(
        row=0, column=2, padx=(20, 8), pady=10, sticky="w")
    export_study_var = ctk.StringVar(value="")
    export_study_menu = ctk.CTkOptionMenu(export_patient_row, variable=export_study_var, width=260, values=["(select a patient)"])
    export_study_menu.grid(row=0, column=3, padx=8, pady=10, sticky="w")

    ctk.CTkLabel(export_patient_row, text="Series:", font=get_font("body")).grid(
        row=0, column=4, padx=(20, 8), pady=10, sticky="w")
    export_series_var = ctk.StringVar(value="")
    export_series_menu = ctk.CTkOptionMenu(export_patient_row, variable=export_series_var, width=260, values=["(select a patient)"])
    export_series_menu.grid(row=0, column=5, padx=8, pady=10, sticky="w")

    # ---- Selected Files picker (multi-select tree of this patient's individual instances) ----
    export_files_frame = ctk.CTkFrame(export_outer, fg_color=THEME_SURFACE)
    export_files_frame.pack(fill="x", pady=(0, 8))
    export_files_hdr_row = ctk.CTkFrame(export_files_frame, fg_color="transparent")
    export_files_hdr_row.pack(fill="x", padx=8, pady=(6, 2))
    ctk.CTkLabel(export_files_hdr_row, text="Ctrl/Shift-click to select individual files:",
                 font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(side="left")
    export_files_select_all_btn = ctk.CTkButton(
        export_files_hdr_row, text="Select All", width=90, height=24,
        font=get_font("small"), fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
        command=lambda: export_files_tree.selection_set(export_files_tree.get_children("")))
    export_files_select_all_btn.pack(side="right", padx=(4, 0))
    export_files_invert_btn = ctk.CTkButton(
        export_files_hdr_row, text="Invert", width=80, height=24,
        font=get_font("small"), fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
        command=lambda: export_files_tree.selection_set(
            [iid for iid in export_files_tree.get_children("") if iid not in export_files_tree.selection()]))
    export_files_invert_btn.pack(side="right", padx=(4, 0))
    export_files_clear_btn = ctk.CTkButton(
        export_files_hdr_row, text="Clear", width=80, height=24,
        font=get_font("small"), fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
        command=lambda: export_files_tree.selection_remove(export_files_tree.selection()))
    export_files_clear_btn.pack(side="right", padx=(4, 0))
    export_files_cols = ("sop_uid", "series_uid", "modality", "size")
    export_files_tree = ttk.Treeview(export_files_frame, columns=export_files_cols, show="headings",
                                      height=6, selectmode="extended")
    for col, label, w in (("sop_uid", "SOP Instance UID", 280), ("series_uid", "Series UID", 260),
                          ("modality", "Modality", 90), ("size", "Size", 90)):
        export_files_tree.heading(col, text=label)
        export_files_tree.column(col, width=w, anchor="w")
    export_files_tree.pack(fill="x", padx=8, pady=(0, 8))
    export_files_tree.bind("<Control-a>", lambda e: (export_files_tree.selection_set(export_files_tree.get_children("")), "break")[1])
    export_files_tree.bind("<Control-A>", lambda e: (export_files_tree.selection_set(export_files_tree.get_children("")), "break")[1])
    export_files_tree.bind("<Escape>", lambda e: (export_files_tree.selection_remove(export_files_tree.selection()), "break")[1])

    # ---- Options ----
    export_options_frame = ctk.CTkFrame(export_outer)
    export_options_frame.pack(fill="x", pady=(4, 8))
    ctk.CTkLabel(export_options_frame, text="Options", font=get_font("body", "bold")).grid(
        row=0, column=0, columnspan=4, padx=12, pady=(10, 2), sticky="w")

    export_reports_var = ctk.BooleanVar(value=_export_tpl.get("include_reports", False))
    ctk.CTkCheckBox(export_options_frame, text="Include Reports", variable=export_reports_var).grid(
        row=1, column=0, padx=12, pady=6, sticky="w")
    export_logs_var = ctk.BooleanVar(value=_export_tpl.get("include_logs", False))
    ctk.CTkCheckBox(export_options_frame, text="Include Logs", variable=export_logs_var).grid(
        row=1, column=1, padx=12, pady=6, sticky="w")
    export_metadata_var = ctk.BooleanVar(value=_export_tpl.get("include_metadata", False))
    ctk.CTkCheckBox(export_options_frame, text="Include Metadata", variable=export_metadata_var).grid(
        row=1, column=2, padx=12, pady=6, sticky="w")
    export_dicomdir_var = ctk.BooleanVar(value=_export_tpl.get("include_dicomdir", False))
    ctk.CTkCheckBox(export_options_frame, text="Include DICOMDIR", variable=export_dicomdir_var).grid(
        row=1, column=3, padx=12, pady=6, sticky="w")

    ctk.CTkLabel(export_options_frame, text="Password (leave blank for a plain ZIP):",
                 font=get_font("body")).grid(row=2, column=0, columnspan=2, padx=12, pady=(6, 12), sticky="w")
    export_password_var = ctk.StringVar(value="")
    export_password_entry = ctk.CTkEntry(export_options_frame, textvariable=export_password_var,
                                          width=220, show="•")
    export_password_entry.grid(row=2, column=2, columnspan=2, padx=12, pady=(6, 12), sticky="w")

    # ---- Run + progress ----
    export_run_row = ctk.CTkFrame(export_outer, fg_color="transparent")
    export_run_row.pack(fill="x", pady=(4, 4))
    export_start_btn = ctk.CTkButton(export_run_row, text="Export…", width=160,
                                      fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    export_start_btn.pack(side="left")
    export_progress_bar = ctk.CTkProgressBar(export_run_row, width=300)
    export_progress_bar.set(0)
    export_progress_bar.pack(side="left", padx=14)
    export_reset_defaults_btn = ctk.CTkButton(export_run_row, text="Reset to Defaults", width=150,
                                               fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
    export_reset_defaults_btn.pack(side="left", padx=(4, 0))
    export_status_lbl = ctk.CTkLabel(export_outer, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    export_status_lbl.pack(anchor="w", pady=(4, 0))

    if not PYZIPPER_AVAILABLE:
        ctk.CTkLabel(
            export_outer,
            text="pyzipper is not installed -- password-protected export will be unavailable "
                 "until you run: pip install pyzipper",
            image=get_icon("triangle-alert", size=14, color=THEME_WARNING), compound="left",
            font=get_font("small"), text_color=THEME_WARNING,
        ).pack(anchor="w", pady=(6, 0))

    export_scope_panels = {
        "patient": [export_pid_menu],
        "study": [export_pid_menu, export_study_menu],
        "series": [export_pid_menu, export_series_menu],
        "files": [export_pid_menu, export_files_frame],
        "worklist": [],
    }

    export_scope_var.trace_add("write", lambda *_: _update_export_scope_visibility())
    export_pid_var.trace_add("write", lambda *_: refresh_export_study_series_options())
    export_start_btn.configure(command=do_start_export)
    export_reset_defaults_btn.configure(command=do_reset_export_defaults)

    # 9.1 -- Full-config export/import bundle (Admin-only, this whole tab
    # already is). Separate section from the per-patient ZIP export above.
    config_bundle_frame = ctk.CTkFrame(export_outer)
    config_bundle_frame.pack(fill="x", pady=(20, 8))
    ctk.CTkLabel(config_bundle_frame, text="Full Configuration Bundle", font=get_font("caption", "bold")).pack(
        anchor="w", padx=12, pady=(10, 2))
    ctk.CTkLabel(
        config_bundle_frame,
        text="Bundles Destinations, Routing Rules, LDAP, Bandwidth, TLS, Notifications, Backup Schedule, "
             "and Log Retention into one file for backup or moving to another instance. Because this includes "
             "LDAP and SMTP credentials, the bundle is encrypted with this installation's key.key -- importing "
             "it elsewhere requires that same key.key.",
        font=get_font("small"), text_color=THEME_TEXT_MUTED, wraplength=760, justify="left",
    ).pack(anchor="w", padx=12, pady=(0, 8))
    config_bundle_btn_row = ctk.CTkFrame(config_bundle_frame, fg_color="transparent")
    config_bundle_btn_row.pack(fill="x", padx=12, pady=(0, 10))
    config_bundle_export_btn = ctk.CTkButton(config_bundle_btn_row, text="Export App Configuration…", width=220)
    config_bundle_export_btn.pack(side="left")
    config_bundle_import_btn = ctk.CTkButton(config_bundle_btn_row, text="Import Configuration…", width=200,
                                              fg_color=THEME_WARNING, hover_color=THEME_WARNING_HOVER)
    config_bundle_import_btn.pack(side="left", padx=8)
    config_bundle_status_lbl = ctk.CTkLabel(config_bundle_frame, text="", font=get_font("small"),
                                             text_color=THEME_TEXT_MUTED)
    config_bundle_status_lbl.pack(anchor="w", padx=12, pady=(0, 10))
    config_bundle_export_btn.configure(command=do_export_config_bundle)
    config_bundle_import_btn.configure(command=do_import_config_bundle)

    # ---- Reports tab ----
    tab_reports = tabview.add("Reports")
    reports_outer = ctk.CTkFrame(tab_reports, fg_color="transparent")
    reports_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(reports_outer, text="PDF Report Generator", font=get_font("section_lg", "bold")).pack(anchor="w")
    make_wrapped_label(
        reports_outer,
        "Studies received/sent, failed studies, success rate, average transfer speed, "
        "top modalities/institutions, destination statistics, queue stats, disk usage, and "
        "uptime -- as a professional PDF with tables and charts.",
        760, font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(4, 14))

    reports_type_var = ctk.StringVar(value="Daily")
    reports_type_row = ctk.CTkFrame(reports_outer, fg_color="transparent")
    reports_type_row.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(reports_type_row, text="Report:", font=get_font("body")).pack(side="left", padx=(0, 8))
    reports_type_menu = ctk.CTkSegmentedButton(
        reports_type_row, variable=reports_type_var,
        values=["Daily", "Weekly", "Monthly", "Year to Date", "Custom Date Range"],
    )
    reports_type_menu.pack(side="left")

    reports_custom_row = ctk.CTkFrame(reports_outer)
    ctk.CTkLabel(reports_custom_row, text="From (YYYY-MM-DD):", font=get_font("body")).grid(
        row=0, column=0, padx=(12, 6), pady=10, sticky="w")
    reports_from_entry = ctk.CTkEntry(reports_custom_row, width=140,
                                      placeholder_text=datetime.date.today().isoformat())
    reports_from_entry.grid(row=0, column=1, padx=6, pady=10, sticky="w")
    ctk.CTkLabel(reports_custom_row, text="To (YYYY-MM-DD):", font=get_font("body")).grid(
        row=0, column=2, padx=(20, 6), pady=10, sticky="w")
    reports_to_entry = ctk.CTkEntry(reports_custom_row, width=140,
                                    placeholder_text=datetime.date.today().isoformat())
    reports_to_entry.grid(row=0, column=3, padx=6, pady=10, sticky="w")

    reports_run_row = ctk.CTkFrame(reports_outer, fg_color="transparent")
    reports_run_row.pack(fill="x", pady=(10, 4))
    reports_generate_btn = ctk.CTkButton(reports_run_row, text="Generate PDF…", width=170,
                                         fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    reports_generate_btn.pack(side="left")
    reports_preview_btn = ctk.CTkButton(reports_run_row, text="Preview", width=130, state="disabled",
                                         fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
    reports_preview_btn.pack(side="left", padx=8)
    reports_print_btn = ctk.CTkButton(reports_run_row, text="Print Last Report", width=170, state="disabled")
    reports_print_btn.pack(side="left", padx=8)
    reports_email_btn = ctk.CTkButton(reports_run_row, text="Email Last Report", width=170, state="disabled")
    reports_email_btn.pack(side="left", padx=8)

    reports_status_lbl = ctk.CTkLabel(reports_outer, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    reports_status_lbl.pack(anchor="w", pady=(8, 0))

    reports_meta_card = make_card(reports_outer, title="Last Report")
    reports_meta_card.pack(fill="x", pady=(14, 0))
    reports_meta_lbl = ctk.CTkLabel(reports_meta_card.body, text="No report generated yet this session.",
                                     font=get_font("small"), text_color=THEME_TEXT_MUTED, justify="left")
    reports_meta_lbl.pack(anchor="w")

    # 8.2 -- Scheduled report email: same lightweight scheduler pattern as
    # Backup's "Scheduled Backup" section, applied to periodically
    # generating + emailing a report PDF instead. Uses the SMTP settings
    # already configured for notifications (load_notifications_config()).
    _report_sched_cfg = load_report_email_schedule()
    report_sched_frame = ctk.CTkFrame(reports_outer)
    report_sched_frame.pack(fill="x", pady=(14, 0))
    ctk.CTkLabel(report_sched_frame, text="Scheduled Report Email", font=get_font("body", "bold")).grid(
        row=0, column=0, columnspan=5, padx=12, pady=(10, 4), sticky="w")

    report_sched_enabled_var = ctk.BooleanVar(value=_report_sched_cfg.get("enabled", False))
    ctk.CTkCheckBox(report_sched_frame, text="Enabled", variable=report_sched_enabled_var).grid(
        row=1, column=0, padx=12, pady=6, sticky="w")

    ctk.CTkLabel(report_sched_frame, text="Frequency:", font=get_font("body")).grid(
        row=1, column=1, padx=(12, 6), pady=6, sticky="w")
    report_sched_freq_var = ctk.StringVar(value=_report_sched_cfg.get("frequency", "Daily"))
    ctk.CTkOptionMenu(report_sched_frame, variable=report_sched_freq_var, width=110,
                       values=["Daily", "Weekly"]).grid(row=1, column=2, padx=6, pady=6, sticky="w")

    ctk.CTkLabel(report_sched_frame, text="Hour (0-23):", font=get_font("body")).grid(
        row=1, column=3, padx=(12, 6), pady=6, sticky="w")
    report_sched_hour_var = ctk.StringVar(value=str(_report_sched_cfg.get("hour", 6)))
    ctk.CTkEntry(report_sched_frame, textvariable=report_sched_hour_var, width=60).grid(
        row=1, column=4, padx=6, pady=6, sticky="w")

    report_sched_save_btn = ctk.CTkButton(report_sched_frame, text="Save Schedule", width=140)
    report_sched_save_btn.grid(row=1, column=5, padx=12, pady=6, sticky="w")

    report_sched_status_lbl = ctk.CTkLabel(report_sched_frame, text="", font=get_font("small"),
                                            text_color=THEME_TEXT_MUTED)
    report_sched_status_lbl.grid(row=2, column=0, columnspan=6, padx=12, pady=(0, 10), sticky="w")
    report_sched_save_btn.configure(command=do_save_report_email_schedule)

    if not REPORTLAB_AVAILABLE:
        ctk.CTkLabel(
            reports_outer,
            text="'reportlab' and/or 'matplotlib' are not installed -- PDF report generation is "
                 "unavailable until you run: pip install reportlab matplotlib",
            image=get_icon("triangle-alert", size=14, color=THEME_WARNING), compound="left",
            font=get_font("small"), text_color=THEME_WARNING,
        ).pack(anchor="w", pady=(6, 0))

    def _update_reports_custom_row_visibility(*_args):
        if reports_type_var.get() == "Custom Date Range":
            reports_custom_row.pack(fill="x", pady=(0, 8))
        else:
            reports_custom_row.pack_forget()

    reports_type_var.trace_add("write", _update_reports_custom_row_visibility)
    reports_generate_btn.configure(command=do_generate_report)
    reports_preview_btn.configure(command=do_preview_last_report)
    reports_print_btn.configure(command=do_print_last_report)
    reports_email_btn.configure(command=do_email_last_report)
    if reports_last_path["value"] and os.path.isfile(reports_last_path["value"]):
        reports_preview_btn.configure(state="normal")
        reports_print_btn.configure(state="normal")
        reports_email_btn.configure(state="normal")
    _refresh_reports_meta_display()

    # ---- Performance Metrics tab ----
    tab_perf = tabview.add("Performance")
    perf_outer = ctk.CTkScrollableFrame(tab_perf, fg_color="transparent")
    perf_outer.pack(fill="both", expand=True, padx=4, pady=4)

    perf_top_row = ctk.CTkFrame(perf_outer, fg_color="transparent")
    perf_top_row.pack(fill="x", padx=10, pady=(8, 4))
    ctk.CTkLabel(perf_top_row, text="Performance Metrics", font=get_font("section_lg", "bold")).pack(side="left")
    perf_export_btn = ctk.CTkButton(perf_top_row, text="Export Metrics…", width=160)
    perf_export_btn.pack(side="right")
    perf_status_lbl = ctk.CTkLabel(perf_top_row, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    perf_status_lbl.pack(side="right", padx=10)

    perf_cards = {}

    def _perf_cards_row(parent):
        return make_card_row(parent)

    def _perf_card(parent, key, title):
        card = make_card(parent)
        add_card_to_row(parent, card)
        ctk.CTkLabel(card.body, text=title, font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 0))
        val_lbl = MarqueeLabel(card.body, text="—", width=1, height=26,
                               font=get_font("value_xl", "bold"), text_color=THEME_TEXT,
                               canvas_bg=THEME_HEADING_BG)
        val_lbl.pack(anchor="w", fill="x", pady=(2, 0))
        perf_cards[key] = val_lbl

    perf_row1 = _perf_cards_row(perf_outer)
    _perf_card(perf_row1, "images_per_sec", "Images/sec")
    _perf_card(perf_row1, "studies_per_sec", "Studies/sec")
    _perf_card(perf_row1, "mb_per_sec", "MB/sec")
    _perf_card(perf_row1, "avg_receive_time", "Avg Receive Time")
    _perf_card(perf_row1, "avg_push_time", "Avg Push Time")

    perf_row2 = _perf_cards_row(perf_outer)
    _perf_card(perf_row2, "avg_association_time", "Avg Association Time")
    _perf_card(perf_row2, "avg_queue_time", "Avg Queue Time")
    _perf_card(perf_row2, "avg_retry_time", "Avg Retry Time")
    _perf_card(perf_row2, "cpu_pct", "CPU Usage")
    _perf_card(perf_row2, "ram_pct", "RAM Usage")

    perf_row3 = _perf_cards_row(perf_outer)
    _perf_card(perf_row3, "disk_io", "Disk I/O")
    _perf_card(perf_row3, "network_io", "Network Usage")
    _perf_card(perf_row3, "database_size", "Database Size")
    _perf_card(perf_row3, "uptime", "Application Uptime")

    # B.1: Document Transfer cards -- derived from the same structured
    # push/receiver JSONL logs on refresh, no new counters.
    perf_row4 = _perf_cards_row(perf_outer)
    _perf_card(perf_row4, "docs_per_sec", "Docs/sec")
    _perf_card(perf_row4, "doc_success_rate_pct", "Doc Transfer Success Rate")
    _perf_card(perf_row4, "avg_doc_transfer_time", "Avg Doc Transfer Time")

    if not PSUTIL_AVAILABLE:
        ctk.CTkLabel(
            perf_outer,
            text="psutil is not installed -- CPU/RAM/Disk I/O/Network metrics will show as N/A.",
            image=get_icon("triangle-alert", size=14, color=THEME_WARNING), compound="left",
            font=get_font("small"), text_color=THEME_WARNING,
        ).pack(anchor="w", padx=16, pady=(0, 4))

    ctk.CTkLabel(perf_outer, text="Historical Graphs (last 2 minutes, auto-refreshing)",
                 font=get_font("caption", "bold")).pack(anchor="w", padx=10, pady=(14, 4))
    perf_graphs_row = ctk.CTkScrollableFrame(perf_outer, fg_color="transparent",
                                              orientation="horizontal", height=200)
    perf_graphs_row.pack(fill="x", padx=6, pady=(0, 12))

    perf_graph_canvases = {}
    PERF_GRAPH_COLORS = {
        "images_per_sec": "#2f8eff", "studies_per_sec": "#2ecc71",
        "cpu_pct": "#f1c40f", "ram_pct": "#9b59b6",
        "disk_io_kbps": "#e67e22", "network_kbps": "#f04747",
    }

    def _perf_sparkline(parent, key, title):
        card = make_card(parent)
        card.pack(side="left", padx=5, pady=5, fill="both", expand=True)
        ctk.CTkLabel(card.body, text=title, font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 0))
        canvas = ctk.CTkCanvas(card.body, width=180, height=110, bg=THEME_SURFACE, highlightthickness=0)
        canvas.pack(pady=(4, 2))
        perf_graph_canvases[key] = (canvas, PERF_GRAPH_COLORS[key])

    _perf_sparkline(perf_graphs_row, "images_per_sec", "Images/sec")
    _perf_sparkline(perf_graphs_row, "studies_per_sec", "Studies/sec")
    _perf_sparkline(perf_graphs_row, "cpu_pct", "CPU %")
    _perf_sparkline(perf_graphs_row, "ram_pct", "RAM %")
    _perf_sparkline(perf_graphs_row, "disk_io_kbps", "Disk I/O (KB/s)")
    _perf_sparkline(perf_graphs_row, "network_kbps", "Network (KB/s)")

    perf_export_btn.configure(command=do_export_performance_metrics)

    # ---- Backup & Restore tab ----
    tab_backup = tabview.add("Backup")
    backup_outer = ctk.CTkScrollableFrame(tab_backup, fg_color="transparent")
    backup_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(backup_outer, text="Backup & Restore", font=get_font("section_lg", "bold")).pack(anchor="w")
    make_wrapped_label(
        backup_outer,
        "Backs up configuration, encryption keys, routing rules, destination profiles, TLS "
        "settings, reports, the CSV database, logs, and audit logs into one validated ZIP. "
        "Restoring automatically takes a safety backup of the current state first.",
        760, font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(4, 14))

    backup_run_row = ctk.CTkFrame(backup_outer, fg_color="transparent")
    backup_run_row.pack(fill="x", pady=(0, 8))
    backup_manual_btn = ctk.CTkButton(backup_run_row, text="Backup Now…", width=160,
                                      fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    backup_manual_btn.pack(side="left")
    backup_restore_btn = ctk.CTkButton(backup_run_row, text="Restore Backup…", width=170,
                                       fg_color=THEME_WARNING, hover_color=THEME_WARNING_HOVER)
    backup_restore_btn.pack(side="left", padx=8)
    backup_last_badge = make_status_badge(backup_run_row, "No backups yet", kind="pending")
    backup_last_badge.pack(side="left", padx=16)
    backup_next_run_lbl = ctk.CTkLabel(backup_run_row, text="", font=get_font("small"),
                                        text_color=THEME_TEXT_MUTED)
    backup_next_run_lbl.pack(side="left", padx=4)
    backup_status_lbl = ctk.CTkLabel(backup_outer, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    backup_status_lbl.pack(anchor="w", pady=(4, 12))

    backup_sched_frame = ctk.CTkFrame(backup_outer)
    backup_sched_frame.pack(fill="x", pady=(0, 12))
    ctk.CTkLabel(backup_sched_frame, text="Scheduled Backup", font=get_font("body", "bold")).grid(
        row=0, column=0, columnspan=5, padx=12, pady=(10, 4), sticky="w")

    _sched_cfg = load_backup_schedule()
    backup_sched_enabled_var = ctk.BooleanVar(value=_sched_cfg.get("enabled", False))
    ctk.CTkCheckBox(backup_sched_frame, text="Enabled", variable=backup_sched_enabled_var).grid(
        row=1, column=0, padx=12, pady=(4, 12), sticky="w")

    ctk.CTkLabel(backup_sched_frame, text="Frequency:", font=get_font("body")).grid(
        row=1, column=1, padx=(12, 6), pady=(4, 12), sticky="w")
    backup_sched_freq_var = ctk.StringVar(value=_sched_cfg.get("frequency", "Daily"))
    ctk.CTkOptionMenu(backup_sched_frame, variable=backup_sched_freq_var, width=110,
                      values=["Daily", "Weekly"]).grid(row=1, column=2, padx=6, pady=(4, 12), sticky="w")

    ctk.CTkLabel(backup_sched_frame, text="At hour (0-23):", font=get_font("body")).grid(
        row=1, column=3, padx=(12, 6), pady=(4, 12), sticky="w")
    backup_sched_hour_var = ctk.StringVar(value=str(_sched_cfg.get("hour", 2)))
    ctk.CTkEntry(backup_sched_frame, textvariable=backup_sched_hour_var, width=60).grid(
        row=1, column=4, padx=6, pady=(4, 12), sticky="w")

    backup_sched_save_btn = ctk.CTkButton(backup_sched_frame, text="Save Schedule", width=130)
    backup_sched_save_btn.grid(row=1, column=5, padx=(20, 12), pady=(4, 12))

    ctk.CTkLabel(backup_outer, text="Backup History", font=get_font("caption", "bold")).pack(
        anchor="w", pady=(4, 4))
    backup_history_cols = ("timestamp", "trigger", "size", "files", "validation", "path")
    backup_history_frame = ctk.CTkFrame(backup_outer, fg_color=THEME_SURFACE)
    backup_history_frame.pack(fill="both", expand=True, pady=(0, 8))
    backup_history_tree = ttk.Treeview(backup_history_frame, columns=backup_history_cols, show="headings", height=8)
    for col, label, w in (("timestamp", "Timestamp", 150), ("trigger", "Trigger", 110),
                          ("size", "Size", 90), ("files", "Files", 60),
                          ("validation", "Validation", 200), ("path", "Path", 260)):
        backup_history_tree.heading(col, text=label,
                                     command=lambda c=col: sort_tree(backup_history_tree, c, False))
        backup_history_tree.column(col, width=w, anchor="w")
    backup_history_tree.pack(fill="both", expand=True)
    backup_history_tree.tag_configure("validation_ok", foreground=THEME_SUCCESS)
    backup_history_tree.tag_configure("validation_fail", foreground=THEME_DANGER)

    ctk.CTkLabel(backup_outer, text="(Double-click a history row to restore that backup)",
                 font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 4))

    backup_manual_btn.configure(command=do_manual_backup)
    backup_restore_btn.configure(command=do_restore_backup_dialog)
    backup_sched_save_btn.configure(command=do_save_backup_schedule)
    backup_history_tree.bind("<Double-1>", _on_backup_history_row_double_click)

    # ---- LDAP / Active Directory tab ----
    tab_ldap = tabview.add("LDAP / AD")
    ldap_outer = ctk.CTkScrollableFrame(tab_ldap, fg_color="transparent")
    ldap_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(ldap_outer, text="LDAP / Active Directory Authentication",
                font=get_font("section_lg", "bold")).pack(anchor="w")
    make_wrapped_label(
        ldap_outer,
        "Enables Domain Login (shown as an option on the Admin login screen) using a real "
        "LDAP/AD bind. Group membership resolves to one of five roles below; only "
        "Administrators unlocks Admin mode in this app today -- Technicians/Radiologists/"
        "Support/Guest all land in the normal User mode, same as any local user. The local "
        "Admin PIN always keeps working as a fallback regardless of this setting.",
        780, font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(4, 14))

    if not LDAP3_AVAILABLE:
        ctk.CTkLabel(
            ldap_outer,
            text="'ldap3' is not installed -- LDAP/AD login is unavailable until you run: pip install ldap3",
            image=get_icon("triangle-alert", size=14, color=THEME_WARNING), compound="left",
            font=get_font("small"), text_color=THEME_WARNING,
        ).pack(anchor="w", pady=(0, 10))

    _ldap_cfg = load_ldap_config()

    ldap_enabled_var = ctk.BooleanVar(value=_ldap_cfg.get("enabled", False))
    ctk.CTkCheckBox(ldap_outer, text="Enable LDAP / AD Authentication", variable=ldap_enabled_var,
                    font=get_font("caption", "bold")).pack(anchor="w", pady=(0, 10))

    ldap_form = ctk.CTkFrame(ldap_outer)
    ldap_form.pack(fill="x", pady=(0, 10))

    def _ldap_row(parent, row, label, width=260):
        ctk.CTkLabel(parent, text=label, font=get_font("body")).grid(
            row=row, column=0, padx=(12, 8), pady=6, sticky="w")
        entry = ctk.CTkEntry(parent, width=width)
        entry.grid(row=row, column=1, padx=8, pady=6, sticky="w")
        return entry

    ldap_server_entry = _ldap_row(ldap_form, 0, "Server URI:", 320)
    ldap_server_entry.insert(0, _ldap_cfg.get("server_uri", ""))
    ctk.CTkLabel(ldap_form, text="e.g. ldap://dc1.company.local:389 or ldaps://dc1.company.local:636",
                font=get_font("micro"), text_color=THEME_TEXT_MUTED).grid(row=0, column=2, padx=8, sticky="w")

    ldap_ssl_var = ctk.BooleanVar(value=_ldap_cfg.get("use_ssl", False))
    ctk.CTkCheckBox(ldap_form, text="Use SSL", variable=ldap_ssl_var).grid(row=1, column=1, padx=8, pady=6, sticky="w")

    ldap_domain_entry = _ldap_row(ldap_form, 2, "AD Domain (UPN bind):")
    ldap_domain_entry.insert(0, _ldap_cfg.get("domain", ""))
    ctk.CTkLabel(ldap_form, text="e.g. company.local -- builds username@domain for AD",
                font=get_font("micro"), text_color=THEME_TEXT_MUTED).grid(row=2, column=2, padx=8, sticky="w")

    ldap_bind_template_entry = _ldap_row(ldap_form, 3, "OR Bind DN Template:", 320)
    ldap_bind_template_entry.insert(0, _ldap_cfg.get("user_bind_dn_template", ""))
    ctk.CTkLabel(ldap_form, text="Generic/OpenLDAP style, e.g. cn={username},ou=Users,dc=company,dc=local",
                font=get_font("micro"), text_color=THEME_TEXT_MUTED).grid(row=3, column=2, padx=8, sticky="w")

    ldap_bind_dn_entry = _ldap_row(ldap_form, 4, "Service Account (bind DN):", 320)
    ldap_bind_dn_entry.insert(0, _ldap_cfg.get("bind_dn", ""))

    ldap_bind_pw_entry = ctk.CTkEntry(ldap_form, width=260, show="•")
    ldap_bind_pw_entry.insert(0, _ldap_cfg.get("bind_password", ""))
    ctk.CTkLabel(ldap_form, text="Service Account Password:", font=get_font("body")).grid(
        row=5, column=0, padx=(12, 8), pady=6, sticky="w")
    ldap_bind_pw_entry.grid(row=5, column=1, padx=8, pady=6, sticky="w")

    ldap_user_base_entry = _ldap_row(ldap_form, 6, "User Search Base:", 320)
    ldap_user_base_entry.insert(0, _ldap_cfg.get("user_search_base", ""))

    ldap_user_filter_entry = _ldap_row(ldap_form, 7, "User Search Filter:", 320)
    ldap_user_filter_entry.insert(0, _ldap_cfg.get("user_search_filter", "(sAMAccountName={username})"))

    ldap_group_base_entry = _ldap_row(ldap_form, 8, "Group Search Base (optional):", 320)
    ldap_group_base_entry.insert(0, _ldap_cfg.get("group_search_base", ""))

    ctk.CTkLabel(ldap_outer, text="Group → Role Mapping", font=get_font("caption", "bold")).pack(
        anchor="w", pady=(6, 4))
    ctk.CTkLabel(
        ldap_outer,
        text="Enter the LDAP/AD group CN (not the full DN) that should map to each role. Leave blank to skip a role.",
        font=get_font("small"), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(0, 6))

    ldap_group_map_frame = ctk.CTkFrame(ldap_outer)
    ldap_group_map_frame.pack(fill="x", pady=(0, 10))
    ldap_group_map_entries = {}
    existing_mappings = _ldap_cfg.get("group_mappings", {})
    # Invert existing cn->role mapping to role->cn for pre-filling the form
    role_to_cn = {role: cn for cn, role in existing_mappings.items()}
    for i, role in enumerate(LDAP_ROLE_NAMES):
        ctk.CTkLabel(ldap_group_map_frame, text=f"{role}:", font=get_font("body")).grid(
            row=i, column=0, padx=(12, 8), pady=6, sticky="w")
        entry = ctk.CTkEntry(ldap_group_map_frame, width=220,
                             placeholder_text="e.g. RAPPS-Admins" if role == "Administrators" else "")
        entry.insert(0, role_to_cn.get(role, ""))
        entry.grid(row=i, column=1, padx=8, pady=6, sticky="w")
        ldap_group_map_entries[role] = entry

    ldap_run_row = ctk.CTkFrame(ldap_outer, fg_color="transparent")
    ldap_run_row.pack(fill="x", pady=(4, 4))
    ldap_save_btn = ctk.CTkButton(ldap_run_row, text="Save LDAP Settings", width=180)
    ldap_save_btn.pack(side="left")
    ldap_import_btn = ctk.CTkButton(ldap_run_row, text="Import Users Now", width=170)
    ldap_import_btn.pack(side="left", padx=8)
    ldap_status_lbl = ctk.CTkLabel(ldap_outer, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    ldap_status_lbl.pack(anchor="w", pady=(6, 12))

    ctk.CTkLabel(ldap_outer, text="Imported User Roster", font=get_font("caption", "bold")).pack(
        anchor="w", pady=(4, 4))
    ldap_roster_cols = ("username", "display_name", "email", "role", "last_login", "access_status")
    ldap_roster_frame = ctk.CTkFrame(ldap_outer, fg_color=THEME_SURFACE)
    ldap_roster_frame.pack(fill="both", expand=True, pady=(0, 8))
    ldap_roster_tree = ttk.Treeview(ldap_roster_frame, columns=ldap_roster_cols, show="headings", height=8)
    for col, label, w in (("username", "Username", 130), ("display_name", "Display Name", 180),
                          ("email", "Email", 220), ("role", "Role", 130), ("last_login", "Last Login", 160),
                          ("access_status", "Access", 90)):
        ldap_roster_tree.heading(col, text=label)
        ldap_roster_tree.column(col, width=w, anchor="w")
    ldap_roster_tree.tag_configure("access_revoked", foreground=THEME_DANGER)
    ldap_roster_tree.pack(fill="both", expand=True)
    ldap_roster_tree.bind("<Button-3>", _build_ldap_roster_context_menu)
    ldap_roster_tree.bind("<Button-2>", _build_ldap_roster_context_menu)  # macOS

    ldap_save_btn.configure(command=do_save_ldap_config)
    ldap_import_btn.configure(command=do_ldap_import_users_now)

    # ---- Query/Retrieve tab (unchanged from the original single-mode app) ----
    tab_qr = tabview.add("Query/Retrieve")


    qr_top = ctk.CTkFrame(tab_qr, fg_color="transparent")
    qr_top.pack(fill="x", padx=10, pady=10)

    qr_cfg_frame = ctk.CTkFrame(qr_top)
    qr_cfg_frame.pack(side="left", padx=(0, 20))

    for i, lbl in enumerate(["Remote AE", "Remote IP", "Remote Port"]):
        ctk.CTkLabel(qr_cfg_frame, text=lbl).grid(row=i, column=0, padx=8, pady=6, sticky="w")
    qr_ae_entry = ctk.CTkEntry(qr_cfg_frame, width=150)
    qr_ae_entry.grid(row=0, column=1, padx=8, pady=6)
    qr_ip_entry = ctk.CTkEntry(qr_cfg_frame, width=150)
    qr_ip_entry.grid(row=1, column=1, padx=8, pady=6)
    qr_port_entry = ctk.CTkEntry(qr_cfg_frame, width=150)
    qr_port_entry.grid(row=2, column=1, padx=8, pady=6)

    qr_filter_frame = ctk.CTkFrame(qr_top)
    qr_filter_frame.pack(side="left", padx=(0, 20))

    qr_filter_entries = {}
    for i, (lbl, key) in enumerate([("Patient ID", "pid"), ("Patient Name", "pname"),
                                      ("Study Date (YYYYMMDD)", "date"), ("Modality", "mod")]):
        ctk.CTkLabel(qr_filter_frame, text=lbl).grid(row=i, column=0, padx=8, pady=4, sticky="w")
        e = ctk.CTkEntry(qr_filter_frame, width=160)
        e.grid(row=i, column=1, padx=8, pady=4)
        qr_filter_entries[key] = e

    qr_btn_col = ctk.CTkFrame(qr_top, fg_color="transparent")
    qr_btn_col.pack(side="left")

    qr_find_btn = ctk.CTkButton(qr_btn_col, text="C-FIND Query", width=190,
                                  fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER)
    qr_find_btn.pack(pady=6)

    qr_retrieve_btn = ctk.CTkButton(qr_btn_col, text="C-MOVE Retrieve Selected", width=215,
                                     fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    qr_retrieve_btn.pack(pady=6)

    qr_status_label = ctk.CTkLabel(qr_btn_col, text="", font=get_font("small"),
                                    text_color=THEME_TEXT_MUTED, wraplength=190)
    qr_status_label.pack(pady=6)

    # Q/R results treeview
    qr_results_cols = ("patient_id", "patient_name", "study_date", "modality",
                        "instances", "description", "study_uid")
    qr_results_headings = {
        "patient_id": "Patient ID", "patient_name": "Patient Name",
        "study_date": "Study Date", "modality": "Modality",
        "instances": "#Images", "description": "Description", "study_uid": "Study UID",
    }
    qr_results_widths = {
        "patient_id": 120, "patient_name": 160, "study_date": 100,
        "modality": 70, "instances": 65, "description": 200, "study_uid": 280,
    }

    qr_tree_frame = ctk.CTkFrame(tab_qr, fg_color=THEME_SURFACE)
    qr_tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

    qr_vsb = ttk.Scrollbar(qr_tree_frame, orient="vertical")
    qr_hsb = ttk.Scrollbar(qr_tree_frame, orient="horizontal")
    qr_tree = ttk.Treeview(
        qr_tree_frame, columns=qr_results_cols, show="headings",
        yscrollcommand=qr_vsb.set, xscrollcommand=qr_hsb.set, selectmode="extended",
    )
    qr_vsb.config(command=qr_tree.yview)
    qr_hsb.config(command=qr_tree.xview)
    for col in qr_results_cols:
        qr_tree.heading(col, text=qr_results_headings[col])
        qr_tree.column(col, width=qr_results_widths[col], minwidth=40, anchor="w")
    qr_vsb.pack(side="right", fill="y")
    qr_hsb.pack(side="bottom", fill="x")
    qr_tree.pack(fill="both", expand=True)

    # ---- Destinations tab (multi-destination manager, unchanged) ----
    tab_destinations = tabview.add("Destinations")

    dest_outer = ctk.CTkFrame(tab_destinations, fg_color="transparent")
    dest_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    dest_list_frame = ctk.CTkFrame(dest_outer)
    dest_list_frame.pack(side="left", fill="y", padx=(0, 10))

    ctk.CTkLabel(dest_list_frame, text="Saved Destinations",
                 font=get_font("caption", "bold")).pack(pady=(8, 4))

    dest_listbox_var = ctk.StringVar(value=[])
    dest_listbox = ctk.CTkTextbox(dest_list_frame, width=220, height=350, state="disabled")
    dest_listbox.pack(padx=8, pady=4)

    dest_select_var = ctk.StringVar()
    dest_names_var = ctk.StringVar(value=[])
    dest_optmenu = ctk.CTkOptionMenu(dest_list_frame, variable=dest_select_var,
                                      values=["(none)"], width=200)
    dest_optmenu.pack(pady=4)

    dest_form_frame = ctk.CTkScrollableFrame(dest_outer)
    dest_form_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

    ctk.CTkLabel(dest_form_frame, text="Destination Profile",
                 font=get_font("caption", "bold")).pack(pady=(8, 4))

    dest_fields = {}
    dest_field_default_border = {}
    for lbl, key, placeholder in [
        ("Profile Name", "name", "e.g. Main PACS"),
        ("Remote AE Title (Called)", "ae", "e.g. ORTHANC"),
        ("Calling AE Title (our identity)", "calling_ae", "e.g. RAPPS  (blank = RAPPS_PUSH)"),
        ("Remote IP", "ip", "e.g. 192.168.1.100"),
        ("Remote Port", "port", "e.g. 4242"),
    ]:
        ctk.CTkLabel(dest_form_frame, text=lbl).pack(anchor="w", padx=10)
        e = ctk.CTkEntry(dest_form_frame, width=220, placeholder_text=placeholder)
        e.pack(padx=10, pady=4)
        dest_fields[key] = e
        dest_field_default_border[key] = e.cget("border_color")
        # 9.3 -- live inline validation as the user types: AE Title/Port
        # fields get validated for real (validate_ae_title/validate_port),
        # other fields just clear the invalid highlight as soon as editing
        # starts, same as before.
        e.bind("<KeyRelease>", lambda _evt, k=key: _live_validate_dest_field(k))

    dest_default_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(dest_form_frame, text="Set as default destination",
                    variable=dest_default_var).pack(anchor="w", padx=10, pady=4)

    # ---- Document Transfer (Radiology_Report.docx / Patient_History.txt) ----
    # Per-destination opt-in on the pusher side. The receiver only actually
    # listens for these connections if the GLOBAL "Document Transfer"
    # setting on the Settings tab is also on (see start_receiver_server()).
    ctk.CTkFrame(dest_form_frame, fg_color=THEME_HEADING_BG, height=1).pack(fill="x", padx=10, pady=(6, 6))
    ctk.CTkLabel(dest_form_frame, text="Document Transfer (Report / History)",
                 font=get_font("caption", "bold")).pack(anchor="w", padx=10, pady=(0, 4))

    dest_doc_transfer_enabled_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(dest_form_frame, text="Send Report/History to this destination",
                    variable=dest_doc_transfer_enabled_var).pack(anchor="w", padx=10, pady=2)

    ctk.CTkLabel(dest_form_frame, text="Document Transfer Port").pack(anchor="w", padx=10)
    e = ctk.CTkEntry(dest_form_frame, width=220, placeholder_text="e.g. 11244")
    e.pack(padx=10, pady=4)
    dest_fields["doc_transfer_port"] = e
    dest_field_default_border["doc_transfer_port"] = e.cget("border_color")
    e.bind("<KeyRelease>", lambda _evt, k="doc_transfer_port": _live_validate_dest_field(k))

    dest_doc_transfer_use_dicom_host_var = ctk.BooleanVar(value=True)
    ctk.CTkCheckBox(dest_form_frame, text="Use this destination's IP (above) for documents too",
                    variable=dest_doc_transfer_use_dicom_host_var).pack(anchor="w", padx=10, pady=2)

    ctk.CTkLabel(dest_form_frame, text="Document Transfer IP (if not using DICOM host)").pack(anchor="w", padx=10)
    e = ctk.CTkEntry(dest_form_frame, width=220, placeholder_text="e.g. 192.168.1.100")
    e.pack(padx=10, pady=4)
    dest_fields["doc_transfer_ip"] = e
    dest_field_default_border["doc_transfer_ip"] = e.cget("border_color")

    auth_key_label_row = ctk.CTkFrame(dest_form_frame, fg_color="transparent")
    auth_key_label_row.pack(fill="x", padx=10)
    ctk.CTkLabel(auth_key_label_row, text="Document Transfer Auth Key").pack(side="left")
    ctk.CTkLabel(auth_key_label_row, text="  must match the receiver's key exactly (Settings tab)",
                image=get_icon("shield", size=12, color=THEME_TEXT_MUTED), compound="left",
                text_color=THEME_TEXT_MUTED, font=get_font("small")).pack(side="left", padx=(4, 0))
    e = ctk.CTkEntry(dest_form_frame, width=220, placeholder_text="optional -- leave blank if the receiver has none set")
    e.pack(padx=10, pady=4)
    dest_fields["doc_transfer_auth_key"] = e
    dest_field_default_border["doc_transfer_auth_key"] = e.cget("border_color")

    # Part 3 trust indicator. Authenticated/Unauthenticated reflects
    # this destination's CONFIGURED key (a real key means the receiver
    # will actually check it -- see Settings tab); "Not Secure" is
    # always shown as-is rather than aspirationally, since the doc-
    # transfer socket has no TLS yet (see the TLS plug-in point comment
    # in push_patient_documents) -- claiming "Secure" would be false.
    dest_trust_row = ctk.CTkFrame(dest_form_frame, fg_color="transparent")
    dest_trust_row.pack(padx=10, pady=(0, 6), anchor="w")
    dest_trust_auth_badge = ctk.CTkLabel(dest_trust_row, text="", font=get_font("small"))
    dest_trust_auth_badge.pack(side="left")
    ctk.CTkLabel(dest_trust_row, text="  Not Secure (plaintext)",
                image=get_icon("lock-open", size=13, color=THEME_DANGER),
                compound="left", text_color=THEME_DANGER, font=get_font("small")).pack(side="left", padx=(10, 0))
    e.bind("<KeyRelease>", lambda _e: update_dest_trust_badge())
    update_dest_trust_badge()

    dest_btn_frame = ctk.CTkFrame(dest_form_frame, fg_color="transparent")
    dest_btn_frame.pack(pady=8)

    dest_add_btn = ctk.CTkButton(dest_btn_frame, text="Add / Update", width=150,
                                  image=get_icon("plus", size=15, color="#ffffff"), compound="left",
                                  fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    dest_add_btn.pack(side="left", padx=4)
    dest_del_btn = ctk.CTkButton(dest_btn_frame, text="Delete Selected", width=150,
                                  image=get_icon("trash-2", size=15, color="#ffffff"), compound="left",
                                  fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER)
    dest_del_btn.pack(side="left", padx=4)

    dest_echo_btn = ctk.CTkButton(dest_form_frame, text="C-ECHO Test", width=220,
                                   image=get_icon("activity", size=15, color=THEME_TEXT), compound="left")
    dest_echo_btn.pack(pady=4)

    dest_status_lbl = ctk.CTkLabel(dest_form_frame, text="", font=get_font("small"),
                                    text_color=THEME_TEXT_MUTED, wraplength=220)
    dest_status_lbl.pack(pady=4)

    # ---- Document Transfer Activity tab (Part 3 of the doc-transfer
    # audit). Reads the per-file JSONL records already written by
    # push_patient_documents() (event_type=DOC-TRANSFER-FILE) and
    # _doc_transfer_receive_one_file() (event_type=DOC-TRANSFER) -- no
    # new persistent state, same pattern as every other stats tab in
    # this app (Admin Dashboard, Performance, PACS Health). ----
    tab_doc_transfer = tabview.add("Doc Transfer")
    dt_outer = ctk.CTkFrame(tab_doc_transfer, fg_color="transparent")
    dt_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    dt_top = ctk.CTkFrame(dt_outer, fg_color="transparent")
    dt_top.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(dt_top, text="Document Transfer Activity", font=get_font("section_lg", "bold")).pack(side="left")
    dt_refresh_btn = ctk.CTkButton(dt_top, text="Refresh", width=110,
                                   image=get_icon("refresh-cw", size=14, color=THEME_TEXT), compound="left")
    dt_refresh_btn.pack(side="right", padx=4)

    dt_stats_row = make_card_row(dt_outer, pady=(0, 10))

    def _dt_stat_card(title):
        card = ctk.CTkFrame(dt_stats_row, corner_radius=12, fg_color=THEME_HEADING_BG)
        add_card_to_row(dt_stats_row, card, padx=4, pady=0)
        ctk.CTkLabel(card, text=title, font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(pady=(8, 0))
        val_lbl = ctk.CTkLabel(card, text="—", font=get_font("kpi", "bold"), text_color=THEME_TEXT)
        val_lbl.pack(pady=(0, 8), padx=6)
        return val_lbl

    dt_docs_today_lbl = _dt_stat_card("Documents Today")
    dt_success_lbl = _dt_stat_card("Success %")
    dt_avg_speed_lbl = _dt_stat_card("Avg Speed")
    dt_avg_duration_lbl = _dt_stat_card("Avg Duration")
    dt_failed_lbl = _dt_stat_card("Failed")
    dt_pending_lbl = _dt_stat_card("Pending")
    dt_retries_lbl = _dt_stat_card("Retries")
    dt_largest_lbl = _dt_stat_card("Largest Transfer")

    # Active Transfers: live progress for whatever push_patient_documents()
    # calls happen to be running right now (manual resend, offline-queue
    # retry, routing-rule auto-push -- all share the same tracking dict).
    # Only shown when something is actually in flight. Data is updated by
    # the background send loop in real time; this panel just polls it on
    # the same 2s periodic_refresh() cadence every other tab already uses.
    dt_active_frame = ctk.CTkFrame(dt_outer, fg_color=THEME_HEADING_BG, corner_radius=10)
    dt_active_rows_container = ctk.CTkFrame(dt_active_frame, fg_color="transparent")
    dt_active_rows_container.pack(fill="x", padx=8, pady=6)
    _dt_active_row_widgets = {}

    def _dt_refresh_active_transfers():
        _dt_progress_sweep_finished()
        with doc_transfer_progress_lock:
            snapshot = {tid: dict(e) for tid, e in doc_transfer_active_transfers.items()}

        for tid in list(_dt_active_row_widgets.keys()):
            if tid not in snapshot:
                _dt_active_row_widgets[tid]["frame"].destroy()
                del _dt_active_row_widgets[tid]

        if not snapshot:
            dt_active_frame.pack_forget()
            return
        if not dt_active_frame.winfo_ismapped():
            dt_active_frame.pack(fill="x", pady=(0, 10), before=dt_view_toggle)

        for tid, entry in snapshot.items():
            if tid not in _dt_active_row_widgets:
                row = ctk.CTkFrame(dt_active_rows_container, fg_color="transparent")
                row.pack(fill="x", pady=4)
                label = ctk.CTkLabel(row, text="", font=get_font("small"), anchor="w")
                label.pack(anchor="w")
                bar_row = ctk.CTkFrame(row, fg_color="transparent")
                bar_row.pack(fill="x", pady=(2, 0))
                bar = ctk.CTkProgressBar(bar_row, width=240)
                bar.set(0)
                bar.pack(side="left")
                detail_lbl = ctk.CTkLabel(bar_row, text="", font=get_font("micro"), text_color=THEME_TEXT_MUTED)
                detail_lbl.pack(side="left", padx=(8, 0))
                cancel_btn = ctk.CTkButton(bar_row, text="Cancel", width=64, height=22, corner_radius=6,
                                           fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER,
                                           command=lambda t=tid: cancel_doc_transfer(t))
                cancel_btn.pack(side="right", padx=(8, 0))
                _dt_active_row_widgets[tid] = {"frame": row, "label": label, "bar": bar,
                                               "detail": detail_lbl, "cancel": cancel_btn}

            w = _dt_active_row_widgets[tid]
            bytes_total = entry.get("bytes_total", 1) or 1
            bytes_sent = entry.get("bytes_sent", 0)
            frac = max(0.0, min(1.0, bytes_sent / bytes_total))
            speed = entry.get("speed_mbps") or 0.0
            remaining_mb = max(0, bytes_total - bytes_sent) / (1024 * 1024)
            eta_txt = f"{remaining_mb / speed:.0f}s left" if speed > 0.01 else "—"
            status = entry.get("status", "running")

            w["label"].configure(
                text=f"{entry.get('pid', '')} → {entry.get('destination_name', '')}  ·  "
                     f"{entry.get('current_filename') or '…'}  "
                     f"({entry.get('files_done', 0)}/{entry.get('files_total', 0)} files)"
                     + {"done": "  ·  Done", "failed": "  ·  Failed",
                        "cancelled": "  ·  Cancelled"}.get(status, ""))
            w["bar"].set(frac)
            w["detail"].configure(
                text=f"{frac * 100:.0f}%  ·  {_human_size(bytes_sent)}/{_human_size(bytes_total)}"
                     f"  ·  {speed:.1f} MB/s  ·  {eta_txt}")
            if status != "running":
                w["cancel"].configure(state="disabled", text="—")

    dt_view_var = ctk.StringVar(value="Sent Activity")
    dt_view_toggle = ctk.CTkSegmentedButton(dt_outer, variable=dt_view_var,
                                            values=["Sent Activity", "Received Activity"])
    dt_view_toggle.pack(anchor="w", pady=(0, 8))

    # --- Sent Activity table ---
    dt_sent_cols = ("timestamp", "patient_id", "patient_name", "destination", "filename",
                    "size", "duration", "speed", "status", "retries", "type", "error")
    dt_sent_headings = {
        "timestamp": "Timestamp", "patient_id": "Patient ID", "patient_name": "Patient Name",
        "destination": "Destination", "filename": "Filename", "size": "Size", "duration": "Duration",
        "speed": "Speed", "status": "Status", "retries": "Retries", "type": "Type", "error": "Error",
    }
    dt_sent_widths = {
        "timestamp": 135, "patient_id": 100, "patient_name": 130, "destination": 120,
        "filename": 170, "size": 70, "duration": 70, "speed": 85, "status": 70,
        "retries": 60, "type": 85, "error": 220,
    }
    dt_sent_frame = ctk.CTkFrame(dt_outer, fg_color=THEME_SURFACE)
    dt_sent_tree = ttk.Treeview(dt_sent_frame, columns=dt_sent_cols, show="headings", height=13)
    for col in dt_sent_cols:
        dt_sent_tree.heading(col, text=dt_sent_headings[col],
                             command=lambda c=col: sort_tree(dt_sent_tree, c, False))
        dt_sent_tree.column(col, width=dt_sent_widths[col], anchor="w")
    dt_sent_tree.pack(fill="both", expand=True)
    dt_sent_tree.tag_configure("status_ok", foreground=THEME_SUCCESS)
    dt_sent_tree.tag_configure("status_failed", foreground=THEME_DANGER)

    dt_sent_action_row = ctk.CTkFrame(dt_outer, fg_color="transparent")
    dt_sent_resend_btn = ctk.CTkButton(dt_sent_action_row, text="Resend Selected", width=160,
                                       image=get_icon("send", size=14, color="#ffffff"), compound="left")
    dt_sent_resend_btn.pack(side="left")
    dt_sent_status_lbl = ctk.CTkLabel(dt_sent_action_row, text="", font=get_font("small"),
                                      text_color=THEME_TEXT_MUTED)
    dt_sent_status_lbl.pack(side="left", padx=10)

    # --- Received Activity table ---
    dt_recv_cols = ("timestamp", "patient_id", "patient_name", "sender", "filename",
                    "destination_folder", "overwrite", "backup", "trust", "status")
    dt_recv_headings = {
        "timestamp": "Timestamp", "patient_id": "Patient", "patient_name": "Patient Name",
        "sender": "Sender", "filename": "File", "destination_folder": "Destination Folder",
        "overwrite": "Overwrite", "backup": "Backup", "trust": "Trust", "status": "Status",
    }
    dt_recv_widths = {
        "timestamp": 135, "patient_id": 100, "patient_name": 130, "sender": 110,
        "filename": 170, "destination_folder": 240, "overwrite": 75, "backup": 70,
        "trust": 110, "status": 70,
    }
    dt_recv_frame = ctk.CTkFrame(dt_outer, fg_color=THEME_SURFACE)
    dt_recv_tree = ttk.Treeview(dt_recv_frame, columns=dt_recv_cols, show="headings", height=13)
    for col in dt_recv_cols:
        dt_recv_tree.heading(col, text=dt_recv_headings[col],
                             command=lambda c=col: sort_tree(dt_recv_tree, c, False))
        dt_recv_tree.column(col, width=dt_recv_widths[col], anchor="w")
    dt_recv_tree.pack(fill="both", expand=True)
    dt_recv_tree.tag_configure("status_ok", foreground=THEME_SUCCESS)
    dt_recv_tree.tag_configure("status_failed", foreground=THEME_DANGER)

    dt_recv_action_row = ctk.CTkFrame(dt_outer, fg_color="transparent")
    dt_recv_open_file_btn = ctk.CTkButton(dt_recv_action_row, text="Open File", width=130,
                                          image=get_icon("file-text", size=14, color=THEME_TEXT), compound="left",
                                          fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
    dt_recv_open_file_btn.pack(side="left")
    dt_recv_open_folder_btn = ctk.CTkButton(dt_recv_action_row, text="Open Folder", width=130,
                                            image=get_icon("folder-open", size=14, color=THEME_TEXT), compound="left",
                                            fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
    dt_recv_open_folder_btn.pack(side="left", padx=(6, 0))

    _dt_sent_row_records = {}
    _dt_recv_row_records = {}

    def _dt_switch_view(*_a):
        if dt_view_var.get() == "Sent Activity":
            dt_recv_frame.pack_forget()
            dt_recv_action_row.pack_forget()
            dt_sent_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))
            dt_sent_action_row.pack(fill="x")
        else:
            dt_sent_frame.pack_forget()
            dt_sent_action_row.pack_forget()
            dt_recv_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))
            dt_recv_action_row.pack(fill="x")

    dt_view_toggle.configure(command=_dt_switch_view)

    global refresh_doc_transfer_tab
    def refresh_doc_transfer_tab():
        try:
            _dt_refresh_active_transfers()
        except Exception:
            log_exception("Failed to refresh Active Transfers panel")

        now = datetime.datetime.now()
        start = now - datetime.timedelta(days=1)
        try:
            stats = compute_doc_transfer_activity_stats(start, now)
        except Exception:
            log_exception("Failed to refresh Doc Transfer tab")
            return

        dt_docs_today_lbl.configure(text=str(stats["documents_today"]))
        dt_success_lbl.configure(text=f"{stats['success_rate_pct']:.0f}%")
        dt_avg_speed_lbl.configure(text=f"{stats['avg_speed_mbps']:.1f} MB/s" if stats["avg_speed_mbps"] else "—")
        dt_avg_duration_lbl.configure(text=f"{stats['avg_duration_sec']:.1f}s" if stats["avg_duration_sec"] else "—")
        dt_failed_lbl.configure(text=str(stats["failed"]))
        dt_pending_lbl.configure(text=str(stats["pending"]))
        dt_retries_lbl.configure(text=str(stats["retries"]))
        dt_largest_lbl.configure(text=_human_size(stats["largest_transfer_bytes"]))

        dt_sent_tree.delete(*dt_sent_tree.get_children())
        _dt_sent_row_records.clear()
        sent_sorted = sorted(stats["sent_records"], key=lambda r: r.get("timestamp", ""), reverse=True)
        for i, r in enumerate(sent_sorted[:300]):
            row_id = f"s{i}"
            status = r.get("final_status", "")
            tag = "status_ok" if status == "OK" else "status_failed"
            speed = r.get("transfer_speed_mbps")
            duration = r.get("duration_sec")
            values = (
                str(r.get("timestamp", ""))[:19], r.get("patient_id", ""), r.get("patient_name", "") or "—",
                r.get("destination_name", ""), r.get("filename", ""), _human_size(r.get("file_size", 0)),
                f"{duration:.1f}s" if isinstance(duration, (int, float)) else "—",
                f"{speed:.1f} MB/s" if isinstance(speed, (int, float)) else "—",
                status, r.get("retry_count", 0), r.get("transfer_type", ""), r.get("error_message", "") or "",
            )
            dt_sent_tree.insert("", "end", iid=row_id, values=values, tags=(tag,))
            _dt_sent_row_records[row_id] = r

        dt_recv_tree.delete(*dt_recv_tree.get_children())
        _dt_recv_row_records.clear()
        recv_sorted = sorted(stats["received_records"], key=lambda r: r.get("timestamp", ""), reverse=True)
        for i, r in enumerate(recv_sorted[:300]):
            row_id = f"r{i}"
            status = r.get("result", "")
            tag = "status_ok" if status == "SUCCESS" else "status_failed"
            trust_text = "Authenticated" if r.get("authenticated") else "Unauthenticated"
            values = (
                str(r.get("timestamp", ""))[:19], r.get("patient_id", ""), r.get("patient_name", "") or "—",
                r.get("source_ip", "") or "—", r.get("filename", ""), r.get("save_location", "") or "",
                "Yes" if r.get("overwrite_occurred") else "No",
                "Yes" if r.get("backup_created") else "No",
                trust_text, status,
            )
            dt_recv_tree.insert("", "end", iid=row_id, values=values, tags=(tag,))
            _dt_recv_row_records[row_id] = r

    def do_resend_selected_document():
        sel = dt_sent_tree.selection()
        if not sel:
            dt_sent_status_lbl.configure(text="Select a row first.", text_color=THEME_TEXT_MUTED)
            return
        record = _dt_sent_row_records.get(sel[0])
        if not record:
            return
        pid = record.get("patient_id", "")
        dest_name = record.get("destination_name", "")
        dest_cfg = get_destination_by_name(dest_name)
        if not pid or not dest_cfg:
            dt_sent_status_lbl.configure(text=f"Can't resend: destination \"{dest_name}\" no longer exists.",
                                         text_color=THEME_DANGER)
            return
        dt_sent_status_lbl.configure(text=f"Resending to {dest_name}…", text_color=THEME_TEXT_MUTED)

        def run():
            try:
                push_patient_documents(pid, dest_cfg)
                app.after(0, lambda: (dt_sent_status_lbl.configure(text="Resend attempted -- refreshing…",
                                                                   text_color=THEME_TEXT_MUTED),
                                      refresh_doc_transfer_tab()))
            except Exception as e:
                log_exception("Resend failed")
                err_msg = str(e)
                app.after(0, lambda: dt_sent_status_lbl.configure(text=f"Resend failed: {err_msg}", text_color=THEME_DANGER))

        threading.Thread(target=run, daemon=True).start()

    def do_open_received_file():
        sel = dt_recv_tree.selection()
        if not sel:
            return
        record = _dt_recv_row_records.get(sel[0])
        path = record.get("save_location") if record else None
        if not path or not os.path.isfile(path):
            modern_showerror("File Not Found", "This file no longer exists at its recorded location.")
            return
        do_open_attachment(path)

    def do_open_received_folder():
        sel = dt_recv_tree.selection()
        if not sel:
            return
        record = _dt_recv_row_records.get(sel[0])
        path = record.get("save_location") if record else None
        if not path:
            return
        ok, err = open_containing_folder(path)
        if not ok:
            modern_showerror("Could Not Open Folder", err)

    dt_refresh_btn.configure(command=refresh_doc_transfer_tab)
    dt_sent_resend_btn.configure(command=do_resend_selected_document)
    dt_recv_open_file_btn.configure(command=do_open_received_file)
    dt_recv_open_folder_btn.configure(command=do_open_received_folder)
    _dt_switch_view()

    # ---- Routing Rules tab (unchanged) ----
    tab_routing = tabview.add("Routing Rules")

    routing_outer = ctk.CTkFrame(tab_routing, fg_color="transparent")
    routing_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    ctk.CTkLabel(routing_outer,
                 text="Rules are evaluated top-down. The first matching rule's destination is used.\n"
                      "Leave a field blank to match any value (wildcard).",
                 font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 8))

    routing_list_frame = ctk.CTkFrame(routing_outer, fg_color=THEME_SURFACE)
    routing_list_frame.pack(fill="both", expand=True)

    routing_cols = ("modality", "institution", "source_ae", "destination")
    routing_headings = {"modality": "Modality", "institution": "Institution",
                        "source_ae": "Source AE", "destination": "→ Destination"}
    routing_widths = {"modality": 120, "institution": 200, "source_ae": 150, "destination": 200}

    routing_vsb = ttk.Scrollbar(routing_list_frame, orient="vertical")
    routing_tree = ttk.Treeview(
        routing_list_frame, columns=routing_cols, show="headings",
        yscrollcommand=routing_vsb.set, height=10,
    )
    routing_vsb.config(command=routing_tree.yview)
    for col in routing_cols:
        routing_tree.heading(col, text=routing_headings[col])
        routing_tree.column(col, width=routing_widths[col], minwidth=40, anchor="w")
    routing_vsb.pack(side="right", fill="y")
    routing_tree.pack(fill="both", expand=True)

    routing_form_frame = ctk.CTkFrame(routing_outer, fg_color="transparent")
    routing_form_frame.pack(fill="x", pady=8)

    routing_fields = {}
    routing_dest_var = ctk.StringVar(value="(none)")

    def _mk_label(text):
        return lambda row: ctk.CTkLabel(row, text=text)

    def _mk_entry(key):
        def factory(row):
            e = ctk.CTkEntry(row, width=130)
            routing_fields[key] = e
            return e
        return factory

    def _mk_dest_menu(row):
        return ctk.CTkOptionMenu(row, variable=routing_dest_var, values=["(none)"], width=160)

    def _mk_btn(**kwargs):
        return lambda row: ctk.CTkButton(row, **kwargs)

    # One unified adaptive row: labels stay their natural size, but the
    # entries, destination dropdown, and buttons all shrink together in
    # equal ratios once the row runs out of room -- not just the buttons.
    routing_row, routing_row_widgets = make_equal_share_row(
        routing_form_frame,
        [
            {"factory": _mk_label("Modality (blank=any)"), "shrink": False},
            {"factory": _mk_entry("modality"), "shrink": True},
            {"factory": _mk_label("Institution (blank=any)"), "shrink": False},
            {"factory": _mk_entry("institution"), "shrink": True},
            {"factory": _mk_label("Source AE (blank=any)"), "shrink": False},
            {"factory": _mk_entry("source_ae"), "shrink": True},
            {"factory": _mk_label("→ Destination"), "shrink": False},
            {"factory": _mk_dest_menu, "shrink": True},
            {"factory": _mk_btn(text="Add Rule", fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER), "shrink": True},
            {"factory": _mk_btn(text="Delete", fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER), "shrink": True},
            {"factory": _mk_btn(text="▲"), "shrink": True},
            {"factory": _mk_btn(text="▼"), "shrink": True},
            {"factory": _mk_btn(text="Test Rule", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER), "shrink": True},
        ],
    )
    (routing_add_btn, routing_del_btn, routing_up_btn, routing_down_btn,
     routing_test_btn) = routing_row_widgets[-5:]
    routing_dest_menu = routing_row_widgets[7]  # the option menu, by its position in the specs list above

    # ---- SOP Classes tab (in-GUI editor, unchanged) ----
    tab_sop = tabview.add("SOP Classes")

    sop_outer = ctk.CTkFrame(tab_sop, fg_color="transparent")
    sop_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    ctk.CTkLabel(sop_outer,
                 text="Edit SOP Classes and Transfer Syntaxes. Changes are written to sopclass.ini and take effect on next Receiver start.",
                 font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 8))

    # 5.3 -- live client-side filter over both boxes; no persistence, same
    # spirit as the worklist search. Filtering is applied by temporarily
    # swapping each box's displayed text for just the matching lines (and
    # locking it read-only while filtered) so an in-progress filter can't
    # accidentally clobber unsaved edits -- clearing the search restores
    # the exact original text.
    sop_search_row = ctk.CTkFrame(sop_outer, fg_color="transparent")
    sop_search_row.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(sop_search_row, text="Filter:", font=get_font("small"),
                 text_color=THEME_TEXT_MUTED).pack(side="left", padx=(0, 6))
    sop_search_var = ctk.StringVar(value="")
    sop_search_entry = ctk.CTkEntry(sop_search_row, textvariable=sop_search_var,
                                     placeholder_text="Search name or UID...", width=280)
    sop_search_entry.pack(side="left")

    sop_split = ctk.CTkFrame(sop_outer, fg_color="transparent")
    sop_split.pack(fill="both", expand=True)

    sop_left = ctk.CTkFrame(sop_split)
    sop_left.pack(side="left", fill="both", expand=True, padx=(0, 8))
    ctk.CTkLabel(sop_left, text="[SOP_CLASSES]  name = UID",
                 font=get_font("body", "bold")).pack(anchor="w", padx=8, pady=4)
    sop_classes_box = ctk.CTkTextbox(sop_left, font=ctk.CTkFont(family="Courier", size=11))
    sop_classes_box.pack(fill="both", expand=True, padx=8, pady=4)

    sop_right = ctk.CTkFrame(sop_split)
    sop_right.pack(side="left", fill="both", expand=True)
    ctk.CTkLabel(sop_right, text="[TRANSFER_SYNTAXES]  name = UID",
                 font=get_font("body", "bold")).pack(anchor="w", padx=8, pady=4)
    sop_ts_box = ctk.CTkTextbox(sop_right, font=ctk.CTkFont(family="Courier", size=11))
    sop_ts_box.pack(fill="both", expand=True, padx=8, pady=4)

    sop_btn_row = ctk.CTkFrame(sop_outer, fg_color="transparent")
    sop_btn_row.pack(fill="x", pady=6)
    sop_load_btn = ctk.CTkButton(sop_btn_row, text="↺ Reload from file", width=180)
    sop_load_btn.pack(side="left", padx=8)
    sop_save_btn = ctk.CTkButton(sop_btn_row, text="Save sopclass.ini", width=180,
                                  fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    sop_save_btn.pack(side="left", padx=8)
    sop_status_lbl = ctk.CTkLabel(sop_btn_row, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    sop_status_lbl.pack(side="left", padx=8)

    # §fix: surfaces allow_lossy_transcode (see _transcode_dataset_for_target())
    # right next to the transfer-syntax list it governs, rather than burying a
    # patient-safety-relevant toggle in a generic settings screen.
    sop_lossy_var = ctk.BooleanVar(value=APP_SETTINGS.get("allow_lossy_transcode", False))
    ctk.CTkCheckBox(
        sop_outer, text="Allow lossy transcode when pushing (only if no lossless/uncompressed "
                         "transfer syntax is accepted by the destination)",
        variable=sop_lossy_var, font=get_font("small"),
        command=lambda: do_settings_set("allow_lossy_transcode", sop_lossy_var.get()),
    ).pack(anchor="w", pady=(4, 0))

    # ---- Logs tab --------------------------------------------------
    # Logs are immutable audit records: there is intentionally no "Clear
    # Log" action anywhere in this UI. Old content is rotated out into
    # LOG_ARCHIVE_DIR (browsable below) on a retention schedule instead
    # of being deleted; the only way logs leave the app is Export.
    tab_logs = tabview.add("Logs")

    logs_top = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_top.pack(fill="x", padx=10, pady=(8, 4))

    log_file_var = ctk.StringVar(value=RECEIVER_LOG)
    log_selector = ctk.CTkSegmentedButton(
        logs_top,
        values=[RECEIVER_LOG, PUSH_LOG, AUDIT_LOG, APP_LOG],
        variable=log_file_var,
    )
    log_selector.pack(side="left")

    log_refresh_btn = ctk.CTkButton(logs_top, text="Refresh", width=110, image=get_icon("refresh-cw", size=14, color=THEME_TEXT), compound="left")
    log_refresh_btn.pack(side="left", padx=10)

    log_tail_var = ctk.BooleanVar(value=True)
    ctk.CTkCheckBox(logs_top, text="Auto-tail", variable=log_tail_var).pack(side="left", padx=10)

    logs_btn_row, (log_archive_now_btn, log_export_btn, log_copy_btn,
                   log_export_filtered_btn) = make_equal_button_row(
        tab_logs,
        [
            dict(text="Archive Now"),
            dict(text="Export Logs…", fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER),
            dict(text="Copy View", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
            dict(text="Export Filtered to CSV", fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER),
        ],
        pady=(0, 4), fill="x",
    )
    logs_btn_row.pack_configure(padx=10)

    logs_row1b = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_row1b.pack(fill="x", padx=10, pady=(0, 4))

    ctk.CTkLabel(logs_row1b, text="View:", font=get_font("small")).pack(side="left", padx=(0, 6))
    log_display_mode_var = ctk.StringVar(value="Raw Text")
    log_display_mode_menu = ctk.CTkSegmentedButton(
        logs_row1b, values=["Raw Text", "Structured"], variable=log_display_mode_var,
    )
    log_display_mode_menu.pack(side="left", padx=(0, 16))

    ctk.CTkLabel(logs_row1b, text="Search:", font=get_font("small")).pack(side="left", padx=(0, 6))
    log_search_var = ctk.StringVar(value="")
    log_search_entry = ctk.CTkEntry(
        logs_row1b, width=280, textvariable=log_search_var,
        placeholder_text="Patient ID/Name, institution, destination, text…")
    log_search_entry.pack(side="left")

    ctk.CTkLabel(logs_row1b, text="Severity:", font=get_font("small")).pack(side="left", padx=(16, 6))
    log_severity_var = ctk.StringVar(value="All")
    log_severity_menu = ctk.CTkSegmentedButton(
        logs_row1b, values=["All", "Errors Only", "Success Only"], variable=log_severity_var,
    )
    log_severity_menu.pack(side="left")

    log_match_count_lbl = ctk.CTkLabel(logs_row1b, text="", font=get_font("small"),
                                        text_color=THEME_TEXT_MUTED)
    log_match_count_lbl.pack(side="left", padx=12)

    logs_row2 = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_row2.pack(fill="x", padx=10, pady=(0, 8))

    ctk.CTkLabel(logs_row2, text="Retention:", font=get_font("small")).pack(side="left", padx=(0, 6))
    log_retention_var = ctk.StringVar(value=str(load_log_retention_days()))
    log_retention_menu = ctk.CTkOptionMenu(
        logs_row2, variable=log_retention_var, width=90,
        values=[str(d) for d in VALID_LOG_RETENTION_DAYS],
    )
    log_retention_menu.pack(side="left")
    ctk.CTkLabel(logs_row2, text="days before archiving", font=get_font("small"),
                 text_color=THEME_TEXT_MUTED).pack(side="left", padx=(4, 20))

    ctk.CTkLabel(logs_row2, text="View:", font=get_font("small")).pack(side="left", padx=(0, 6))
    log_view_mode_var = ctk.StringVar(value="Live")
    log_view_mode_menu = ctk.CTkSegmentedButton(
        logs_row2, values=["Live", "Archived"], variable=log_view_mode_var,
    )
    log_view_mode_menu.pack(side="left")

    log_archive_var = ctk.StringVar(value="")
    log_archive_menu = ctk.CTkOptionMenu(logs_row2, variable=log_archive_var, width=280, values=["(no archives yet)"])
    log_archive_menu.pack(side="left", padx=(10, 0))

    log_status_lbl = ctk.CTkLabel(logs_row2, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    log_status_lbl.pack(side="left", padx=12)

    logs_row3 = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_row3.pack(fill="x", padx=10, pady=(0, 4))
    ctk.CTkLabel(logs_row3, text="Destination:", font=get_font("small")).pack(side="left", padx=(0, 6))
    log_struct_dest_var = ctk.StringVar(value="All")
    log_struct_dest_menu = ctk.CTkOptionMenu(logs_row3, variable=log_struct_dest_var, width=200,
                                              values=["All"])
    log_struct_dest_menu.pack(side="left", padx=(0, 16))
    ctk.CTkLabel(logs_row3, text="Range:", font=get_font("small")).pack(side="left", padx=(0, 6))
    log_struct_range_var = ctk.StringVar(value="Last 7 Days")
    log_struct_range_menu = ctk.CTkOptionMenu(logs_row3, variable=log_struct_range_var, width=140,
                                               values=LOG_STRUCT_RANGE_OPTIONS)
    log_struct_range_menu.pack(side="left")

    # B.3: DOC-TRANSFER* event types as first-class filters, not just
    # unfiltered rows a user has to already know to search for.
    ctk.CTkLabel(logs_row3, text="Event Type:", font=get_font("small")).pack(side="left", padx=(16, 6))
    log_struct_event_type_var = ctk.StringVar(value="All")
    log_struct_event_type_menu = ctk.CTkOptionMenu(
        logs_row3, variable=log_struct_event_type_var, width=200,
        values=["All", "DOC-TRANSFER", "DOC-TRANSFER-RECEIVED", "DOC-TRANSFER-FAILED",
                "DOC-TRANSFER-OK", "DOC-TRANSFER-SERVER-START", "DOC-TRANSFER-SERVER-STOP",
                "DOC-TRANSFER-SERVER-CRASH"])
    log_struct_event_type_menu.pack(side="left")

    log_open_related_btn = ctk.CTkButton(logs_row3, text="Open Related Object", width=180,
                                          fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER, state="disabled")
    log_open_related_btn.pack(side="right", padx=4)

    logs_body = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # ---- Raw Text container (existing behavior, unchanged) ----
    log_raw_container = ctk.CTkFrame(logs_body, fg_color="transparent")
    logs_text = ctk.CTkTextbox(log_raw_container, font=ctk.CTkFont(family="Courier", size=11), state="disabled")
    logs_text.pack(fill="both", expand=True)
    log_text_ref["box"] = logs_text

    # ---- Structured container (new): table + expandable detail panel ----
    log_structured_container = ctk.CTkFrame(logs_body, fg_color="transparent")
    log_struct_split = ctk.CTkFrame(log_structured_container, fg_color="transparent")
    log_struct_split.pack(fill="both", expand=True)

    log_struct_tree_frame = ttk.Frame(log_struct_split)
    log_struct_tree_frame.pack(fill="both", expand=True, side="top")
    log_struct_cols = ("time", "event", "result", "patient_id", "patient_name", "destination", "error")
    log_struct_tree = ttk.Treeview(log_struct_tree_frame, columns=log_struct_cols, show="headings", height=12)
    log_struct_headings = {
        "time": "Time", "event": "Event", "result": "Result", "patient_id": "Patient ID",
        "patient_name": "Patient Name", "destination": "Destination", "error": "Error",
    }
    log_struct_widths = {
        "time": 150, "event": 90, "result": 80, "patient_id": 110,
        "patient_name": 160, "destination": 140, "error": 260,
    }
    for col in log_struct_cols:
        log_struct_tree.heading(col, text=log_struct_headings[col],
                                 command=lambda c=col: sort_tree(log_struct_tree, c, False))
        log_struct_tree.column(col, width=log_struct_widths[col], anchor="w")
    log_struct_vsb = ttk.Scrollbar(log_struct_tree_frame, orient="vertical", command=log_struct_tree.yview)
    log_struct_tree.configure(yscrollcommand=log_struct_vsb.set)
    log_struct_tree.pack(side="left", fill="both", expand=True)
    log_struct_vsb.pack(side="right", fill="y")

    ctk.CTkLabel(log_structured_container, text="Details (select a row — includes stack trace when present)",
                 font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(8, 2))
    log_detail_box = ctk.CTkTextbox(log_structured_container, height=140,
                                     font=ctk.CTkFont(family="Courier", size=11), state="disabled")
    log_detail_box.pack(fill="x")

    log_raw_container.pack(fill="both", expand=True)  # default mode = Raw Text

    # =========================================================
    # SETTINGS TAB (Priority 10 -- Settings & Performance)
    # =========================================================
    tab_settings = tabview.add("Settings")
    settings_outer = ctk.CTkScrollableFrame(tab_settings, fg_color="transparent")
    settings_outer.pack(fill="both", expand=True, padx=10, pady=(6, 14))

    settings_header = ctk.CTkFrame(settings_outer, fg_color="transparent")
    settings_header.pack(fill="x", pady=(4, 10))
    ctk.CTkLabel(settings_header, text="", image=get_icon("settings-2", size=20, color=THEME_TEXT),
                 compound="left").pack(side="left", padx=(0, 8))
    ctk.CTkLabel(settings_header, text="Settings", font=get_font("title", "bold")).pack(side="left")
    settings_status_lbl = ctk.CTkLabel(settings_header, text="", font=get_font("small"), text_color=THEME_TEXT_MUTED)
    settings_status_lbl.pack(side="right")

    def _settings_section(title, subtitle=None):
        card = make_card(settings_outer, title=title, subtitle=subtitle)
        card.pack(fill="x", pady=(0, 12))
        return card.body

    def _settings_row(parent, label_text, hint=None, hint_icon=None, hint_color=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=4)
        left = ctk.CTkFrame(row, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(left, text=label_text, font=get_font("body"), text_color=THEME_TEXT).pack(anchor="w")
        if hint:
            color = hint_color or THEME_TEXT_MUTED
            if hint_icon:
                ctk.CTkLabel(left, text=hint, image=get_icon(hint_icon, size=12, color=color),
                             compound="left", font=get_font("micro"), text_color=color,
                             justify="left", wraplength=500, anchor="w").pack(anchor="w")
            else:
                make_wrapped_label(left, hint, 520, font=get_font("micro"),
                                    text_color=color).pack(anchor="w")
        ctrl_holder = ctk.CTkFrame(row, fg_color="transparent")
        ctrl_holder.pack(side="right")
        return ctrl_holder

    # ---- Performance ----
    perf_body = _settings_section("Performance", "Controls how often the app refreshes itself.")

    make_wrapped_label(
        perf_body,
        "Lower refresh rates use less CPU on slower computers; higher refresh "
        "rates feel more real-time on faster ones. This applies to Dashboard, Receiver/Pusher "
        "monitoring, the Offline Queue, Logs, Destination Health, and Statistics alike.",
        760, font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 8))

    rr_ctrl = _settings_row(perf_body, "Application Refresh Rate")
    settings_refresh_rate_var = ctk.StringVar(value=f"{APP_SETTINGS['refresh_rate_ms']} ms")
    ctk.CTkOptionMenu(rr_ctrl, variable=settings_refresh_rate_var, width=120,
                      values=[f"{ms} ms" for ms in REFRESH_RATE_CHOICES_MS],
                      command=do_settings_refresh_rate_change).pack()

    pm_ctrl = _settings_row(perf_body, "Performance Mode",
                            "Picking a mode seeds the refresh rate and worker-thread settings below with a sensible starting point.")
    settings_perf_mode_var = ctk.StringVar(value=APP_SETTINGS["performance_mode"])
    ctk.CTkSegmentedButton(pm_ctrl, variable=settings_perf_mode_var,
                           values=list(PERFORMANCE_MODE_PRESETS.keys()),
                           command=do_settings_performance_mode_change).pack()

    graphs_ctrl = _settings_row(perf_body, "Disable Graphs",
                                "Turns off the trend-line sparklines on the Dashboard and Performance tabs. "
                                "They're redrawn on every refresh tick, so this is one of the quicker ways "
                                "to lighten the load on a slower machine.")
    settings_disable_graphs_var = ctk.BooleanVar(value=APP_SETTINGS.get("disable_graphs", False))

    def do_settings_disable_graphs_change():
        _update_app_setting("disable_graphs", settings_disable_graphs_var.get())
        # Force every already-drawn sparkline to repaint immediately
        # (as a placeholder or a real chart) instead of waiting for the
        # next periodic refresh tick.
        for _canvas in list(dash_graph_canvases.values()):
            try:
                if _canvas.winfo_exists():
                    _draw_sparkline(_canvas, [], THEME_ACCENT)
            except Exception:
                pass
        for _canvas, _color in list(perf_graph_canvases.values()):
            try:
                if _canvas.winfo_exists():
                    _draw_sparkline(_canvas, [], _color)
            except Exception:
                pass
        try:
            refresh_home_dashboard()
        except Exception:
            pass
        try:
            refresh_performance_tab()
        except Exception:
            pass

    ctk.CTkSwitch(graphs_ctrl, text="", variable=settings_disable_graphs_var,
                  command=do_settings_disable_graphs_change).pack()

    perf_stats_row = make_card_row(perf_body, pady=(8, 0))
    settings_perf_stat_lbls = {}

    def _settings_perf_stat(key, title):
        card = ctk.CTkFrame(perf_stats_row, corner_radius=12, fg_color=THEME_HEADING_BG)
        add_card_to_row(perf_stats_row, card, padx=4, pady=0)
        ctk.CTkLabel(card, text=title, font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(pady=(8, 0))
        lbl = ctk.CTkLabel(card, text="—", font=get_font("section", "bold"), text_color=THEME_TEXT)
        lbl.pack(pady=(0, 8))
        settings_perf_stat_lbls[key] = lbl

    _settings_perf_stat("refresh_freq", "Current UI Refresh Frequency")
    _settings_perf_stat("avg_update_ms", "Avg UI Update Time")
    _settings_perf_stat("cpu", "CPU Usage")
    _settings_perf_stat("memory", "Memory Usage")

    # ---- Appearance ----
    appear_body = _settings_section("Appearance")

    ui_scale_ctrl = _settings_row(appear_body, "UI Scaling")
    settings_ui_scale_var = ctk.StringVar(value=f"{APP_SETTINGS['ui_scaling_pct']}%")
    ctk.CTkOptionMenu(ui_scale_ctrl, variable=settings_ui_scale_var, width=100,
                      values=["90%", "100%", "110%", "125%", "150%"],
                      command=do_settings_ui_scaling_change).pack()

    font_scale_ctrl = _settings_row(appear_body, "Font Size")
    settings_font_scale_var = ctk.StringVar(value=f"{APP_SETTINGS['font_scale_pct']}%")
    ctk.CTkOptionMenu(font_scale_ctrl, variable=settings_font_scale_var, width=100,
                      values=["90%", "100%", "110%", "125%", "150%"],
                      command=do_settings_font_scale_change).pack()

    compact_ctrl = _settings_row(appear_body, "Compact Mode", "Tighter spacing throughout the app. Applies on restart.")
    settings_compact_var = ctk.BooleanVar(value=APP_SETTINGS["compact_mode"])
    ctk.CTkCheckBox(compact_ctrl, text="", variable=settings_compact_var,
                 command=lambda: do_settings_set("compact_mode", settings_compact_var.get(),
                                                 "Compact Mode saved -- restart to apply.")).pack()

    remember_geom_ctrl = _settings_row(appear_body, "Remember Window Size and Position")
    settings_remember_geom_var = ctk.BooleanVar(value=APP_SETTINGS["remember_window_geometry"])
    ctk.CTkCheckBox(remember_geom_ctrl, text="", variable=settings_remember_geom_var,
                 command=lambda: do_settings_set("remember_window_geometry", settings_remember_geom_var.get())).pack()

    launch_max_ctrl = _settings_row(appear_body, "Launch Application Maximized", "Applies the next time the app starts.")
    settings_launch_max_var = ctk.BooleanVar(value=APP_SETTINGS["launch_maximized"])
    ctk.CTkCheckBox(launch_max_ctrl, text="", variable=settings_launch_max_var,
                 command=lambda: do_settings_set("launch_maximized", settings_launch_max_var.get(),
                                                 "Saved -- takes effect on next launch.")).pack()

    # ---- Notifications ----
    notif_body = _settings_section("Notifications")

    toast_ctrl = _settings_row(notif_body, "Toast Notifications")
    settings_toast_var = ctk.BooleanVar(value=APP_SETTINGS["toast_notifications_enabled"])
    ctk.CTkCheckBox(toast_ctrl, text="", variable=settings_toast_var,
                 command=lambda: do_settings_set("toast_notifications_enabled", settings_toast_var.get())).pack()

    sound_ctrl = _settings_row(notif_body, "Notification Sounds")
    settings_sound_var = ctk.BooleanVar(value=APP_SETTINGS["notification_sounds_enabled"])
    ctk.CTkCheckBox(sound_ctrl, text="", variable=settings_sound_var,
                 command=lambda: do_settings_set("notification_sounds_enabled", settings_sound_var.get())).pack()

    dur_ctrl = _settings_row(notif_body, "Notification Duration")
    settings_notif_duration_var = ctk.StringVar(value=f"{APP_SETTINGS['notification_duration_sec']} sec")
    ctk.CTkOptionMenu(dur_ctrl, variable=settings_notif_duration_var, width=100,
                      values=["2 sec", "4 sec", "6 sec", "8 sec", "10 sec"],
                      command=lambda v: do_settings_set(
                          "notification_duration_sec", int(v.split()[0]),
                          f"Notification duration set to {v}.")).pack()

    crit_ctrl = _settings_row(notif_body, "Critical-Only Notifications",
                              "Only show toasts for failures/offline/critical events.")
    settings_critical_only_var = ctk.BooleanVar(value=APP_SETTINGS["critical_only_notifications"])
    ctk.CTkCheckBox(crit_ctrl, text="", variable=settings_critical_only_var,
                 command=lambda: do_settings_set("critical_only_notifications",
                                                 settings_critical_only_var.get())).pack()

    # ---- Startup & Behavior ----
    startup_body = _settings_section("Startup & Behavior")

    launch_startup_ctrl = _settings_row(startup_body, "Launch on System Startup")
    settings_launch_startup_var = ctk.BooleanVar(value=APP_SETTINGS["launch_on_system_startup"])
    ctk.CTkCheckBox(launch_startup_ctrl, text="", variable=settings_launch_startup_var,
                 command=lambda: do_settings_launch_on_startup_change(settings_launch_startup_var.get())).pack()

    auto_start_recv_ctrl = _settings_row(startup_body, "Automatically Start Receiver When App Opens",
                                         "Applies on next launch.")
    settings_auto_start_recv_var = ctk.BooleanVar(value=APP_SETTINGS["auto_start_receiver_on_launch"])
    ctk.CTkCheckBox(auto_start_recv_ctrl, text="", variable=settings_auto_start_recv_var,
                 command=lambda: do_settings_set("auto_start_receiver_on_launch",
                                                 settings_auto_start_recv_var.get(),
                                                 "Saved -- takes effect on next launch.")).pack()

    auto_restart_ctrl = _settings_row(startup_body, "Automatically Restart Receiver If It Stops Unexpectedly")
    settings_auto_restart_var = ctk.BooleanVar(value=APP_SETTINGS["auto_restart_receiver"])
    ctk.CTkCheckBox(auto_restart_ctrl, text="", variable=settings_auto_restart_var,
                 command=lambda: do_settings_set("auto_restart_receiver", settings_auto_restart_var.get())).pack()

    restore_session_ctrl = _settings_row(startup_body, "Restore Previous Session on Startup", "Applies on next launch.")
    settings_restore_session_var = ctk.BooleanVar(value=APP_SETTINGS["restore_previous_session"])
    ctk.CTkCheckBox(restore_session_ctrl, text="", variable=settings_restore_session_var,
                 command=lambda: do_settings_set("restore_previous_session",
                                                 settings_restore_session_var.get(),
                                                 "Saved -- takes effect on next launch.")).pack()

    minimize_tray_ctrl = _settings_row(startup_body, "Minimize to System Tray Instead of Exiting")
    settings_minimize_tray_var = ctk.BooleanVar(value=APP_SETTINGS["minimize_to_tray_on_close"])
    ctk.CTkCheckBox(minimize_tray_ctrl, text="", variable=settings_minimize_tray_var,
                 command=lambda: do_settings_set("minimize_to_tray_on_close",
                                                 settings_minimize_tray_var.get())).pack()

    confirm_close_ctrl = _settings_row(startup_body, "Confirmation Dialog Before Closing")
    settings_confirm_close_var = ctk.BooleanVar(value=APP_SETTINGS["confirm_before_close"])
    ctk.CTkCheckBox(confirm_close_ctrl, text="", variable=settings_confirm_close_var,
                 command=lambda: do_settings_set("confirm_before_close", settings_confirm_close_var.get())).pack()

    check_updates_ctrl = _settings_row(startup_body, "Automatically Check for Updates",
                                       "No updater is configured yet -- this just reserves the preference for when one is added.")
    settings_check_updates_var = ctk.BooleanVar(value=APP_SETTINGS["auto_check_for_updates"])
    ctk.CTkCheckBox(check_updates_ctrl, text="", variable=settings_check_updates_var,
                 command=lambda: do_settings_set("auto_check_for_updates", settings_check_updates_var.get())).pack()

    # ---- Logging & Diagnostics ----
    diag_body = _settings_section("Logging & Diagnostics")

    log_level_ctrl = _settings_row(diag_body, "Logging Level")
    settings_log_level_var = ctk.StringVar(value=APP_SETTINGS["logging_level"])
    ctk.CTkSegmentedButton(log_level_ctrl, variable=settings_log_level_var, values=LOG_LEVEL_CHOICES,
                           command=do_settings_logging_level_change).pack()

    # §fix: live toggle for pynetdicom's own verbose PDU/DIMSE tracing (see
    # the debug_logger() gating note near the pynetdicom imports, and in
    # startup()). Applies immediately -- no restart needed in either
    # direction -- so it's actually usable as a "flip on, reproduce the
    # issue, flip back off" diagnostic tool instead of a permanent,
    # unmanaged stdout log stream.
    def _on_verbose_protocol_logging_toggle(value):
        do_settings_set("verbose_dicom_protocol_logging", value)
        try:
            if value:
                debug_logger()
            else:
                logging.getLogger("pynetdicom").handlers = []
                logging.getLogger("pynetdicom").setLevel(logging.WARNING)
        except Exception:
            log_exception("Failed to toggle verbose DICOM protocol logging")

    verbose_proto_var = ctk.BooleanVar(value=APP_SETTINGS.get("verbose_dicom_protocol_logging", False))
    ctk.CTkCheckBox(
        diag_body, text="Verbose DICOM protocol logging (full PDU/DIMSE traces to console — "
                         "for diagnosing transfer-syntax/negotiation issues only)",
        variable=verbose_proto_var, font=get_font("small"),
        command=lambda: _on_verbose_protocol_logging_toggle(verbose_proto_var.get()),
    ).pack(anchor="w", pady=(4, 10))

    diag_btn_row = ctk.CTkFrame(diag_body, fg_color="transparent")
    diag_btn_row.pack(fill="x", pady=(4, 10))
    ctk.CTkButton(diag_btn_row, text="Clear Temporary Cache", width=180,
                 command=do_settings_clear_cache).pack(side="left", padx=(0, 8))
    ctk.CTkButton(diag_btn_row, text="Remove Archived Logs", width=180,
                 fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER,
                 command=do_settings_remove_archived_logs).pack(side="left", padx=8)
    ctk.CTkButton(diag_btn_row, text="Open Config Directory", width=180,
                 fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                 command=do_settings_open_config_dir).pack(side="left", padx=8)

    diag_grid = ctk.CTkFrame(diag_body, fg_color=THEME_SURFACE, corner_radius=10)
    diag_grid.pack(fill="x", pady=(4, 0))
    settings_diag_lbls = {}

    def _diag_line(label_text, key):
        row = ctk.CTkFrame(diag_grid, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=label_text, font=get_font("small"), text_color=THEME_TEXT_MUTED,
                     width=200, anchor="w").pack(side="left")
        lbl = ctk.CTkLabel(row, text="—", font=get_font("small", "bold"), text_color=THEME_TEXT, anchor="w")
        lbl.pack(side="left")
        settings_diag_lbls[key] = lbl

    _diag_line("Application Version", "version")
    _diag_line("Python Version", "python")
    _diag_line("Operating System", "os")
    _diag_line("Application Uptime", "uptime")
    _diag_line("Current Memory Usage", "memory_usage")
    _diag_line("Active Background Workers", "workers")
    _diag_line("Offline Queue Size", "queue_size")
    _diag_line("Current Receiver Status", "receiver_status")

    # ---- Data & Safety ----
    safety_body = _settings_section("Data & Safety")
    safety_btn_row = ctk.CTkFrame(safety_body, fg_color="transparent")
    safety_btn_row.pack(fill="x")
    ctk.CTkButton(safety_btn_row, text="Export All Settings…", width=180,
                 fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER,
                 command=do_settings_export).pack(side="left", padx=(0, 8), pady=4)
    ctk.CTkButton(safety_btn_row, text="Import Settings…", width=160,
                 command=do_settings_import).pack(side="left", padx=8, pady=4)
    ctk.CTkButton(safety_btn_row, text="Restore Factory Defaults", width=200,
                 fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER,
                 command=do_settings_restore_defaults).pack(side="left", padx=8, pady=4)
    safety_btn_row2 = ctk.CTkFrame(safety_body, fg_color="transparent")
    safety_btn_row2.pack(fill="x")
    ctk.CTkButton(safety_btn_row2, text="Backup Every Configuration File…", width=240,
                 command=do_settings_manual_backup).pack(side="left", padx=(0, 8), pady=4)
    ctk.CTkButton(safety_btn_row2, text="Verify Configuration Integrity", width=220,
                 fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                 command=do_settings_verify_integrity).pack(side="left", padx=8, pady=4)

    # ---- Accessibility ----
    a11y_body = _settings_section("Accessibility")

    hc_ctrl = _settings_row(a11y_body, "High Contrast Mode", "Boosts text/background contrast throughout the app. Applies immediately.")
    settings_hc_var = ctk.BooleanVar(value=APP_SETTINGS["high_contrast_mode"])
    ctk.CTkCheckBox(hc_ctrl, text="", variable=settings_hc_var,
                 command=lambda: do_settings_accessibility_change(
                     "high_contrast_mode", settings_hc_var.get(), "High Contrast Mode")).pack()

    click_ctrl = _settings_row(a11y_body, "Larger Click Targets", "Increases worklist row height and nav-button height throughout the app. Applies immediately.")
    settings_click_targets_var = ctk.BooleanVar(value=APP_SETTINGS["larger_click_targets"])
    ctk.CTkCheckBox(click_ctrl, text="", variable=settings_click_targets_var,
                 command=lambda: do_settings_accessibility_change(
                     "larger_click_targets", settings_click_targets_var.get(), "Larger Click Targets")).pack()

    motion_ctrl = _settings_row(a11y_body, "Reduced Motion Mode", "Stops long text (status bar, KPI tiles) from scrolling -- shows it truncated instead. Applies immediately.")
    settings_reduced_motion_var = ctk.BooleanVar(value=APP_SETTINGS["reduced_motion"])
    ctk.CTkCheckBox(motion_ctrl, text="", variable=settings_reduced_motion_var,
                 command=lambda: do_settings_reduced_motion_change(settings_reduced_motion_var.get())).pack()

    shortcuts_ctrl = _settings_row(a11y_body, "Keyboard Shortcuts")
    ctk.CTkButton(shortcuts_ctrl, text="View Shortcuts", width=150,
                 command=do_settings_show_shortcuts).pack()

    # ---- Advanced ----
    adv_body = _settings_section("Advanced")

    workers_ctrl = _settings_row(adv_body, "Maximum Background Worker Threads")
    settings_max_workers_var = ctk.StringVar(value=str(APP_SETTINGS["max_worker_threads"]))
    ctk.CTkOptionMenu(workers_ctrl, variable=settings_max_workers_var, width=90,
                      values=[str(n) for n in (1, 2, 4, 6, 8, 12, 16)],
                      command=lambda v: do_settings_set("max_worker_threads", int(v),
                                                        f"Max worker threads set to {v}.")).pack()

    pushes_ctrl = _settings_row(adv_body, "Maximum Simultaneous Push Operations")
    settings_max_pushes_var = ctk.StringVar(value=str(APP_SETTINGS["max_simultaneous_pushes"]))
    ctk.CTkOptionMenu(pushes_ctrl, variable=settings_max_pushes_var, width=90,
                      values=[str(n) for n in (1, 2, 4, 6, 8, 12, 16)],
                      command=lambda v: do_settings_set("max_simultaneous_pushes", int(v),
                                                        f"Max simultaneous pushes set to {v}.")).pack()

    timeout_ctrl = _settings_row(adv_body, "Network Timeout")
    settings_net_timeout_var = ctk.StringVar(value=f"{APP_SETTINGS['network_timeout_sec']} sec")
    ctk.CTkOptionMenu(timeout_ctrl, variable=settings_net_timeout_var, width=100,
                      values=["5 sec", "10 sec", "15 sec", "30 sec", "60 sec", "120 sec"],
                      command=lambda v: do_settings_set("network_timeout_sec", int(v.split()[0]),
                                                        f"Network timeout set to {v}.")).pack()

    cache_limit_ctrl = _settings_row(adv_body, "Cache Size Limit", "Also used as the log-rotation size threshold.")
    settings_cache_limit_var = ctk.StringVar(value=f"{APP_SETTINGS['cache_size_limit_mb']} MB")
    ctk.CTkOptionMenu(cache_limit_ctrl, variable=settings_cache_limit_var, width=100,
                      values=["5 MB", "10 MB", "25 MB", "50 MB", "100 MB", "250 MB"],
                      command=lambda v: do_settings_set("cache_size_limit_mb", int(v.split()[0]),
                                                        f"Cache size limit set to {v}.")).pack()

    exp_ctrl = _settings_row(adv_body, "Experimental Features",
                             "Unvalidated and subject to change without notice.",
                             hint_icon="triangle-alert", hint_color=THEME_WARNING)
    settings_experimental_var = ctk.BooleanVar(value=APP_SETTINGS["experimental_features_enabled"])
    ctk.CTkCheckBox(exp_ctrl, text="", variable=settings_experimental_var,
                 command=lambda: do_settings_toggle_experimental(settings_experimental_var.get())).pack()

    doc_transfer_ctrl = _settings_row(
        adv_body, "Document Transfer (Receiver)",
        "Lets the receiver accept Radiology_Report.docx / Patient_History.txt "
        "from the pusher over a separate small TCP connection. Each destination "
        "still opts in separately on the Destinations tab. Takes effect the next "
        "time the receiver is started.")
    settings_doc_transfer_var = ctk.BooleanVar(value=APP_SETTINGS.get("doc_transfer_enabled", False))
    ctk.CTkCheckBox(doc_transfer_ctrl, text="", variable=settings_doc_transfer_var,
                 command=lambda: do_settings_set(
                     "doc_transfer_enabled", settings_doc_transfer_var.get(),
                     "Document transfer enabled." if settings_doc_transfer_var.get()
                     else "Document transfer disabled.")).pack()

    doc_transfer_port_ctrl = _settings_row(
        adv_body, "Document Transfer Port",
        "TCP port the receiver listens on for document transfers.")
    settings_doc_transfer_port_var = ctk.StringVar(
        value=str(APP_SETTINGS.get("doc_transfer_receiver_port", "11244")))
    ctk.CTkEntry(doc_transfer_port_ctrl, textvariable=settings_doc_transfer_port_var, width=90).pack(side="left")
    ctk.CTkButton(doc_transfer_port_ctrl, text="Save", width=60,
                 command=lambda: do_settings_save_doc_transfer_port(
                     settings_doc_transfer_port_var.get())).pack(side="left", padx=(6, 0))

    # A.5: configurable chunk size / connect timeout, read by
    # push_patient_documents / doc_transfer_accept_loop /
    # _handle_doc_transfer_connection instead of the bare constants.
    doc_transfer_chunk_ctrl = _settings_row(
        adv_body, "Document Transfer Chunk Size",
        "Bytes per network read/write while streaming Report/History files.")
    settings_doc_transfer_chunk_var = ctk.StringVar(
        value=f"{APP_SETTINGS.get('doc_transfer_chunk_size_kb', 64)} KB")
    ctk.CTkOptionMenu(doc_transfer_chunk_ctrl, variable=settings_doc_transfer_chunk_var, width=100,
                      values=["16 KB", "32 KB", "64 KB", "128 KB", "256 KB", "512 KB"],
                      command=lambda v: do_settings_set("doc_transfer_chunk_size_kb", int(v.split()[0]),
                                                        f"Document transfer chunk size set to {v}.")).pack()

    doc_transfer_timeout_ctrl = _settings_row(
        adv_body, "Document Transfer Timeout",
        "Connect/socket timeout for the document-transfer connection.")
    settings_doc_transfer_timeout_var = ctk.StringVar(
        value=f"{APP_SETTINGS.get('doc_transfer_timeout_sec', DOC_TRANSFER_CONNECT_TIMEOUT_SEC)} sec")
    ctk.CTkOptionMenu(doc_transfer_timeout_ctrl, variable=settings_doc_transfer_timeout_var, width=100,
                      values=["5 sec", "10 sec", "15 sec", "30 sec", "60 sec"],
                      command=lambda v: do_settings_set("doc_transfer_timeout_sec", int(v.split()[0]),
                                                        f"Document transfer timeout set to {v}.")).pack()

    doc_transfer_max_size_ctrl = _settings_row(
        adv_body, "Document Transfer Max File Size",
        "Files larger than this are rejected by the receiver before any data is written to disk.")
    settings_doc_transfer_max_size_var = ctk.StringVar(
        value=f"{APP_SETTINGS.get('doc_transfer_max_size_mb', 500)} MB")
    ctk.CTkOptionMenu(doc_transfer_max_size_ctrl, variable=settings_doc_transfer_max_size_var, width=100,
                      values=["10 MB", "25 MB", "50 MB", "100 MB", "250 MB", "500 MB", "1000 MB", "2000 MB"],
                      command=lambda v: do_settings_set("doc_transfer_max_size_mb", int(v.split()[0]),
                                                        f"Document transfer size limit set to {v}.")).pack()

    doc_transfer_auth_ctrl = _settings_row(
        adv_body, "Document Transfer Auth Key",
        "Shared secret every sender must present. Leave blank to accept unauthenticated "
        "transfers (matches every existing destination's key until you set one on both sides).",
        hint_icon="shield", hint_color=THEME_TEXT_MUTED)
    settings_doc_transfer_auth_var = ctk.StringVar(
        value=str(APP_SETTINGS.get("doc_transfer_receiver_auth_key", "")))
    ctk.CTkEntry(doc_transfer_auth_ctrl, textvariable=settings_doc_transfer_auth_var,
                 width=140, show="•").pack(side="left")
    ctk.CTkButton(doc_transfer_auth_ctrl, text="Save", width=60,
                 command=lambda: do_settings_save_doc_transfer_auth_key(
                     settings_doc_transfer_auth_var.get())).pack(side="left", padx=(6, 0))

    # §3.6 fix: an empty auth key means the receiver silently accepts
    # unauthenticated document-transfer connections from anyone on the
    # network -- previously there was no proactive nudge anywhere telling
    # an operator that. This banner sits directly under the field it's
    # about, shows/hides itself based on the CURRENT saved key (not just
    # whatever's typed and unsaved in the box), and is refreshed by
    # do_settings_save_doc_transfer_auth_key() after every Save -- no tab
    # rebuild or app restart needed to see it appear/disappear.
    global doc_transfer_auth_warning_label
    doc_transfer_auth_warning_label = make_wrapped_label(
        adv_body, "", 760, font=get_font("micro"), text_color=THEME_WARNING)
    doc_transfer_auth_warning_label.pack(anchor="w", pady=(0, 8))
    _refresh_doc_transfer_auth_warning()

    # D.3: admin idle-timeout auto-lock.
    idle_timeout_ctrl = _settings_row(
        adv_body, "Admin Idle Auto-Lock",
        "Automatically returns to the Lock screen after this many minutes of "
        "inactivity in Admin mode. Never interrupts an in-flight push/receive job. "
        "0 = disabled.")
    settings_admin_idle_timeout_var = ctk.StringVar(
        value=f"{APP_SETTINGS.get('admin_idle_timeout_min', 15)} min" if APP_SETTINGS.get('admin_idle_timeout_min', 15) else "Disabled")
    ctk.CTkOptionMenu(idle_timeout_ctrl, variable=settings_admin_idle_timeout_var, width=100,
                      values=["Disabled", "5 min", "10 min", "15 min", "30 min", "60 min"],
                      command=lambda v: do_settings_set(
                          "admin_idle_timeout_min", 0 if v == "Disabled" else int(v.split()[0]),
                          "Admin idle auto-lock disabled." if v == "Disabled"
                          else f"Admin idle auto-lock set to {v}.")).pack()

    # D.2 / C.6: passive inline warnings, recomputed whenever this tab is
    # opened/refreshed (see refresh_settings_diagnostics) -- never a
    # blocking dialog, never a new background thread.
    global settings_warnings_frame
    settings_warnings_frame = ctk.CTkFrame(adv_body, fg_color="transparent")
    settings_warnings_frame.pack(fill="x", padx=8, pady=(10, 4))

    refresh_settings_diagnostics()
    # correctly every time these tabs are rebuilt) ----
    qr_find_btn.configure(command=do_qr_find)
    qr_retrieve_btn.configure(command=do_qr_retrieve)

    dest_add_btn.configure(command=do_add_update_dest)
    dest_del_btn.configure(command=do_del_dest)
    dest_echo_btn.configure(command=do_echo_dest)
    dest_select_var.trace_add("write", lambda *_: do_load_dest_into_form())

    routing_add_btn.configure(command=do_add_routing_rule)
    routing_del_btn.configure(command=do_del_routing_rule)
    routing_up_btn.configure(command=lambda: do_move_rule(-1))
    routing_down_btn.configure(command=lambda: do_move_rule(1))
    routing_test_btn.configure(command=do_test_routing_rule)

    sop_save_btn.configure(command=do_save_sop)
    sop_load_btn.configure(command=lambda: (load_sop_into_editor(),
                                             sop_status_lbl.configure(text="Reloaded from sopclass.ini")))
    sop_search_var.trace_add("write", lambda *_: _apply_sop_filter())

    log_refresh_btn.configure(command=_refresh_current_log_view)
    log_archive_now_btn.configure(command=do_archive_logs_now)
    log_export_btn.configure(command=do_export_logs)
    log_file_var.trace_add("write", lambda *_: _refresh_current_log_view())
    log_view_mode_var.trace_add("write", lambda *_: (refresh_archive_list(), refresh_log_view()))
    log_archive_var.trace_add("write", lambda *_: refresh_log_view())
    log_retention_var.trace_add("write", lambda *_: do_change_log_retention())
    log_search_var.trace_add("write", lambda *_: _refresh_current_log_view())
    log_severity_var.trace_add("write", lambda *_: _refresh_current_log_view())
    log_copy_btn.configure(command=do_copy_log_view)
    log_export_filtered_btn.configure(command=export_structured_log_view_to_csv)
    log_display_mode_var.trace_add("write", _on_log_display_mode_change)
    log_struct_dest_var.trace_add("write", lambda *_: refresh_structured_log_view())
    log_struct_range_var.trace_add("write", lambda *_: refresh_structured_log_view())
    if "log_struct_event_type_var" in globals():
        log_struct_event_type_var.trace_add("write", lambda *_: refresh_structured_log_view())
    log_struct_tree.bind("<<TreeviewSelect>>", _on_log_struct_tree_select)
    log_open_related_btn.configure(command=do_open_related_log_object)
    _refresh_log_struct_destination_choices()

    for _name in ADMIN_ONLY_TAB_NAMES:
        add_nav_button(_name)
    admin_tabs_active["value"] = True

    # Populate the tabs. Widget construction above is already done and
    # painted; the actual data population (dashboard scan, log tail,
    # audit-log parse, disk_usage syscall) is deferred by one event-loop
    # tick via app.after(0, ...) so Tk gets a chance to render the newly
    # built tabs BEFORE doing this work. Previously all of this ran
    # synchronously in the same call that built ~6 tabs' worth of
    # widgets, which is what made switching into Admin mode visibly
    # freeze for a moment.
    def _populate_admin_tabs_deferred():
        refresh_destinations_ui()
        refresh_routing_ui()
        load_sop_into_editor()
        refresh_archive_list()
        refresh_log_view()
        refresh_health_tree()
        refresh_admin_dashboard()
        _update_export_scope_visibility()
        refresh_backup_history_ui()
        refresh_ldap_roster_ui()
        try:
            refresh_settings_diagnostics()
        except Exception:
            log_exception("Initial Settings diagnostics refresh failed")
        try:
            refresh_doc_transfer_tab()
        except Exception:
            log_exception("Initial Doc Transfer tab refresh failed")

    app.after(0, _populate_admin_tabs_deferred)


def tear_down_admin_only_tabs():
    """Removes the Admin-only tabs (and, per CTkTabview.delete(), everything
    built inside them) on the way back to User mode. build_admin_only_tabs()
    fully reconstructs them if Admin is unlocked again later -- there is no
    separate/duplicated construction path for that case."""
    for name in ADMIN_ONLY_TAB_NAMES:
        try:
            tabview.delete(name)
        except ValueError:
            pass
        remove_nav_button(name)
    admin_tabs_active["value"] = False
    _sync_nav_highlight()

# =========================================================
# OTP CONFIRMATION DIALOG (local, no external dependency)
# =========================================================

def local_otp_confirm(title, summary, on_confirmed):
    """Display a randomly generated 6-digit code that the user must retype
    before a destructive or permanent action proceeds. The code never leaves
    this machine."""
    code = "".join(random.choices(string.digits, k=6))

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title(title)
    win.geometry("420x280")
    _safe_grab_set(win)

    def _on_close():
        win.grab_release()
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", _on_close)

    make_wrapped_label(win, summary, 380, font=get_font("body")).pack(padx=15, pady=(15, 5))
    ctk.CTkLabel(win, text=f"Verification code:  {code}",
                 font=get_font("value_lg", "bold")).pack(pady=10)
    ctk.CTkLabel(win, text="Re-type the code above to confirm:").pack()
    code_entry = ctk.CTkEntry(win, width=180, justify="center")
    code_entry.pack(pady=8)

    def verify():
        if code_entry.get().strip() == code:
            win.grab_release()
            win.destroy()
            on_confirmed()
        else:
            modern_showerror("Invalid", "Code does not match.", parent=win)

    ctk.CTkButton(win, text="Confirm & Proceed", command=verify).pack(pady=12)

# =========================================================
# REFRESH HELPERS
# =========================================================

def refresh_worklists(search=""):
    # Each tree always filters using its own search box's current text --
    # previously whichever entry fired the refresh (Receiver or Pusher)
    # passed its text into BOTH trees, so typing in one tab's search could
    # silently filter the other tab's worklist too. populate_tree() always
    # matches against the full patient_data set before paginating, so
    # search here already covers every record, not just the visible page.
    populate_tree(rec_tree, rec_search_var.get())
    populate_tree(push_tree, push_search_var.get())
    refresh_receiver_badge()


def _log_line_matches_filters(line, search_lower, severity):
    if severity == "Errors Only" and "ERROR=" in line and line.rstrip().endswith("ERROR="):
        return False  # blank ERROR= means success; exclude from Errors Only
    if severity == "Success Only" and "ERROR=" in line and not line.rstrip().endswith("ERROR="):
        return False
    if search_lower and search_lower not in line.lower():
        return False
    return True


def _apply_raw_log_severity_tags(box, content):
    """6.2 -- colors raw-view lines by severity using the exact same
    classification _log_line_matches_filters() already applies (a
    non-blank 'ERROR=' means failure, everything else non-blank counts as
    success), so raw mode visually matches structured mode's tree-tag
    color-coding (log_struct_tree's 'failure'/'success' tags) instead of
    being plain monochrome text."""
    try:
        target = box if hasattr(box, "tag_config") else box._textbox
        target.tag_config("raw_log_error", foreground=THEME_DANGER)
        target.tag_config("raw_log_success", foreground=THEME_TEXT)
        target.tag_remove("raw_log_error", "1.0", "end")
        target.tag_remove("raw_log_success", "1.0", "end")
        for i, line in enumerate(content.splitlines(), start=1):
            if not line.strip():
                continue
            if "ERROR=" in line and not line.rstrip().endswith("ERROR="):
                target.tag_add("raw_log_error", f"{i}.0", f"{i}.end")
            else:
                target.tag_add("raw_log_success", f"{i}.0", f"{i}.end")
    except Exception:
        log_exception("Failed to apply raw log severity coloring")


def refresh_log_view():
    if not admin_tabs_active["value"]:
        return  # Logs tab doesn't exist in User mode
    box = log_text_ref.get("box")
    if not box:
        return

    if log_view_mode_var.get() == "Archived":
        archive_path = _resolve_selected_archive_path()
        content = read_archived_log(archive_path) if archive_path and os.path.exists(archive_path) else "(select an archive above)"
    else:
        path = log_file_var.get()
        lines = tail_log_file(path, max_lines=500)
        content = "".join(lines) if lines else "(log is empty)"

    search_lower = log_search_var.get().strip().lower() if "log_search_var" in globals() else ""
    severity = log_severity_var.get() if "log_severity_var" in globals() else "All"
    if search_lower or severity != "All":
        raw_lines = content.splitlines(keepends=True)
        filtered = [ln for ln in raw_lines if _log_line_matches_filters(ln, search_lower, severity)]
        if "log_match_count_lbl" in globals():
            log_match_count_lbl.configure(text=f"{len(filtered)} / {len(raw_lines)} lines match")
        content = "".join(filtered) if filtered else "(no lines match the current filters)"
    elif "log_match_count_lbl" in globals():
        log_match_count_lbl.configure(text="")

    box.configure(state="normal")
    box.delete("1.0", "end")
    box.insert("end", content)
    _apply_raw_log_severity_tags(box, content)
    if log_tail_var.get() and log_view_mode_var.get() == "Live" and not search_lower and severity == "All":
        box.see("end")
    box.configure(state="disabled")


def do_copy_log_view():
    """Copies exactly what's currently shown (post-filter) to the
    clipboard -- same clipboard_clear/append pattern used elsewhere."""
    box = log_text_ref.get("box")
    if not box:
        return
    try:
        text = box.get("1.0", "end-1c")
        app.clipboard_clear()
        app.clipboard_append(text)
        if "log_match_count_lbl" in globals() and log_match_count_lbl is not None:
            prior = log_match_count_lbl.cget("text")
            log_match_count_lbl.configure(text="Copied to clipboard")
            app.after(1500, lambda: log_match_count_lbl.configure(text=prior))
    except Exception:
        log_exception("Failed to copy log view to clipboard")


def _refresh_log_struct_destination_choices():
    """Populates the structured view's Destination filter from the same
    configured destinations list used everywhere else in the app."""
    try:
        names = sorted({d.get("name", "") for d in load_destinations() if d.get("name")})
    except Exception:
        names = []
    values = ["All"] + names
    log_struct_dest_menu.configure(values=values)
    if log_struct_dest_var.get() not in values:
        log_struct_dest_var.set("All")


def refresh_structured_log_view():
    if not admin_tabs_active["value"]:
        return
    if "log_struct_tree" not in globals():
        return
    for iid in log_struct_tree.get_children():
        log_struct_tree.delete(iid)
    log_struct_records_by_iid.clear()

    records = get_structured_log_records(
        module_text_path=log_file_var.get(),
        search_text=log_search_var.get(),
        destination=log_struct_dest_var.get(),
        severity=log_severity_var.get(),
        range_option=log_struct_range_var.get(),
        event_type=log_struct_event_type_var.get() if "log_struct_event_type_var" in globals() else "All",
    )
    for i, rec in enumerate(records):
        iid = f"log{i}"
        is_failure = _log_record_is_failure(rec)
        values = (
            rec.get("timestamp", ""),
            rec.get("event", ""),
            "FAILURE" if is_failure else "SUCCESS",
            rec.get("patient_id", ""),
            rec.get("patient_name", ""),
            rec.get("destination_name", ""),
            rec.get("error_message") or rec.get("details", "") or "",
        )
        log_struct_tree.insert("", "end", iid=iid, values=values,
                                tags=("failure",) if is_failure else ("success",))
        log_struct_records_by_iid[iid] = rec

    log_struct_tree.tag_configure("failure", foreground=THEME_DANGER)
    log_struct_tree.tag_configure("success", foreground=THEME_TEXT)
    if "log_match_count_lbl" in globals():
        log_match_count_lbl.configure(text=f"{len(records)} record(s)")
    ensure_tree_empty_state(
        log_struct_tree,
        "No log entries match the current filters.\nTry a wider date range or clearing Search/Severity/Destination.",
        icon="",
    )

    log_detail_box.configure(state="normal")
    log_detail_box.delete("1.0", "end")
    log_detail_box.configure(state="disabled")
    log_open_related_btn.configure(state="disabled")


def _on_log_struct_tree_select(_event=None):
    sel = log_struct_tree.selection()
    log_detail_box.configure(state="normal")
    log_detail_box.delete("1.0", "end")
    if not sel:
        log_detail_box.configure(state="disabled")
        log_open_related_btn.configure(state="disabled")
        return
    rec = log_struct_records_by_iid.get(sel[0])
    if rec is None:
        log_detail_box.configure(state="disabled")
        log_open_related_btn.configure(state="disabled")
        return
    try:
        pretty = json.dumps(rec, indent=2, ensure_ascii=False)
    except Exception:
        pretty = str(rec)
    log_detail_box.insert("end", pretty)
    log_detail_box.configure(state="disabled")
    log_open_related_btn.configure(state="normal" if rec.get("patient_id") else "disabled")


def do_open_related_log_object():
    sel = log_struct_tree.selection()
    if not sel:
        return
    rec = log_struct_records_by_iid.get(sel[0])
    if not rec or not rec.get("patient_id"):
        return
    ok, msg = open_patient_folder(rec["patient_id"])
    if not ok:
        modern_showwarning("Open Related Object", msg)


def _refresh_current_log_view():
    """Dispatches to whichever view (Raw Text / Structured) is currently
    selected -- called from every place that used to call
    refresh_log_view() directly (filter changes, auto-tail, tab open)."""
    if log_display_mode_var.get() == "Structured":
        refresh_structured_log_view()
    else:
        refresh_log_view()


def _on_log_display_mode_change(*_args):
    if log_display_mode_var.get() == "Structured":
        log_raw_container.pack_forget()
        _refresh_log_struct_destination_choices()
        log_structured_container.pack(fill="both", expand=True)
    else:
        log_structured_container.pack_forget()
        log_raw_container.pack(fill="both", expand=True)
    _refresh_current_log_view()


def refresh_archive_list():
    """Repopulates the archive dropdown with every rotated log currently
    in LOG_ARCHIVE_DIR, newest first."""
    if not admin_tabs_active["value"]:
        return
    archives = list_log_archives()
    if not archives:
        log_archive_menu.configure(values=["(no archives yet)"])
        log_archive_var.set("(no archives yet)")
        return
    labels_to_paths = {}
    labels = []
    for fname, fpath, size, created in archives:
        created_str = created.strftime("%Y-%m-%d %H:%M") if created else "?"
        label = f"{created_str}  —  {fname}  ({size/1024:.0f} KB)"
        labels.append(label)
        labels_to_paths[label] = fpath
    log_archive_menu.configure(values=labels)
    # Store the mapping so refresh_log_view can resolve label -> path
    log_text_ref["archive_label_to_path"] = labels_to_paths
    if log_archive_var.get() not in labels:
        log_archive_var.set(labels[0])


def _resolve_selected_archive_path():
    label = log_archive_var.get()
    return log_text_ref.get("archive_label_to_path", {}).get(label, "")


def do_archive_logs_now():
    """Manual 'Archive Now' -- force-rotates every log file immediately,
    regardless of its age/size, into a fresh timestamped ZIP. Nothing is
    ever deleted; the live log is simply reset to empty afterward."""
    try:
        rotate_logs_if_needed(force=True)
        for log_path, jsonl_path in zip(ALL_LOG_FILES, ALL_LOG_FILES_JSONL):
            if os.path.exists(log_path) and os.path.getsize(log_path) > 0:
                archive_log_file(log_path, jsonl_path, reason="manual")
        refresh_archive_list()
        refresh_log_view()
        log_status_lbl.configure(text="Logs archived.", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Manual log archive failed")
        log_status_lbl.configure(text=f"Archive failed: {e}", text_color=THEME_DANGER)


def do_export_logs():
    """Lets the user save a single ZIP containing every live + archived
    log. This is the only way logs leave the app besides on-screen
    viewing -- there is no delete path."""
    dest_path = filedialog.asksaveasfilename(
        title="Export Logs",
        defaultextension=".zip",
        filetypes=[("ZIP archive", "*.zip")],
        initialfile=f"rapps_logs_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
    )
    if not dest_path:
        return
    ok = export_logs_zip(dest_path, include_archives=True)
    if ok:
        log_status_lbl.configure(text=f"Exported to {os.path.basename(dest_path)}", text_color=THEME_SUCCESS)
    else:
        log_status_lbl.configure(text="Export failed — see app.log", text_color=THEME_DANGER)


def do_change_log_retention():
    try:
        days = int(log_retention_var.get())
        save_log_retention_days(days)
        log_status_lbl.configure(text=f"Retention set to {days} days", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Failed to change log retention")
        log_status_lbl.configure(text=f"Retention change failed: {e}", text_color=THEME_DANGER)


def do_save_bandwidth_config():
    preset = bw_preset_var.get()
    try:
        custom_mbps = float(bw_custom_entry.get() or 0)
    except ValueError:
        bw_status_lbl.configure(text="Custom value must be a number.", text_color=THEME_DANGER)
        return
    if preset == "Custom" and custom_mbps <= 0:
        bw_status_lbl.configure(text="Enter a custom Mbps value greater than 0.", text_color=THEME_DANGER)
        return
    try:
        cfg = {"preset": preset, "custom_mbps": custom_mbps}
        save_bandwidth_config(cfg)
        effective = get_effective_bandwidth_mbps(cfg)
        label = "Unlimited" if effective <= 0 else f"{effective:g} Mbps"
        bw_active_limit_lbl.configure(text=label)
        bw_status_lbl.configure(text=f"Saved. Effective limit: {label}", text_color=THEME_SUCCESS)
    except Exception as e:
        bw_status_lbl.configure(text=f"Save failed: {e}", text_color=THEME_DANGER)


def refresh_destinations_ui():
    """Rebuild the destination list display and sync all option menus
    that reference destinations (pusher, routing). push_dest_menu lives on
    the Pusher tab, which exists in BOTH modes, so it's always refreshed;
    the Destinations/Routing-tab widgets only exist in Admin mode."""
    dests = load_destinations()
    names = [d["name"] for d in dests]
    opts = names if names else ["(none configured)"]

    push_dest_menu.configure(values=opts)
    if names and push_dest_var.get() not in names:
        push_dest_var.set(names[0])
    _refresh_multi_dest_checkboxes()

    if not admin_tabs_active["value"]:
        return

    display = "\n".join(
        f"{'[default] ' if d.get('default') else '           '}"
        f"{'[docs] ' if d.get('doc_transfer_enabled') else '       '}{d['name']:20s}  "
        f"{d.get('calling_ae') or DEFAULT_PUSH_CALLING_AE} → {d['ae']}@{d['ip']}:{d['port']}"
        for d in dests
    )
    dest_listbox.configure(state="normal")
    dest_listbox.delete("1.0", "end")
    dest_listbox.insert("end", display or "(no destinations saved)")
    dest_listbox.configure(state="disabled")

    dest_optmenu.configure(values=opts)
    routing_dest_menu.configure(values=opts)
    if names and routing_dest_var.get() not in names:
        routing_dest_var.set(names[0])
    if names and dest_select_var.get() not in names:
        dest_select_var.set(names[0])


def refresh_routing_ui():
    if not admin_tabs_active["value"]:
        return
    routing_tree.delete(*routing_tree.get_children())
    for rule in load_routing_rules():
        routing_tree.insert("", "end", values=(
            rule.get("modality", ""), rule.get("institution", ""),
            rule.get("source_ae", ""), rule.get("destination", ""),
        ))


_sop_filter_cache = {"classes": None, "ts": None}


def _reset_sop_filter():
    """Called whenever the underlying text is replaced from outside the
    filter itself (reload from file, save) so a stale cached 'full text'
    can't silently overwrite the fresh content next time the filter runs."""
    _sop_filter_cache["classes"] = None
    _sop_filter_cache["ts"] = None
    if "sop_search_var" in globals():
        sop_search_var.set("")


def _apply_sop_filter(*_a):
    query = sop_search_var.get().strip().lower()
    for box, key in ((sop_classes_box, "classes"), (sop_ts_box, "ts")):
        if query:
            if _sop_filter_cache[key] is None:
                _sop_filter_cache[key] = box.get("1.0", "end-1c")
            full_text = _sop_filter_cache[key]
            matches = [ln for ln in full_text.split("\n") if query in ln.lower()]
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.insert("end", "\n".join(matches) if matches else "(no matches)")
            box.configure(state="disabled")
        elif _sop_filter_cache[key] is not None:
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.insert("end", _sop_filter_cache[key])
            _sop_filter_cache[key] = None


def load_sop_into_editor():
    if not admin_tabs_active["value"]:
        return
    _reset_sop_filter()
    sop_classes, transfer_syntaxes = read_sop_ini_raw()
    sop_classes_box.configure(state="normal")
    sop_classes_box.delete("1.0", "end")
    sop_classes_box.insert("end", "\n".join(f"{k} = {v}" for k, v in sop_classes.items()))
    sop_ts_box.configure(state="normal")
    sop_ts_box.delete("1.0", "end")
    sop_ts_box.insert("end", "\n".join(f"{k} = {v}" for k, v in transfer_syntaxes.items()))


def refresh_receiver_badge():
    """User-mode Receiver tab badge: 'X reports opened today', built
    entirely from timestamps already tracked in patient_data."""
    today = datetime.date.today().isoformat()
    reports_today = 0
    with data_lock:
        for d in patient_data.values():
            last_opened = d.get("report_last_opened", "")
            if isinstance(last_opened, str) and last_opened.startswith(today):
                reports_today += 1
    try:
        rec_badge_label.configure(text=f"{reports_today} reports opened today")
    except Exception:
        pass


_toasted_stale_pids = set()  # tracks which stale patients we've already alerted on, this session


def refresh_push_badge():
    """Pusher tab badge: live failed / pending / stale counts, so problems
    are visible without having to scan the whole worklist or dig into the
    Admin Dashboard. Also fires a one-time toast per patient the first
    time it's noticed sitting stale (pending too long with no activity),
    so a backlog doesn't go unnoticed."""
    failed = pending = stale = 0
    newly_stale = []
    with data_lock:
        for pid, d in patient_data.items():
            status = d.get("status", "")
            if status == STATUS_FAILED:
                failed += 1
            if status in (STATUS_PENDING, STATUS_RETRYING, STATUS_QUEUED):
                pending += 1
            if is_pending_stale(d):
                stale += 1
                if pid not in _toasted_stale_pids:
                    newly_stale.append(pid)
                    _toasted_stale_pids.add(pid)
    try:
        color = THEME_DANGER if failed else THEME_TEXT_MUTED
        push_badge_label.configure(
            text=f"{failed} failed  •  {pending} pending  •  {stale} stale", text_color=color)
    except Exception:
        pass

    if newly_stale:
        if len(newly_stale) == 1:
            msg = f"Patient {newly_stale[0]} has been pending with no activity for a while."
        else:
            msg = f"{len(newly_stale)} patients have been pending with no activity for a while."
        ui_event_queue.put(("toast", ("Stale Pending Items", msg)))


def select_all_by_status(tree, statuses):
    """Part 4 quick-filter: select every row whose status is in `statuses`,
    so the person doesn't have to hand multi-select rows before a bulk
    retry."""
    tree.selection_remove(*tree.selection())
    to_select = []
    with data_lock:
        for pid, d in patient_data.items():
            if d.get("status") in statuses and tree.exists(pid):
                to_select.append(pid)
    if to_select:
        tree.selection_set(to_select)
        tree.see(to_select[0])


def jump_to_receiver_stale_filtered():
    """Admin Dashboard 'Jump to Receiver, filtered' button: switch to the
    Receiver tab and select the stale-pending rows there."""
    tabview.set("Receiver")
    stale_ids = []
    with data_lock:
        for pid, d in patient_data.items():
            if is_pending_stale(d):
                stale_ids.append(pid)
    rec_tree.selection_remove(*rec_tree.selection())
    existing = [pid for pid in stale_ids if rec_tree.exists(pid)]
    if existing:
        rec_tree.selection_set(existing)
        rec_tree.see(existing[0])


def _sync_tree_rows(tree, rows, empty_message=None, empty_icon=""):
    """Updates a Treeview to match `rows` (list of (iid, values_tuple))
    while only touching rows that actually changed -- no full
    delete()+reinsert() every call. Used by the dashboard's small
    sub-trees, which otherwise got rebuilt from scratch on every 2s
    refresh tick even when nothing in them had changed."""
    wanted_ids = [iid for iid, _ in rows]
    wanted_set = set(wanted_ids)
    existing = list(tree.get_children(""))

    # Drop rows that no longer belong.
    for iid in existing:
        if iid not in wanted_set:
            tree.delete(iid)

    # Insert/update in the desired order.
    for index, (iid, values) in enumerate(rows):
        if tree.exists(iid):
            if tree.item(iid, "values") != values:
                tree.item(iid, values=values)
            if tree.index(iid) != index:
                tree.move(iid, "", index)
        else:
            tree.insert("", index, iid=iid, values=values)

    if empty_message:
        ensure_tree_empty_state(tree, empty_message, icon=empty_icon)


QUEUE_GROWING_THRESHOLD = 50
_queue_growing_notify_throttle = [0.0]

_dest_health_check_inflight = {"value": False}
_last_dest_health_check_at = [0.0]


DESTINATION_UPTIME_HISTORY_LEN = 200

# =========================================================
# PACS HEALTH MONITOR (continuous background C-ECHO scheduling)
# =========================================================

PACS_HEALTH_CHECK_INTERVAL_SEC = 60
_pacs_health_thread_started = {"value": False}


EXPORT_SCOPE_LABEL_TO_KEY = {
    "Entire Patient": "patient", "Study": "study", "Series": "series",
    "Selected Files": "files", "Entire Worklist": "worklist",
}


def _update_export_scope_visibility():
    if not admin_tabs_active["value"]:
        return
    try:
        scope_key = EXPORT_SCOPE_LABEL_TO_KEY.get(export_scope_var.get(), "patient")
    except Exception:
        return

    show_pid = scope_key != "worklist"
    show_study = scope_key == "study"
    show_series = scope_key == "series"
    show_files = scope_key == "files"

    try:
        if show_pid:
            export_pid_menu.grid()
        else:
            export_pid_menu.grid_remove()
        if show_study:
            export_study_menu.grid()
        else:
            export_study_menu.grid_remove()
        if show_series:
            export_series_menu.grid()
        else:
            export_series_menu.grid_remove()
        if show_files:
            export_files_frame.pack(fill="x", pady=(0, 8), after=export_patient_row)
        else:
            export_files_frame.pack_forget()
    except Exception:
        log_exception("Failed to update Export tab scope visibility")

    if show_pid:
        refresh_export_patient_options()
    refresh_export_study_series_options()


def refresh_export_patient_options():
    if not admin_tabs_active["value"]:
        return
    with data_lock:
        pids = sorted(patient_data.keys())
    values = pids if pids else ["(no patients)"]
    export_pid_menu.configure(values=values)
    if export_pid_var.get() not in values:
        export_pid_var.set(values[0])


def refresh_export_study_series_options():
    """Populates the Study/Series dropdowns and the Selected-Files tree
    for whichever patient is currently chosen. Cheap-ish (reads DICOM
    headers only, no pixel data) but only actually needed for
    study/series/files scopes, so it's only called when relevant."""
    if not admin_tabs_active["value"]:
        return
    scope_key = EXPORT_SCOPE_LABEL_TO_KEY.get(export_scope_var.get(), "patient")
    pid = export_pid_var.get()
    if scope_key not in ("study", "series", "files") or not pid or pid == "(no patients)":
        return

    index = build_file_index_for_patient(pid)

    if scope_key == "study":
        studies = sorted({e["study_uid"] for e in index if e["study_uid"]})
        values = studies if studies else ["(no studies found)"]
        export_study_menu.configure(values=values)
        if export_study_var.get() not in values:
            export_study_var.set(values[0])

    if scope_key == "series":
        series_labels = sorted({
            f"{e['series_uid']}  ({e['modality']}{', ' + e['series_description'] if e['series_description'] else ''})"
            for e in index if e["series_uid"]
        })
        values = series_labels if series_labels else ["(no series found)"]
        export_series_menu.configure(values=values)
        if export_series_var.get() not in values:
            export_series_var.set(values[0])

    if scope_key == "files":
        export_files_tree.delete(*export_files_tree.get_children())
        for e in index:
            size_kb = f"{e['size'] / 1024:.1f} KB"
            export_files_tree.insert("", "end", iid=e["path"],
                                     values=(e["sop_uid"], e["series_uid"], e["modality"], size_kb))


CONFIG_BUNDLE_KEYS = (
    "destinations", "routing_rules", "ldap_config", "bandwidth_config",
    "tls_config", "notifications_config", "backup_schedule", "log_retention_days",
    "app_settings",
)


def do_export_config_bundle():
    """9.1 -- gathers the output of every load_*() config function listed
    in CONFIG_BUNDLE_KEYS into one JSON bundle. Because that includes the
    LDAP service-account password and SMTP credentials, the bundle is
    encrypted with this installation's own key.key (same Fernet key
    everything else already uses) rather than written as plain JSON."""
    if not modern_confirm(
        "Export App Configuration",
        "This bundle includes LDAP and SMTP credentials. It will be encrypted with "
        "this installation's key.key, so it can only be imported using that same "
        "key.key (e.g. on this machine, or a copy of it moved together with key.key).\n\nContinue?",
    ):
        return
    bundle = {
        "bundle_version": 1,
        "exported_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "destinations": load_destinations(),
        "routing_rules": load_routing_rules(),
        "ldap_config": load_ldap_config(),
        "bandwidth_config": load_bandwidth_config(),
        "tls_config": load_tls_config(),
        "notifications_config": load_notifications_config(),
        "backup_schedule": load_backup_schedule(),
        "log_retention_days": load_log_retention_days(),
        "app_settings": load_app_settings(),
    }
    default_name = f"rapps_config_bundle_{datetime.date.today().isoformat()}.enc"
    out_path = filedialog.asksaveasfilename(
        title="Export App Configuration", defaultextension=".enc",
        initialfile=default_name, filetypes=[("Encrypted config bundle", "*.enc")])
    if not out_path:
        return
    try:
        fernet = Fernet(load_key())
        with open(out_path, "wb") as f:
            f.write(fernet.encrypt(json.dumps(bundle, indent=2).encode()))
        write_audit_log("CONFIG-BUNDLE-EXPORTED", f"path={out_path}")
        config_bundle_status_lbl.configure(text=f"Exported to {out_path}", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Failed to export config bundle")
        config_bundle_status_lbl.configure(text=f"Export failed: {e}", text_color=THEME_DANGER)


def do_import_config_bundle():
    """9.1 -- reverse of do_export_config_bundle(): decrypts the chosen
    bundle with this installation's key.key, confirms, then calls the
    matching save_*() for each key present, and refreshes the UI those
    settings feed."""
    in_path = filedialog.askopenfilename(
        title="Import App Configuration",
        filetypes=[("Encrypted config bundle", "*.enc"), ("All files", "*.*")])
    if not in_path:
        return
    try:
        fernet = Fernet(load_key())
        with open(in_path, "rb") as f:
            raw = fernet.decrypt(f.read())
        bundle = json.loads(raw.decode())
    except Exception as e:
        log_exception("Failed to read/decrypt config bundle")
        config_bundle_status_lbl.configure(
            text="Could not read this bundle -- wrong key.key, or not a bundle file.", text_color=THEME_DANGER)
        return

    exported_at = bundle.get("exported_at", "an unknown date")
    if not modern_confirm(
        "Import App Configuration",
        f"This will OVERWRITE current Destinations, Routing Rules, LDAP, Bandwidth, TLS, Notifications, "
        f"Backup Schedule, Log Retention, and App Settings with the values from this bundle "
        f"(exported {exported_at}).\n\nThis cannot be undone. Continue?",
        danger=True,
    ):
        return

    try:
        if "destinations" in bundle:
            save_destinations(bundle["destinations"])
        if "routing_rules" in bundle:
            save_routing_rules(bundle["routing_rules"])
        if "ldap_config" in bundle:
            save_ldap_config(bundle["ldap_config"])
        if "bandwidth_config" in bundle:
            save_bandwidth_config(bundle["bandwidth_config"])
        if "tls_config" in bundle:
            save_tls_config(bundle["tls_config"])
        if "notifications_config" in bundle:
            save_notifications_config(bundle["notifications_config"])
        if "backup_schedule" in bundle:
            save_backup_schedule(bundle["backup_schedule"])
        if "log_retention_days" in bundle:
            save_log_retention_days(bundle["log_retention_days"])
        if "app_settings" in bundle:
            save_app_settings(bundle["app_settings"])
    except Exception as e:
        log_exception("Failed to apply imported config bundle")
        config_bundle_status_lbl.configure(text=f"Bundle was read, but applying it failed: {e}", text_color=THEME_DANGER)
        return

    write_audit_log("CONFIG-BUNDLE-IMPORTED", f"path={in_path}")

    try:
        refresh_destinations_ui()
        refresh_routing_ui()
        if "backup_next_run_lbl" in globals():
            refresh_next_backup_indicator()
        if "app_settings" in bundle:
            _app_settings_cache["value"] = None  # force reload from disk
            global APP_SETTINGS
            APP_SETTINGS = load_app_settings()
            _populate_settings_controls_from_app_settings()
    except Exception:
        log_exception("Failed to refresh UI after config bundle import")

    config_bundle_status_lbl.configure(
        text="Configuration imported. Reopen the LDAP/Notifications/Bandwidth/TLS tabs (or restart the app) "
             "to see their reloaded values.",
        text_color=THEME_SUCCESS)


def do_reset_export_defaults():
    """8.1 -- restores the Export tab's controls (and the persisted
    template) to DEFAULT_EXPORT_TEMPLATE. Leaves the password field alone
    since it's never part of the persisted template to begin with."""
    save_export_template(dict(DEFAULT_EXPORT_TEMPLATE))
    export_scope_var.set(DEFAULT_EXPORT_TEMPLATE["scope"])
    export_reports_var.set(DEFAULT_EXPORT_TEMPLATE["include_reports"])
    export_logs_var.set(DEFAULT_EXPORT_TEMPLATE["include_logs"])
    export_metadata_var.set(DEFAULT_EXPORT_TEMPLATE["include_metadata"])
    export_dicomdir_var.set(DEFAULT_EXPORT_TEMPLATE["include_dicomdir"])
    export_status_lbl.configure(text="Export options reset to defaults.", text_color=THEME_TEXT_MUTED)


def do_start_export():
    scope_key = EXPORT_SCOPE_LABEL_TO_KEY.get(export_scope_var.get(), "patient")
    options = {
        "scope": scope_key,
        "include_reports": export_reports_var.get(),
        "include_logs": export_logs_var.get(),
        "include_metadata": export_metadata_var.get(),
        "include_dicomdir": export_dicomdir_var.get(),
        "password": export_password_var.get().strip(),
    }

    if scope_key in ("patient", "study", "series", "files"):
        pid = export_pid_var.get()
        if not pid or pid == "(no patients)":
            export_status_lbl.configure(text="Choose a patient first.", text_color=THEME_DANGER)
            return
        options["pid"] = pid

    if scope_key == "study":
        study_uid = export_study_var.get()
        if not study_uid or study_uid.startswith("("):
            export_status_lbl.configure(text="Choose a study first.", text_color=THEME_DANGER)
            return
        options["study_uid"] = study_uid

    if scope_key == "series":
        series_label = export_series_var.get()
        if not series_label or series_label.startswith("("):
            export_status_lbl.configure(text="Choose a series first.", text_color=THEME_DANGER)
            return
        options["series_uid"] = series_label.split("  (")[0]

    if scope_key == "files":
        selected = export_files_tree.selection()
        if not selected:
            export_status_lbl.configure(text="Select at least one file (Ctrl/Shift-click).", text_color=THEME_DANGER)
            return
        options["file_paths"] = list(selected)

    if options["password"] and not PYZIPPER_AVAILABLE:
        export_status_lbl.configure(text="Password-protected export needs 'pyzipper' installed.", text_color=THEME_DANGER)
        return

    save_export_template({
        "scope": export_scope_var.get(),
        "include_reports": options["include_reports"],
        "include_logs": options["include_logs"],
        "include_metadata": options["include_metadata"],
        "include_dicomdir": options["include_dicomdir"],
        "had_password": bool(options["password"]),
    })

    default_name = f"rapps_export_{scope_key}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    dest_path = filedialog.asksaveasfilename(
        title="Export To…", defaultextension=".zip",
        filetypes=[("ZIP archive", "*.zip")], initialfile=default_name,
    )
    if not dest_path:
        return
    options["dest_path"] = dest_path

    export_start_btn.configure(state="disabled")
    export_progress_bar.set(0)
    export_status_lbl.configure(text="Starting export…", text_color=THEME_TEXT_MUTED)

    def progress_cb(done, total, message):
        def on_ui():
            export_progress_bar.set(done / total if total else 0)
            export_status_lbl.configure(text=message, text_color=THEME_TEXT_MUTED)
        app.after(0, on_ui)

    def run():
        try:
            file_count, pids = run_export_job(options, progress_cb=progress_cb)
            def on_done():
                export_progress_bar.set(1.0)
                export_status_lbl.configure(
                    text=f"Exported {file_count} file(s) across {len(pids)} patient(s) to {os.path.basename(dest_path)}",
                    text_color=THEME_SUCCESS)
                export_start_btn.configure(state="normal")
                notify_event("push_complete", "Export Complete",
                            f"{file_count} file(s) exported to {os.path.basename(dest_path)}")
            app.after(0, on_done)
        except Exception as e:
            log_exception("Export job failed")
            err_msg = str(e)
            def on_fail():
                export_status_lbl.configure(text=f"Export failed: {err_msg}", text_color=THEME_DANGER)
                export_start_btn.configure(state="normal")
            app.after(0, on_fail)

    threading.Thread(target=run, daemon=True).start()


def _resolve_report_date_range():
    """Turns the Reports tab's selection into (start_dt, end_dt).
    Returns (start_dt, end_dt, label) or (None, None, error_message)."""
    report_type = reports_type_var.get()
    now = datetime.datetime.now()
    if report_type == "Daily":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        return start, now, "Daily"
    if report_type == "Weekly":
        return now - datetime.timedelta(days=7), now, "Weekly"
    if report_type == "Monthly":
        return now - datetime.timedelta(days=30), now, "Monthly"
    if report_type == "Year to Date":
        start = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
        return start, now, "Year to Date"
    # Custom Date Range
    try:
        from_str = reports_from_entry.get().strip() or datetime.date.today().isoformat()
        to_str = reports_to_entry.get().strip() or datetime.date.today().isoformat()
        start = datetime.datetime.strptime(from_str, "%Y-%m-%d")
        end = datetime.datetime.strptime(to_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
        if end < start:
            return None, None, "'To' date must be on or after 'From' date."
        return start, end, "Custom Date Range"
    except ValueError:
        return None, None, "Dates must be in YYYY-MM-DD format."


def _format_file_size(num_bytes):
    for unit in ("B", "KB", "MB", "GB"):
        if num_bytes < 1024:
            return f"{num_bytes:.0f} {unit}" if unit == "B" else f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} TB"


def _refresh_reports_meta_display():
    """Metadata display for the last generated report -- file name, size,
    date range covered, and when it was generated. Purely a presentation
    layer; does not call into generate_report_pdf/compute_report_stats
    (the report engine itself is untouched)."""
    path = reports_last_path["value"]
    if not path or not os.path.isfile(path):
        reports_meta_lbl.configure(text="No report generated yet this session.")
        return
    try:
        size = _format_file_size(os.path.getsize(path))
    except Exception:
        size = "unknown size"
    label = reports_last_meta.get("label") or "—"
    start_dt = reports_last_meta.get("start_dt")
    end_dt = reports_last_meta.get("end_dt")
    generated_at = reports_last_meta.get("generated_at")
    date_range_txt = f"{start_dt:%Y-%m-%d %H:%M} → {end_dt:%Y-%m-%d %H:%M}" if start_dt and end_dt else "—"
    generated_txt = generated_at.strftime("%Y-%m-%d %H:%M:%S") if generated_at else "—"
    reports_meta_lbl.configure(text=(
        f"File:          {os.path.basename(path)}  ({size})\n"
        f"Report type:   {label}\n"
        f"Data covers:   {date_range_txt}\n"
        f"Generated at:  {generated_txt}"
    ))


def do_preview_last_report():
    path = reports_last_path["value"]
    if not path or not os.path.isfile(path):
        reports_status_lbl.configure(text="Generate a report first.", text_color=THEME_DANGER)
        return
    ok, msg = open_document(path)  # reuses the same OS-launch utility as Open Report/Open History
    if not ok:
        reports_status_lbl.configure(text=f"Preview failed: {msg}", text_color=THEME_DANGER)


def do_generate_report():
    if not REPORTLAB_AVAILABLE:
        reports_status_lbl.configure(text="PDF generation needs 'reportlab' and 'matplotlib' installed.",
                                     text_color=THEME_DANGER)
        return
    start_dt, end_dt, label = _resolve_report_date_range()
    if start_dt is None:
        reports_status_lbl.configure(text=label, text_color=THEME_DANGER)
        return

    default_name = f"rapps_report_{label.lower().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    dest_path = filedialog.asksaveasfilename(
        title="Save Report As…", defaultextension=".pdf",
        filetypes=[("PDF document", "*.pdf")], initialfile=default_name,
    )
    if not dest_path:
        return

    reports_generate_btn.configure(state="disabled")
    reports_status_lbl.configure(text="Generating report…", text_color=THEME_TEXT_MUTED)

    def run():
        try:
            generate_report_pdf(start_dt, end_dt, dest_path, label)
            def on_done():
                reports_last_path["value"] = dest_path
                reports_last_meta["label"] = label
                reports_last_meta["start_dt"] = start_dt
                reports_last_meta["end_dt"] = end_dt
                reports_last_meta["generated_at"] = datetime.datetime.now()
                reports_status_lbl.configure(text=f"Saved {os.path.basename(dest_path)}", text_color=THEME_SUCCESS)
                reports_generate_btn.configure(state="normal")
                reports_preview_btn.configure(state="normal")
                reports_print_btn.configure(state="normal")
                reports_email_btn.configure(state="normal")
                _refresh_reports_meta_display()
                notify_event("push_complete", "Report Generated", f"{label} report saved to {os.path.basename(dest_path)}")
            app.after(0, on_done)
        except Exception as e:
            log_exception("Report generation failed")
            err_msg = str(e)
            def on_fail():
                reports_status_lbl.configure(text=f"Report generation failed: {err_msg}", text_color=THEME_DANGER)
                reports_generate_btn.configure(state="normal")
            app.after(0, on_fail)

    threading.Thread(target=run, daemon=True).start()


def do_print_last_report():
    path = reports_last_path["value"]
    if not path or not os.path.isfile(path):
        reports_status_lbl.configure(text="Generate a report first.", text_color=THEME_DANGER)
        return

    def run():
        ok, msg = print_pdf_file(path)
        def on_ui():
            reports_status_lbl.configure(
                text=("Sent: " if ok else "Print failed: ") + msg,
                text_color=THEME_SUCCESS if ok else THEME_DANGER,
            )
        app.after(0, on_ui)

    threading.Thread(target=run, daemon=True).start()


def do_email_last_report():
    path = reports_last_path["value"]
    if not path or not os.path.isfile(path):
        reports_status_lbl.configure(text="Generate a report first.", text_color=THEME_DANGER)
        return

    def run():
        ok, msg = email_pdf_file(path, "PDF Report", "Please find the attached report.")
        def on_ui():
            reports_status_lbl.configure(
                text=("Sent: " if ok else "Email failed: ") + msg,
                text_color=THEME_SUCCESS if ok else THEME_DANGER,
            )
        app.after(0, on_ui)

    threading.Thread(target=run, daemon=True).start()


def refresh_destination_filter_options():
    """Keeps the Pusher tab's Destination filter dropdown in sync with
    whatever push_target values actually appear in the worklist right
    now -- dynamic, since destinations are user-configurable and a
    patient's push_target is only known once something has actually been
    pushed."""
    try:
        with data_lock:
            targets = sorted({d.get("push_target", "") or "(none)" for d in patient_data.values()})
        values = ["All"] + targets
        current = push_dest_filter_var.get()
        push_dest_filter_menu.configure(values=values)
        if current not in values:
            push_dest_filter_var.set("All")
    except Exception:
        pass  # widgets not built yet


def _fmt_ts_or_dash(ts):
    if not ts:
        return "—"
    return datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def check_all_destinations_health():
    """Runs a C-ECHO against every configured destination, one at a time,
    from whatever thread calls this. Always call this off the main/UI
    thread -- it does real network I/O per destination."""
    for dest in load_destinations():
        try:
            check_destination_health(dest)
        except Exception:
            log_exception(f"PACS health check failed for destination {dest.get('name')}")


def pacs_health_monitor_loop():
    """Background daemon loop: checks every destination's health every
    PACS_HEALTH_CHECK_INTERVAL_SEC seconds, for as long as the app runs.
    Started once at startup (see start_pacs_health_monitor_thread)."""
    while not app_shutdown_event.is_set():
        try:
            check_all_destinations_health()
        except Exception:
            log_exception("PACS health monitor loop iteration failed")
        app_shutdown_event.wait(PACS_HEALTH_CHECK_INTERVAL_SEC)


def start_pacs_health_monitor_thread():
    if _pacs_health_thread_started["value"]:
        return
    _pacs_health_thread_started["value"] = True
    threading.Thread(target=pacs_health_monitor_loop, daemon=True).start()


def do_run_all_health_checks_now():
    """'Check All Now' button -- runs immediately in a background thread
    so the UI never freezes waiting on network timeouts."""
    health_status_lbl.configure(text="Checking all destinations…")

    def run():
        check_all_destinations_health()
        def on_ui():
            refresh_health_tree()
            health_status_lbl.configure(text=f"Last checked: {datetime.datetime.now().strftime('%H:%M:%S')}")
        app.after(0, on_ui)

    threading.Thread(target=run, daemon=True).start()


def _fmt_offlineq_ts(epoch_ts):
    if not epoch_ts:
        return "—"
    return datetime.datetime.fromtimestamp(epoch_ts).strftime("%Y-%m-%d %H:%M:%S")


_queue_threshold_alert_state = {"armed": True}


def _check_offline_queue_threshold(size):
    """3.2 -- fires notify_event('queue_growing', ...) once when the queue
    crosses the configured threshold, then re-arms only once it drops back
    below it, so a queue that stays large doesn't spam a notification on
    every refresh tick."""
    try:
        threshold = int(load_notifications_config()["thresholds"].get("offline_queue_size", 20))
    except (TypeError, ValueError):
        threshold = 20
    if threshold <= 0:
        return
    if size >= threshold and _queue_threshold_alert_state["armed"]:
        _queue_threshold_alert_state["armed"] = False
        notify_event("queue_growing", "Offline Queue Growing",
                     f"The offline queue has reached {size} item(s) (alert threshold: {threshold}).")
    elif size < threshold and not _queue_threshold_alert_state["armed"]:
        _queue_threshold_alert_state["armed"] = True


def refresh_offline_queue_ui():
    """Populates the Offline Queue panel at the top of the Pusher tab.
    Cheap (in-memory cache read, no network I/O) -- safe on every
    periodic_refresh tick."""
    try:
        existing_ids = set(offlineq_tree.get_children())
    except Exception:
        return  # widgets not built yet

    size, oldest, soonest_retry, soonest_interval = get_offline_queue_summary()
    offlineq_size_lbl.configure(text=str(size))
    offlineq_oldest_lbl.configure(text=oldest["queued_at"] if oldest else "—")
    offlineq_next_retry_lbl.configure(text=_fmt_offlineq_ts(soonest_retry) if soonest_retry else "—")
    offlineq_interval_lbl.configure(text=f"{soonest_interval}s" if soonest_interval else "—")
    _check_offline_queue_threshold(size)

    items = get_offline_queue_items()
    seen_ids = set()
    for item in items:
        pid = item["pid"]
        seen_ids.add(pid)
        row = (
            pid, item.get("destination_name", ""), item.get("queued_at", ""),
            item.get("attempt_count", 0), _fmt_offlineq_ts(item.get("next_retry_at")),
            (item.get("last_error", "") or "")[:80],
        )
        if pid in existing_ids:
            offlineq_tree.item(pid, values=row)
        else:
            offlineq_tree.insert("", "end", iid=pid, values=row)
    for stale_id in existing_ids - seen_ids:
        offlineq_tree.delete(stale_id)
    ensure_tree_empty_state(offlineq_tree, "Offline queue is empty.\nStudies land here only when a destination is unreachable.", icon="")


def _retry_offline_item_now(pid):
    """Runs one queued patient's push immediately in a background thread,
    regardless of its scheduled next_retry_at -- used by both the
    'Retry Due Items Now' button and double-clicking a row."""
    item = next((i for i in get_offline_queue_items() if i["pid"] == pid), None)
    if not item:
        return
    dest = get_destination_by_name(item["destination_name"])
    if not dest:
        modern_showerror("Retry Failed", f"Destination '{item['destination_name']}' no longer exists.")
        return

    def run():
        try:
            sent, total, ok = push_single_patient(pid, destination=dest)
            if ok:
                dequeue_offline(pid)
        except Exception:
            log_exception(f"Manual offline-queue retry failed for {pid}")
        app.after(0, refresh_offline_queue_ui)

    threading.Thread(target=run, daemon=True).start()


def do_retry_due_offline_items_now():
    """'Retry Due Items Now' button -- immediately attempts every queued
    item regardless of its backoff schedule (still one at a time, still
    fully backgrounded)."""
    items = get_offline_queue_items()
    if not items:
        return

    def run():
        for item in items:
            if app_shutdown_event.is_set():
                break
            dest = get_destination_by_name(item["destination_name"])
            if not dest:
                continue
            try:
                sent, total, ok = push_single_patient(item["pid"], destination=dest)
                if ok:
                    dequeue_offline(item["pid"])
            except Exception:
                log_exception(f"Bulk offline-queue retry failed for {item['pid']}")
        app.after(0, refresh_offline_queue_ui)

    threading.Thread(target=run, daemon=True).start()


def _on_offlineq_row_double_click(event):
    row_id = offlineq_tree.identify_row(event.y)
    if row_id:
        _retry_offline_item_now(row_id)


def _abandon_offline_queue_item(pid):
    """3.1 -- manually removes one item from the offline queue without
    attempting a retry. Confirms first if attempts have already been
    invested, since that work (and the checkpoint progress) is lost."""
    item = next((i for i in get_offline_queue_items() if i["pid"] == pid), None)
    if not item:
        return
    attempts = item.get("attempt_count", 0)
    if attempts > 0:
        if not modern_confirm(
            "Remove from Queue",
            f"{pid} has already had {attempts} retry attempt(s) invested.\n"
            "Remove it from the offline queue without retrying?",
            danger=True,
        ):
            return
    dequeue_offline(pid)
    write_audit_log("QUEUE-ABANDONED", f"pid={pid} attempts={attempts}")
    refresh_offline_queue_ui()


def _build_offlineq_context_menu(event):
    row_id = offlineq_tree.identify_row(event.y)
    if not row_id:
        return
    offlineq_tree.selection_set(row_id)
    import tkinter as tk
    menu = tk.Menu(app, tearoff=0)
    menu.add_command(label="Retry Now", command=lambda: _retry_offline_item_now(row_id))
    menu.add_separator()
    menu.add_command(label="Remove from Queue", command=lambda: _abandon_offline_queue_item(row_id))
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


offlineq_tree.bind("<Button-3>", _build_offlineq_context_menu)
offlineq_tree.bind("<Button-2>", _build_offlineq_context_menu)  # macOS


offlineq_tree.bind("<Double-1>", _on_offlineq_row_double_click)
offlineq_retry_all_btn.configure(command=do_retry_due_offline_items_now)


def refresh_performance_tab():
    """Populates every stat card + sparkline graph on the Performance
    Metrics tab. Only called while that tab is actually visible (see
    periodic_refresh) since it involves a log scan and several psutil
    syscalls -- not something to do on every 2s tick regardless."""
    if not admin_tabs_active["value"]:
        return
    try:
        m = compute_live_performance_metrics()
    except Exception:
        log_exception("Failed to compute live performance metrics")
        return

    perf_cards["images_per_sec"].configure(text=f"{m['images_per_sec']:.3f}")
    perf_cards["studies_per_sec"].configure(text=f"{m['studies_per_sec']:.4f}")
    perf_cards["mb_per_sec"].configure(text=f"{m['mb_per_sec']:.2f}")
    perf_cards["avg_receive_time"].configure(text=f"{m['avg_receive_time']:.2f}s")
    perf_cards["avg_push_time"].configure(text=f"{m['avg_push_time']:.2f}s")
    perf_cards["avg_association_time"].configure(text=f"{m['avg_association_time']:.2f}s")
    perf_cards["avg_queue_time"].configure(text=f"{m['avg_queue_time']:.1f}s")
    perf_cards["avg_retry_time"].configure(text=f"{m['avg_retry_time']:.0f}s")
    perf_cards["cpu_pct"].configure(text=f"{m['cpu_pct']:.0f}%" if PSUTIL_AVAILABLE else "N/A")
    perf_cards["ram_pct"].configure(text=f"{m['ram_pct']:.0f}%" if PSUTIL_AVAILABLE else "N/A")
    disk_total_kbps = (m["disk_read_bps"] + m["disk_write_bps"]) / 1024.0
    perf_cards["disk_io"].configure(
        text=f"R {m['disk_read_bps']/1024:.1f} / W {m['disk_write_bps']/1024:.1f} KB/s" if PSUTIL_AVAILABLE else "N/A")
    net_total_kbps = (m["network_up_bps"] + m["network_down_bps"]) / 1024.0
    perf_cards["network_io"].configure(
        text=f"↑{m['network_up_bps']/1024:.1f} ↓{m['network_down_bps']/1024:.1f} KB/s" if PSUTIL_AVAILABLE else "N/A")
    perf_cards["database_size"].configure(text=f"{m['database_size_bytes']/1024:.1f} KB")
    perf_cards["uptime"].configure(text=_fmt_uptime(m["uptime_sec"]))
    perf_cards["docs_per_sec"].configure(text=f"{m['docs_per_sec']:.3f}")
    perf_cards["doc_success_rate_pct"].configure(text=f"{m['doc_success_rate_pct']:.1f}%")
    perf_cards["avg_doc_transfer_time"].configure(text=f"{m['avg_doc_transfer_time']:.2f}s")
    _check_doc_transfer_failure_rate_alert()

    perf_history["images_per_sec"].append(m["images_per_sec"])
    perf_history["studies_per_sec"].append(m["studies_per_sec"])
    perf_history["cpu_pct"].append(m["cpu_pct"])
    perf_history["ram_pct"].append(m["ram_pct"])
    perf_history["disk_io_kbps"].append(disk_total_kbps)
    perf_history["network_kbps"].append(net_total_kbps)

    for key, (canvas, color) in perf_graph_canvases.items():
        _draw_sparkline(canvas, list(perf_history[key]), color)


def do_export_performance_metrics():
    dest_path = filedialog.asksaveasfilename(
        title="Export Performance Metrics", defaultextension=".json",
        filetypes=[("JSON", "*.json")],
        initialfile=f"rapps_metrics_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
    )
    if not dest_path:
        return
    try:
        export_performance_metrics(dest_path)
        perf_status_lbl.configure(text=f"Exported to {os.path.basename(dest_path)}", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Failed to export performance metrics")
        perf_status_lbl.configure(text=f"Export failed: {e}", text_color=THEME_DANGER)


def _build_ldap_roster_context_menu(event):
    row_id = ldap_roster_tree.identify_row(event.y)
    if not row_id:
        return
    ldap_roster_tree.selection_set(row_id)
    users = load_ldap_users()
    is_revoked = bool(users.get(row_id, {}).get("revoked"))
    import tkinter as tk
    menu = tk.Menu(app, tearoff=0)
    if is_revoked:
        menu.add_command(label="Restore Access", command=lambda: (_set_ldap_user_revoked(row_id, False), refresh_ldap_roster_ui()))
    else:
        menu.add_command(label="Revoke Access", command=lambda: (
            modern_confirm("Revoke Access", f"Revoke login access for {row_id}?\nThey will be blocked even if their directory credentials are still valid.", danger=True)
            and (_set_ldap_user_revoked(row_id, True), refresh_ldap_roster_ui())))
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


def refresh_ldap_roster_ui():
    """Read-only refresh of the imported-users roster tree. Must never have
    side effects on disk -- this is called from the periodic UI refresh loop,
    so anything that writes config here would silently stomp saved LDAP
    settings on every tick."""
    if not admin_tabs_active["value"]:
        return
    try:
        existing_ids = set(ldap_roster_tree.get_children())
    except Exception:
        return
    users = load_ldap_users()
    seen_ids = set()
    for username, u in users.items():
        seen_ids.add(username)
        revoked = bool(u.get("revoked"))
        row = (u.get("username", ""), u.get("display_name", ""), u.get("email", ""),
              u.get("role", ""), u.get("last_login", ""), "Revoked" if revoked else "Active")
        tags = ("access_revoked",) if revoked else ()
        if username in existing_ids:
            ldap_roster_tree.item(username, values=row, tags=tags)
        else:
            ldap_roster_tree.insert("", "end", iid=username, values=row, tags=tags)
    for stale_id in existing_ids - seen_ids:
        ldap_roster_tree.delete(stale_id)
    ensure_tree_empty_state(ldap_roster_tree, "No LDAP users imported yet.\nUse 'Import Users' above once LDAP is configured.", icon="")


def do_save_ldap_config():
    """Handler for the 'Save LDAP Settings' button. Gathers the LDAP form
    fields, persists them, and refreshes the roster tree. Previously this
    logic lived (unreachable, under the wrong name) inside
    refresh_ldap_roster_ui and ran on every periodic refresh tick instead of
    only on demand; the Save button itself called a function
    (`do_save_ldap_config`) that didn't exist anywhere, which raised a
    NameError the instant an admin unlocked Admin mode -- aborting
    set_view() before it could reach apply_receiver_mode_visibility() or
    update_admin_bar_ui(). That's why the Lock button, admin-only controls,
    and the 'Mode: Administrator' label never updated after login."""
    group_mappings = {}
    for role, entry in ldap_group_map_entries.items():
        cn = entry.get().strip()
        if cn:
            group_mappings[cn] = role

    cfg = load_ldap_config()
    cfg.update({
        "enabled": ldap_enabled_var.get(),
        "server_uri": ldap_server_entry.get().strip(),
        "use_ssl": ldap_ssl_var.get(),
        "domain": ldap_domain_entry.get().strip(),
        "user_bind_dn_template": ldap_bind_template_entry.get().strip(),
        "bind_dn": ldap_bind_dn_entry.get().strip(),
        "bind_password": ldap_bind_pw_entry.get(),
        "user_search_base": ldap_user_base_entry.get().strip(),
        "user_search_filter": ldap_user_filter_entry.get().strip() or "(sAMAccountName={username})",
        "group_search_base": ldap_group_base_entry.get().strip(),
        "group_mappings": group_mappings,
    })
    try:
        save_ldap_config(cfg)
        write_audit_log("LDAP-CONFIG-SAVED", "LDAP settings updated by administrator")
        ldap_status_lbl.configure(text="LDAP settings saved.", text_color=THEME_SUCCESS)
        refresh_ldap_roster_ui()
    except Exception as e:
        log_exception("Failed to save LDAP config")
        ldap_status_lbl.configure(text=f"Save failed: {e}", text_color=THEME_DANGER)


def do_ldap_import_users_now():
    if not LDAP3_AVAILABLE:
        ldap_status_lbl.configure(text="LDAP support needs 'ldap3' installed.", text_color=THEME_DANGER)
        return
    ldap_import_btn.configure(state="disabled")
    ldap_status_lbl.configure(text="Importing users from directory…", text_color=THEME_TEXT_MUTED)

    def run():
        count, err = ldap_import_all_users()
        def on_done():
            ldap_import_btn.configure(state="normal")
            if err:
                ldap_status_lbl.configure(text=f"Import failed: {err}", text_color=THEME_DANGER)
            else:
                ldap_status_lbl.configure(text=f"Imported/updated {count} user(s).", text_color=THEME_SUCCESS)
                refresh_ldap_roster_ui()
        app.after(0, on_done)

    threading.Thread(target=run, daemon=True).start()


def refresh_backup_history_ui():
    if not admin_tabs_active["value"]:
        return
    try:
        existing_ids = set(backup_history_tree.get_children())
    except Exception:
        return
    history = load_backup_history()
    seen_ids = set()
    for i, record in enumerate(history):
        row_id = str(i)
        seen_ids.add(row_id)
        size_kb = f"{record.get('size_bytes', 0) / 1024:.1f} KB"
        validation = record.get("validation_message", "")
        tag = "validation_ok" if record.get("validation_ok") else "validation_fail"
        row = (record.get("timestamp", ""), record.get("triggered_by", ""), size_kb,
              record.get("file_count", 0), validation, record.get("path", ""))
        if row_id in existing_ids:
            backup_history_tree.item(row_id, values=row, tags=(tag,))
        else:
            backup_history_tree.insert("", "end", iid=row_id, values=row, tags=(tag,))
    for stale_id in existing_ids - seen_ids:
        backup_history_tree.delete(stale_id)
    ensure_tree_empty_state(backup_history_tree, "No backups yet.\nRun 'Backup Now' or enable scheduled backups below.", icon="")

    if "backup_last_badge" in globals():
        if history:
            latest = history[-1]
            when = latest.get("timestamp", "unknown time")
            if latest.get("validation_ok"):
                backup_last_badge.update_status(f"Last backup OK — {when}", "success")
            else:
                backup_last_badge.update_status(f"Last backup FAILED — {when}", "danger")
        else:
            backup_last_badge.update_status("No backups yet", "pending")


def do_manual_backup():
    default_name = f"rapps_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    dest_path = filedialog.asksaveasfilename(
        title="Save Backup As…", defaultextension=".zip",
        filetypes=[("ZIP archive", "*.zip")], initialfile=default_name,
    )
    if not dest_path:
        return
    backup_manual_btn.configure(state="disabled")
    backup_status_lbl.configure(text="Backing up…", text_color=THEME_TEXT_MUTED)

    def run():
        try:
            record = run_backup_job(dest_path, triggered_by="manual")
            def on_done():
                backup_manual_btn.configure(state="normal")
                if record["validation_ok"]:
                    backup_status_lbl.configure(text=f"{record['validation_message']}", text_color=THEME_SUCCESS)
                else:
                    backup_status_lbl.configure(text=f"Backup completed but validation failed: {record['validation_message']}",
                                                text_color=THEME_WARNING)
                refresh_backup_history_ui()
            app.after(0, on_done)
        except Exception as e:
            log_exception("Manual backup failed")
            err_msg = str(e)
            def on_fail():
                backup_manual_btn.configure(state="normal")
                backup_status_lbl.configure(text=f"Backup failed: {err_msg}", text_color=THEME_DANGER)
            app.after(0, on_fail)

    threading.Thread(target=run, daemon=True).start()


def _run_restore(backup_path):
    backup_restore_btn.configure(state="disabled")
    backup_status_lbl.configure(text="Restoring… (a safety backup of the current state is being taken first)",
                                text_color=THEME_TEXT_MUTED)

    def run():
        try:
            restored, safety_path = restore_backup(backup_path)
            def on_done():
                backup_restore_btn.configure(state="normal")
                backup_status_lbl.configure(
                    text=f"Restored {len(restored)} file(s). A restart is recommended to fully apply "
                         f"the restored configuration. Safety backup: {os.path.basename(safety_path)}",
                    text_color=THEME_SUCCESS)
                refresh_backup_history_ui()
            app.after(0, on_done)
        except Exception as e:
            log_exception("Restore failed")
            err_msg = str(e)
            def on_fail():
                backup_restore_btn.configure(state="normal")
                backup_status_lbl.configure(text=f"Restore failed: {err_msg}", text_color=THEME_DANGER)
            app.after(0, on_fail)

    threading.Thread(target=run, daemon=True).start()


def do_restore_backup_dialog():
    backup_path = filedialog.askopenfilename(
        title="Select Backup to Restore", filetypes=[("ZIP archive", "*.zip")],
    )
    if not backup_path:
        return
    if not modern_askyesno(
        "Confirm Restore",
        "Restoring will overwrite current configuration, keys, routing rules, destination "
        "profiles, and other files with the contents of this backup.\n\n"
        "A safety backup of the CURRENT state will be taken automatically first, so this can "
        "be undone if needed.\n\nContinue?",
    ):
        return
    _run_restore(backup_path)


def _on_backup_history_row_double_click(event):
    row_id = backup_history_tree.identify_row(event.y)
    if not row_id:
        return
    history = load_backup_history()
    try:
        record = history[int(row_id)]
    except (ValueError, IndexError):
        return
    path = record.get("path", "")
    if not os.path.isfile(path):
        modern_showerror("Restore Failed", f"Backup file no longer exists at:\n{path}")
        return
    if not modern_askyesno(
        "Confirm Restore",
        f"Restore from this backup?\n\n{path}\n\n"
        "A safety backup of the CURRENT state will be taken automatically first.",
    ):
        return
    _run_restore(path)


def do_save_report_email_schedule():
    try:
        hour = int(report_sched_hour_var.get())
        if not (0 <= hour <= 23):
            raise ValueError
    except ValueError:
        report_sched_status_lbl.configure(text="Hour must be a whole number from 0-23.", text_color=THEME_DANGER)
        return
    cfg = {
        "enabled": report_sched_enabled_var.get(),
        "frequency": report_sched_freq_var.get(),
        "hour": hour,
    }
    try:
        save_report_email_schedule(cfg)
        status = "enabled" if cfg["enabled"] else "disabled"
        report_sched_status_lbl.configure(
            text=f"Scheduled report email {status} -- {cfg['frequency']} at {hour:02d}:00", text_color=THEME_SUCCESS)
    except Exception as e:
        log_exception("Failed to save report email schedule")
        report_sched_status_lbl.configure(text=f"Failed to save schedule: {e}", text_color=THEME_DANGER)


def do_save_backup_schedule():
    try:
        hour = int(backup_sched_hour_var.get())
        if not (0 <= hour <= 23):
            raise ValueError
    except ValueError:
        backup_status_lbl.configure(text="Hour must be a whole number from 0-23.", text_color=THEME_DANGER)
        return
    cfg = {
        "enabled": backup_sched_enabled_var.get(),
        "frequency": backup_sched_freq_var.get(),
        "hour": hour,
    }
    try:
        save_backup_schedule(cfg)
        status = "enabled" if cfg["enabled"] else "disabled"
        backup_status_lbl.configure(
            text=f"Scheduled backup {status} -- {cfg['frequency']} at {hour:02d}:00", text_color=THEME_SUCCESS)
        refresh_next_backup_indicator()
    except Exception as e:
        log_exception("Failed to save backup schedule")
        backup_status_lbl.configure(text=f"Failed to save schedule: {e}", text_color=THEME_DANGER)


def refresh_health_tree():
    """Repopulates the PACS Health tree from destination_health_cache.
    Cheap (no network I/O) -- safe to call on every periodic_refresh tick."""
    if not admin_tabs_active["value"]:
        return
    try:
        existing_ids = set(health_tree.get_children())
    except Exception:
        return  # tab not built yet

    dests = load_destinations()
    seen_ids = set()
    counts = {"online": 0, "degraded": 0, "offline": 0, "unchecked": 0}
    for dest in dests:
        name = dest["name"]
        seen_ids.add(name)
        record = destination_health_cache.get(name)

        if record is None:
            status_text, tag = "Not yet checked", ""
            resp_time = avg_latency = last_success = last_failure = "—"
            consecutive = 0
            uptime_pct = "—"
            counts["unchecked"] += 1
        else:
            consecutive = record.get("consecutive_failures", 0)
            if record["online"]:
                status_text, tag = "● Online", "status_online"
                counts["online"] += 1
            elif consecutive >= 3:
                status_text, tag = "● Offline", "status_offline"
                counts["offline"] += 1
            else:
                status_text, tag = "● Degraded", "status_degraded"
                counts["degraded"] += 1
            resp_time = f"{record['response_time_ms']:.0f} ms" if record.get("response_time_ms") is not None else "—"
            history = record.get("uptime_history") or []
            latencies = [record["response_time_ms"]] if record.get("response_time_ms") is not None else []
            avg_latency = f"{(sum(latencies) / len(latencies)):.0f} ms" if latencies else "—"
            last_success = _fmt_ts_or_dash(record.get("last_success"))
            last_failure = _fmt_ts_or_dash(record.get("last_failure"))
            uptime_pct = f"{(sum(1 for h in history if h) / len(history) * 100):.1f}%" if history else "—"

            # A.2: small secondary indicator, not a replacement for the
            # DICOM status above -- only shown for destinations that have
            # document transfer enabled at all.
            doc_reachable = record.get("doc_transfer_reachable")
            if doc_reachable is not None:
                status_text += "  \u00b7 Docs OK" if doc_reachable else "  \u00b7 Docs unreachable"

        row = (name, status_text, resp_time, avg_latency, last_success, last_failure, consecutive, uptime_pct)
        if name in existing_ids:
            health_tree.item(name, values=row, tags=(tag,) if tag else ())
        else:
            health_tree.insert("", "end", iid=name, values=row, tags=(tag,) if tag else ())

    for stale_id in existing_ids - seen_ids:
        health_tree.delete(stale_id)
    ensure_tree_empty_state(health_tree, "No destinations configured yet.\nAdd one in the Destinations tab to start monitoring.", icon="")

    if "health_summary_online_badge" in globals():
        health_summary_online_badge.update_status(str(counts["online"]), "connected")
        health_summary_degraded_badge.update_status(str(counts["degraded"]), "warning")
        health_summary_offline_badge.update_status(str(counts["offline"]), "offline")
        health_summary_unchecked_badge.update_status(str(counts["unchecked"]), "pending")





# 5.1 -- destination_health_cache above is live/in-memory only (reset on
# restart). This persists periodic (timestamp, online) snapshots per
# destination to a small rolling JSON log, capped per destination, so the
# Destinations tab can show an uptime% / sparkline that survives restarts.
DESTINATION_HEALTH_HISTORY_FILE = "destination_health_history.json"  # not secret -- just online/offline timestamps
DEST_HEALTH_HISTORY_MAX_PER_DEST = 500

_dest_health_history_cache = {"value": None}
_dest_health_history_last_flush = {"value": 0}


def load_destination_health_history():
    if _dest_health_history_cache["value"] is not None:
        return _dest_health_history_cache["value"]
    data = {}
    try:
        if os.path.exists(DESTINATION_HEALTH_HISTORY_FILE):
            with open(DESTINATION_HEALTH_HISTORY_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                data = loaded
    except Exception:
        log_exception("Failed to load destination_health_history.json")
    _dest_health_history_cache["value"] = data
    return data


def save_destination_health_history(data):
    try:
        atomic_write(DESTINATION_HEALTH_HISTORY_FILE, json.dumps(data, indent=2))
    except Exception:
        log_exception("Failed to save destination_health_history.json")
    _dest_health_history_cache["value"] = data


def record_destination_health_snapshot(name, online):
    """Appends one (timestamp, online) snapshot for `name`, capped at
    DEST_HEALTH_HISTORY_MAX_PER_DEST entries. Fed by the exact same
    C-ECHO checks that already update the in-memory
    destination_health_cache -- see record_destination_health_result().
    Disk writes are throttled to at most once every 30s since this can be
    called as often as every health-check tick."""
    history = load_destination_health_history()
    entries = history.setdefault(name, [])
    entries.append({"ts": time.time(), "online": bool(online)})
    if len(entries) > DEST_HEALTH_HISTORY_MAX_PER_DEST:
        del entries[:-DEST_HEALTH_HISTORY_MAX_PER_DEST]
    now = time.time()
    if now - _dest_health_history_last_flush["value"] >= 30:
        _dest_health_history_last_flush["value"] = now
        save_destination_health_history(history)


def get_destination_uptime_history(name, limit=60):
    """Most recent `limit` persisted snapshots for `name`, oldest first."""
    return load_destination_health_history().get(name, [])[-limit:]


def get_destination_uptime_pct(name, limit=60):
    entries = get_destination_uptime_history(name, limit=limit)
    if not entries:
        return None
    return sum(1 for e in entries if e.get("online")) * 100.0 / len(entries)


def record_destination_health_result(name, ok, msg, response_time_ms):
    """Updates destination_health_cache[name] with a new C-ECHO result and
    fires destination_online/destination_offline notifications on state
    transitions (never on every check -- only when status actually
    changes, to avoid notification spam)."""
    now = time.time()
    existing = destination_health_cache.get(name)
    previous_online = existing["online"] if existing else None

    record = existing or {
        "online": None, "message": "", "response_time_ms": None,
        "checked_at": None, "last_success": None, "last_failure": None,
        "consecutive_failures": 0,
        "uptime_history": deque(maxlen=DESTINATION_UPTIME_HISTORY_LEN),
    }
    record["online"] = ok
    record["message"] = msg
    record["response_time_ms"] = response_time_ms
    record["checked_at"] = now
    if ok:
        record["last_success"] = now
        record["consecutive_failures"] = 0
    else:
        record["last_failure"] = now
        record["consecutive_failures"] = record.get("consecutive_failures", 0) + 1
    record["uptime_history"].append(ok)
    destination_health_cache[name] = record
    record_destination_health_snapshot(name, ok)

    if previous_online is not None and previous_online != ok:
        if ok:
            notify_event("destination_online", "Destination Online", f"{name} is back online.")
        else:
            notify_event("destination_offline", "Destination Offline", f"{name} went offline: {msg}")


def check_doc_transfer_reachability(dest):
    """A.2: plain TCP connect probe (no protocol handshake) against a
    destination's doc_transfer_port, independent of the DICOM C-ECHO check
    above. No-ops (returns None) for destinations that don't have document
    transfer enabled, so sites not using this feature never open an extra
    connection or gain an extra cache field."""
    if not dest.get("doc_transfer_enabled"):
        return None
    host = dest.get("ip") if dest.get("doc_transfer_use_dicom_host", True) else dest.get("doc_transfer_ip")
    ok, port_val = validate_port(dest.get("doc_transfer_port"), "Document Transfer Port")
    if not ok or not host:
        return False
    try:
        with socket.create_connection((host, port_val), timeout=DOC_TRANSFER_CONNECT_TIMEOUT_SEC):
            return True
    except Exception:
        return False


def check_destination_health(dest):
    """Blocking C-ECHO + timing against one destination profile dict.
    Always call this from a background thread -- it does real network
    I/O. Updates destination_health_cache via record_destination_health_result."""
    start = time.time()
    try:
        ok, msg = dicom_echo(dest["ae"], dest["ip"], int(dest["port"]),
                              calling_ae=dest.get("calling_ae") or DEFAULT_PUSH_CALLING_AE,
                              dest_label=dest.get("name"))
    except Exception as e:
        ok, msg = False, str(e)
    response_time_ms = round((time.time() - start) * 1000, 1)
    tls_verified = get_tls_verification_state(dest.get("name"))
    if tls_verified is False:
        msg = f"{msg} (⚠ TLS unverified -- no ca_cert configured)"
    record_destination_health_result(dest["name"], ok, msg, response_time_ms)

    # A.2: same PACS_HEALTH_CHECK_INTERVAL_SEC cadence, no separate polling
    # loop -- just a second field written onto the same cache entry.
    doc_reachable = check_doc_transfer_reachability(dest)
    if doc_reachable is not None:
        record = destination_health_cache.get(dest["name"])
        if record is not None:
            record["doc_transfer_reachable"] = doc_reachable

    return ok, msg, response_time_ms


def _kick_off_background_destination_check():
    """Non-blocking C-ECHO health check against the default push
    destination, throttled to at most once every 30s. Result lands in
    destination_health_cache for the NEXT dashboard refresh to display --
    the current tick never blocks waiting on network I/O."""
    now = time.time()
    if _dest_health_check_inflight["value"] or (now - _last_dest_health_check_at[0]) < 30:
        return
    dest = get_default_destination()
    if not dest:
        return
    _dest_health_check_inflight["value"] = True
    _last_dest_health_check_at[0] = now

    def run():
        try:
            check_destination_health(dest)
        except Exception:
            log_exception("Background destination health check failed")
        finally:
            _dest_health_check_inflight["value"] = False

    threading.Thread(target=run, daemon=True).start()


def _format_bytes_per_sec(bps):
    if bps < 1024:
        return f"{bps:.0f} B/s"
    if bps < 1024 * 1024:
        return f"{bps / 1024:.1f} KB/s"
    return f"{bps / (1024 * 1024):.1f} MB/s"


def refresh_home_dashboard():
    """Populates the new landing-page Dashboard tab. Runs on every
    periodic_refresh() tick (2s) regardless of admin/user mode -- this
    tab is visible in both, unlike the Admin-only operational dashboard."""
    _roll_daily_stats_if_needed()

    # ---- Receiver ----
    running = receiver_state.get("running")
    if running:
        dash_home_recv_status.configure(text="● Running", text_color=THEME_SUCCESS)
    else:
        dash_home_recv_status.configure(text="● Stopped", text_color=THEME_DANGER)
    dash_home_recv_ae.configure(text=ae_entry.get() or "—")
    dash_home_recv_port.configure(text=port_entry.get() or "—")
    try:
        dash_home_recv_ip.configure(text=socket.gethostbyname(socket.gethostname()))
    except Exception:
        dash_home_recv_ip.configure(text="—")

    active_assoc = 0
    try:
        server_ae = receiver_state.get("server_ae")
        if server_ae is not None:
            active_assoc = len(getattr(server_ae, "active_associations", []) or [])
    except Exception:
        active_assoc = 0
    dash_home_recv_assoc.configure(text=str(active_assoc))

    with data_lock:
        total_patients = len(patient_data)
        queue_size = sum(1 for d in patient_data.values()
                          if d.get("status") in (STATUS_PENDING, STATUS_RECEIVED, STATUS_IMPORTED))
        pending_count = sum(1 for d in patient_data.values() if d.get("status") == STATUS_PENDING)
        retry_count_total = sum(1 for d in patient_data.values() if d.get("status") == STATUS_RETRYING)

    with _dashboard_stats_lock:
        studies_recv_today = len(daily_stats["studies_received_uids"])
        images_recv_today = daily_stats["images_received"]
        docs_recv_today = daily_stats["documents_received"]
        studies_sent_today = len(daily_stats["studies_sent_uids"])
        images_sent_today = daily_stats["images_sent"]
        failed_today = daily_stats["failed_transfers"]

    dash_home_studies_recv.configure(text=str(studies_recv_today))
    dash_home_images_recv.configure(text=str(images_recv_today))
    dash_home_docs_recv.configure(text=str(docs_recv_today))
    dash_home_total_patients.configure(text=str(total_patients))
    dash_home_recv_queue.configure(text=str(queue_size))

    if queue_size >= QUEUE_GROWING_THRESHOLD:
        now = time.time()
        if now - _queue_growing_notify_throttle[0] > 600:  # once per 10 min
            _queue_growing_notify_throttle[0] = now
            notify_event("queue_growing", "Queue Growing",
                         f"Current queue size is {queue_size} (threshold: {QUEUE_GROWING_THRESHOLD}).")

    # ---- Pusher ----
    if push_job.get("running"):
        dash_home_push_status.configure(text="● Sending", text_color=STATUS_COLORS[STATUS_SENDING])
    else:
        dash_home_push_status.configure(text="Idle", text_color=THEME_TEXT_MUTED)

    _kick_off_background_destination_check()
    dest = get_default_destination()
    if dest and dest["name"] in destination_health_cache:
        health = destination_health_cache[dest["name"]]
        if health["online"]:
            dash_home_dest_status.configure(text=f"● Online ({dest['name']})", text_color=THEME_SUCCESS)
        else:
            dash_home_dest_status.configure(text=f"● Offline ({dest['name']})", text_color=THEME_DANGER)
    elif dest:
        dash_home_dest_status.configure(text=f"Checking… ({dest['name']})", text_color=THEME_TEXT_MUTED)
    else:
        dash_home_dest_status.configure(text="No destination configured", text_color=THEME_TEXT_MUTED)

    dash_home_images_sent.configure(text=str(images_sent_today))
    dash_home_studies_sent.configure(text=str(studies_sent_today))
    dash_home_failed_transfers.configure(text=str(failed_today))
    dash_home_pending_queue.configure(text=str(pending_count))
    dash_home_retry_queue.configure(text=str(retry_count_total))

    img_rate, _eta = get_push_throughput_eta()
    dash_home_throughput_img.configure(text=f"{img_rate:.2f}")
    # Approximate MB/sec from images/sec using the average file size of
    # whatever was most recently received (no extra disk scanning added).
    avg_file_size_mb = 0.0
    try:
        if os.path.isdir(OUTPUT_DIR):
            sample = []
            with data_lock:
                sample_pids = list(patient_data.keys())[:5]
            for pid in sample_pids:
                folder = get_patient_folder(pid)
                if os.path.isdir(folder):
                    for fname in os.listdir(folder)[:3]:
                        fpath = os.path.join(folder, fname)
                        if os.path.isfile(fpath):
                            sample.append(os.path.getsize(fpath))
            if sample:
                avg_file_size_mb = (sum(sample) / len(sample)) / (1024 * 1024)
    except Exception:
        avg_file_size_mb = 0.0
    dash_home_throughput_mb.configure(text=f"{(img_rate * avg_file_size_mb):.2f}")

    # ---- System ----
    if PSUTIL_AVAILABLE:
        try:
            cpu_pct = psutil.cpu_percent(interval=None)
            _last_cpu_percent[0] = cpu_pct
            dash_home_cpu.configure(text=f"{cpu_pct:.0f}%")
            mem = psutil.virtual_memory()
            dash_home_ram.configure(text=f"{mem.percent:.0f}%  ({mem.used / (1024**3):.1f}/{mem.total / (1024**3):.1f} GB)")
        except Exception:
            dash_home_cpu.configure(text="N/A")
            dash_home_ram.configure(text="N/A")
    else:
        dash_home_cpu.configure(text="N/A")
        dash_home_ram.configure(text="N/A")

    try:
        total_b, used_b, free_b = shutil.disk_usage(os.path.abspath("."))
        total_gb, free_gb = total_b / (1024 ** 3), free_b / (1024 ** 3)
        used_pct = (used_b / total_b * 100) if total_b else 0
        dash_home_disk.configure(text=f"{used_pct:.0f}% used")
        dash_home_disk_free.configure(text=f"{free_gb:.1f} GB free")
    except Exception:
        dash_home_disk.configure(text="N/A")
        dash_home_disk_free.configure(text="N/A")

    uptime_sec = int(time.time() - APP_START_TIME)
    hours, rem = divmod(uptime_sec, 3600)
    minutes, seconds = divmod(rem, 60)
    dash_home_uptime.configure(text=f"{hours}h {minutes}m {seconds}s")

    net_up_bps = net_down_bps = 0.0
    if PSUTIL_AVAILABLE:
        try:
            counters = psutil.net_io_counters()
            now = time.time()
            if _net_io_baseline["time"] is not None:
                dt = max(now - _net_io_baseline["time"], 0.001)
                net_up_bps = max(0.0, (counters.bytes_sent - _net_io_baseline["bytes_sent"]) / dt)
                net_down_bps = max(0.0, (counters.bytes_recv - _net_io_baseline["bytes_recv"]) / dt)
            _net_io_baseline.update(time=now, bytes_sent=counters.bytes_sent, bytes_recv=counters.bytes_recv)
            dash_home_net_up.configure(text=_format_bytes_per_sec(net_up_bps))
            dash_home_net_down.configure(text=_format_bytes_per_sec(net_down_bps))
        except Exception:
            dash_home_net_up.configure(text="N/A")
            dash_home_net_down.configure(text="N/A")
    else:
        dash_home_net_up.configure(text="N/A")
        dash_home_net_down.configure(text="N/A")

    # ---- Graph history + redraw ----
    dash_history["studies_received"].append(studies_recv_today)
    dash_history["studies_sent"].append(studies_sent_today)
    dash_history["failed_transfers"].append(failed_today)
    dash_history["queue_size"].append(queue_size)
    dash_history["network_kbps"].append((net_up_bps + net_down_bps) / 1024.0)

    graph_range = dash_graph_range_var.get() if "dash_graph_range_var" in globals() else "Live"
    daily_trend = None
    if graph_range in ("7d", "30d"):
        daily_trend = get_daily_stats_trend(days=7 if graph_range == "7d" else 30)
    # B.5: docs_delivered/doc_failure_rate have no "Live" (in-session)
    # equivalent -- always show a daily trend, defaulting to 7 days when
    # "Live" is selected (mirroring how queue_size/network_kbps always
    # fall back to live regardless of range, just the opposite case).
    doc_daily_trend = get_doc_transfer_daily_trend(days=7 if graph_range != "30d" else 30)

    for key, canvas in dash_graph_canvases.items():
        if key in ("docs_delivered", "doc_failure_rate"):
            trend_field = "docs_delivered" if key == "docs_delivered" else "failure_rate_pct"
            _draw_sparkline(canvas, [d.get(trend_field, 0) for d in doc_daily_trend], DASH_GRAPH_COLORS[key])
        elif daily_trend is not None and key in ("studies_received", "studies_sent", "failed_transfers"):
            _draw_sparkline(canvas, [d.get(key, 0) for d in daily_trend], DASH_GRAPH_COLORS[key])
        else:
            # queue_size / network_kbps have no daily equivalent -- always
            # show the live in-session window for those regardless of range.
            _draw_sparkline(canvas, list(dash_history[key]), DASH_GRAPH_COLORS[key])


def refresh_admin_dashboard():
    """Populates the Admin Dashboard from data that already exists --
    patient_data, receiver_state, destinations, and audit.log. No new
    background polling/scanning is added; this reuses the existing
    data_lock-guarded refresh pattern and is called on-demand (mode
    switch, manual Refresh, or the normal periodic UI refresh)."""
    if not admin_tabs_active["value"]:
        return

    today = datetime.date.today().isoformat()
    received = pushed = failed = 0
    stale = 0
    docs_pending = 0
    failures = []
    with data_lock:
        for pid, d in patient_data.items():
            t = d.get("time", "")
            if isinstance(t, str) and t.startswith(today):
                received += 1
            sent_t = d.get("sent_time", "")
            status = d.get("status", "")
            if isinstance(sent_t, str) and sent_t.startswith(today) and status == STATUS_SENT:
                pushed += 1
            if isinstance(t, str) and t.startswith(today) and status == STATUS_FAILED:
                failed += 1
            if is_pending_stale(d):
                stale += 1
            if status == STATUS_FAILED:
                failures.append((pid, d))
            # B.4: DICOM delivered (STATUS_SENT) but the document leg
            # either errored out or was never confirmed at all, despite a
            # local report/history file existing.
            if d.get("last_doc_transfer_error"):
                docs_pending += 1
            elif status == STATUS_SENT and not d.get("doc_transfer_sent_at"):
                if os.path.isfile(get_report_path(pid)) or os.path.isfile(get_history_path(pid)):
                    docs_pending += 1

    dash_received_lbl.configure(text=str(received))
    dash_pushed_lbl.configure(text=str(pushed))
    dash_failed_lbl.configure(text=str(failed))
    dash_docs_pending_lbl.configure(text=str(docs_pending))
    dash_stale_lbl.configure(text=str(stale))

    # Disk gauge (reuses get_free_disk_gb(), just adds a visible bar+label)
    try:
        free_gb = get_free_disk_gb()
        total, used, free_bytes = shutil.disk_usage(os.path.abspath("."))
        total_gb = total / (1024 ** 3)
        frac_free = max(0.0, min(1.0, free_gb / total_gb)) if total_gb else 0.0
        dash_disk_bar.set(frac_free)
        dash_disk_lbl.configure(text=f"{free_gb:.1f} GB free of {total_gb:.1f} GB")
    except Exception:
        log_exception("Dashboard: failed to read disk usage")

    # Receiver status (reuses receiver_state, no new state)
    if receiver_state.get("running"):
        ae = receiver_state.get("server_ae") or "?"
        dash_receiver_status_badge.update_status(f"Running ({ae}@{port_entry.get() or '?'})", "running")
    else:
        dash_receiver_status_badge.update_status("Stopped", "stopped")

    # Recent failures panel (last N, newest first). Diff-based: only
    # touch rows that actually changed, instead of delete+reinsert every
    # single 2s tick -- that full rebuild is what made the Dashboard feel
    # laggy while just sitting open with a lot of history.
    failures.sort(key=lambda kv: kv[1].get("time", ""), reverse=True)
    _sync_tree_rows(dash_failures_tree, [
        (pid, (pid, d.get("patient_name", ""), d.get("time", ""), (d.get("last_error", "") or "")[:80]))
        for pid, d in failures[:25]
    ], empty_message="No recent failures — all transfers succeeding.", empty_icon="")

    # Per-destination push health: rolling PUSH-OK/PUSH-FAILED counts parsed
    # from the existing, already-written, plain-text append-only audit.log
    # (no new logging added).
    ok_counts, fail_counts = parse_audit_log_push_counts()
    _sync_tree_rows(dash_dest_health_tree, [
        (dest["name"], (dest["name"], "(double-click to test)",
                        ok_counts.get(dest["name"], 0), fail_counts.get(dest["name"], 0)))
        for dest in load_destinations()
    ], empty_message="No destinations configured yet.\nAdd one in the Destinations tab.", empty_icon="")


_audit_push_count_cache = {"offset": 0, "ok": {}, "fail": {}, "inode": None}


def parse_audit_log_push_counts():
    """Parse audit.log's existing PUSH-OK/PUSH-FAILED lines into rolling
    per-destination success/fail counts. audit.log's format is untouched;
    this is read-only parsing for the dashboard, matching the existing
    'PUSH-OK: <pid> -> <destination> ...' / 'PUSH-FAILED: <pid> -> <destination> ...'
    style already written by write_audit_log() elsewhere in this file.

    Incremental: only the bytes appended since the last call are scanned
    (tracked by file offset), so this stays cheap no matter how large
    audit.log grows -- the dashboard used to re-read the entire file from
    byte 0 on every 2s refresh tick, which got slower over the life of the
    log."""
    cache = _audit_push_count_cache
    try:
        size = os.path.getsize(AUDIT_LOG)
        # If the file shrank (cleared/rotated) or doesn't match what we
        # last saw, start over from scratch instead of seeking past EOF.
        if size < cache["offset"]:
            cache["offset"] = 0
            cache["ok"] = {}
            cache["fail"] = {}

        with open(AUDIT_LOG, "r", encoding="utf-8", errors="replace") as f:
            f.seek(cache["offset"])
            for line in f:
                if "PUSH-OK" not in line and "PUSH-FAILED" not in line:
                    continue
                is_ok = "PUSH-OK" in line
                # Lines look like: [timestamp] PUSH-OK: <pid> -> <dest> (...)
                dest_name = None
                if "->" in line:
                    tail = line.split("->", 1)[1].strip()
                    dest_name = tail.split()[0].rstrip(":") if tail else None
                if not dest_name:
                    continue
                bucket = cache["ok"] if is_ok else cache["fail"]
                bucket[dest_name] = bucket.get(dest_name, 0) + 1
            cache["offset"] = f.tell()
    except FileNotFoundError:
        pass
    except Exception:
        log_exception("Dashboard: failed to parse audit.log for push health")
    return dict(cache["ok"]), dict(cache["fail"])


def export_worklist_view_to_csv(tree):
    """Exports exactly what's currently visible in the given worklist tree
    (respecting its search text + Status/Date filters, and current sort
    order) to a CSV -- handy for quick ad-hoc reporting without having to
    dig through the full audit log."""
    row_ids = tree.get_children("")
    if not row_ids:
        modern_showinfo("Export View to CSV", "There are no rows in the current view to export.")
        return
    default_name = f"worklist_view_{datetime.date.today().isoformat()}.csv"
    out_path = filedialog.asksaveasfilename(
        title="Export current view to CSV", defaultextension=".csv",
        initialfile=default_name, filetypes=[("CSV files", "*.csv")])
    if not out_path:
        return
    try:
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            export_cols = [c for c in WL_COLUMNS if c != "sel"]
            writer.writerow([WL_HEADINGS[c] for c in export_cols])
            for iid in row_ids:
                values = tree.item(iid, "values")
                writer.writerow([v for c, v in zip(WL_COLUMNS, values) if c != "sel"])
        modern_showinfo("Export View to CSV", f"Exported {len(row_ids)} row(s) to:\n{out_path}")
    except Exception as e:
        log_exception("Failed to export worklist view to CSV")
        modern_showerror("Export Failed", str(e))


def export_structured_log_view_to_csv():
    """6.1 -- mirrors export_worklist_view_to_csv's 'export exactly what's
    currently rendered' behavior, but for the Logs tab's structured view:
    whatever refresh_structured_log_view() currently has filtered/rendered
    into log_struct_tree (search text + severity + destination + date
    range), not the full log via export_logs_zip."""
    if "log_struct_tree" not in globals():
        modern_showerror("Export Failed", "Structured log view isn't built yet.")
        return
    row_ids = log_struct_tree.get_children("")
    if not row_ids:
        modern_showinfo("Export Filtered to CSV", "There are no rows in the current filtered view to export.")
        return
    default_name = f"log_view_{datetime.date.today().isoformat()}.csv"
    out_path = filedialog.asksaveasfilename(
        title="Export filtered log view to CSV", defaultextension=".csv",
        initialfile=default_name, filetypes=[("CSV files", "*.csv")])
    if not out_path:
        return
    try:
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([log_struct_headings[c] for c in log_struct_cols])
            for iid in row_ids:
                writer.writerow(log_struct_tree.item(iid, "values"))
        modern_showinfo("Export Filtered to CSV", f"Exported {len(row_ids)} row(s) to:\n{out_path}")
    except Exception as e:
        log_exception("Failed to export filtered log view to CSV")
        modern_showerror("Export Failed", str(e))


def do_export_audit_log_csv():
    """Admin-only 'Export audit log to CSV' button. Parses the existing
    plain-text, append-only audit.log into structured rows for external
    reporting -- the log FILE format itself is never changed."""
    try:
        default_name = f"audit_log_export_{datetime.date.today().isoformat()}.csv"
        out_path = filedialog.asksaveasfilename(
            title="Export audit log to CSV", defaultextension=".csv",
            initialfile=default_name, filetypes=[("CSV files", "*.csv")])
        if not out_path:
            return
        rows = []
        with open(AUDIT_LOG, "r", encoding="utf-8", errors="replace") as f:
            for line in f:
                line = line.rstrip("\n")
                if not line.startswith("["):
                    continue
                ts_end = line.find("]")
                if ts_end == -1:
                    continue
                timestamp = line[1:ts_end]
                rest = line[ts_end + 1:].strip()
                event_type, _, detail = rest.partition(":")
                rows.append((timestamp, event_type.strip(), detail.strip()))
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "event_type", "detail"])
            writer.writerows(rows)
        modern_showinfo("Export complete", f"Exported {len(rows)} audit log entries to:\n{out_path}")
    except Exception as e:
        log_exception("Failed to export audit log to CSV")
        modern_showerror("Export failed", str(e))

# =========================================================
# RECEIVER CALLBACKS
# =========================================================

def _load_receiver_config():
    config = decrypt_and_load(RECEIVER_CONFIG)
    if config:
        parts = config.split("|")
        if len(parts) == 2:
            ae_entry.configure(state="normal")
            ae_entry.delete(0, "end")
            ae_entry.insert(0, parts[0])
            port_entry.configure(state="normal")
            port_entry.delete(0, "end")
            port_entry.insert(0, parts[1])
            ae_entry.configure(state="disabled")
            port_entry.configure(state="disabled")


def do_save_receiver():
    ae = ae_entry.get().strip()
    port = port_entry.get().strip()
    if not ae or not port:
        modern_showerror("Error", "AE Title and Port are required.")
        return

    ok, ae_clean_or_err = validate_ae_title(ae, "Receiver AE Title")
    if not ok:
        modern_showerror("Invalid AE Title", ae_clean_or_err)
        return

    ok, port_val_or_err = validate_port(port, "Port")
    if not ok:
        modern_showerror("Invalid Port", port_val_or_err)
        return

    def proceed():
        encrypt_and_save(RECEIVER_CONFIG, f"{ae}|{port}")
        ae_entry.configure(state="disabled")
        port_entry.configure(state="disabled")
        modern_showinfo("Saved", "Receiver configuration saved.")

    local_otp_confirm(
        "Verify Receiver Configuration",
        f"AE Title: {ae}\nPort: {port}\n\nConfirm you want to save this configuration.",
        proceed,
    )


def do_edit_receiver():
    ae_entry.configure(state="normal")
    port_entry.configure(state="normal")


def do_start_receiver():
    config = decrypt_and_load(RECEIVER_CONFIG)
    if not config:
        modern_showerror("Error", "Receiver config missing. Save it first.")
        return
    parts = config.split("|")
    if len(parts) != 2:
        modern_showerror("Error", "Receiver config is corrupt. Re-save it.")
        return
    ae_title, port = parts
    if receiver_state.get("running"):
        modern_showwarning("Already Running", "Receiver is already running.")
        return
    threading.Thread(target=start_receiver_server, args=(ae_title, port), daemon=True).start()


def do_stop_receiver():
    _receiver_stop_was_user_initiated["value"] = True
    threading.Thread(target=stop_receiver_server, daemon=True).start()


def do_import_folder():
    folder = filedialog.askdirectory(title="Select folder containing DICOM files")
    if not folder:
        return

    def progress_cb(done, total):
        ui_event_queue.put(("import_progress", (done, max(total, 1))))

    import_btn.configure(state="disabled")
    threading.Thread(target=import_folder, args=(folder, progress_cb), daemon=True).start()


receiver_save_btn.configure(command=do_save_receiver)
receiver_edit_btn.configure(command=do_edit_receiver)
start_btn.configure(command=do_start_receiver)
stop_btn.configure(command=do_stop_receiver)
tray_btn.configure(command=minimize_to_tray)
import_btn.configure(command=do_import_folder)
rec_refresh_btn.configure(command=lambda: refresh_worklists(rec_search_var.get()))
rec_search_var.trace_add("write", lambda *_: refresh_worklists(rec_search_var.get()))
rec_status_filter_var.trace_add("write", lambda *_: populate_tree(rec_tree, rec_search_var.get(), force=True))
rec_date_filter_var.trace_add("write", lambda *_: populate_tree(rec_tree, rec_search_var.get(), force=True))
rec_report_filter_var.trace_add("write", lambda *_: populate_tree(rec_tree, rec_search_var.get(), force=True))
rec_export_btn.configure(command=lambda: export_worklist_view_to_csv(rec_tree))

# =========================================================
# PUSHER CALLBACKS
# =========================================================

def _get_active_dest():
    name = push_dest_var.get()
    d = get_destination_by_name(name)
    if not d:
        d = get_default_destination()
    return d


# 2.3 -- Pre-push size/ETA estimate, shown in a confirmation dialog for
# batches over the thresholds below.
LARGE_PUSH_STUDY_THRESHOLD = 20
LARGE_PUSH_SIZE_THRESHOLD_BYTES = 500 * 1024 * 1024


def _estimate_push_size_and_eta(pids):
    """Total on-disk size (bytes) for every .dcm file across `pids`, and
    the estimated transfer duration (seconds) at the currently configured
    bandwidth limit -- None if unlimited (no ceiling to estimate against)."""
    total_bytes = 0
    for pid in pids:
        for fpath in list_patient_dcm_files(pid):
            try:
                total_bytes += os.path.getsize(fpath)
            except OSError:
                pass
    mbps = get_effective_bandwidth_mbps()
    eta_seconds = None
    if mbps > 0:
        total_megabits = (total_bytes * 8) / 1_000_000
        eta_seconds = total_megabits / mbps
    return total_bytes, eta_seconds


def _format_bytes(n):
    n = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"


def _format_duration(seconds):
    if seconds is None:
        return "unknown (bandwidth unlimited)"
    seconds = int(seconds)
    if seconds < 60:
        return f"~{seconds}s"
    minutes, sec = divmod(seconds, 60)
    if minutes < 60:
        return f"~{minutes}m {sec}s"
    hours, minutes = divmod(minutes, 60)
    return f"~{hours}h {minutes}m"


def _confirm_large_push_if_needed(pids, dest_count=1, dests=None):
    """Shows a modern_confirm with total size + estimated duration for
    batches over LARGE_PUSH_STUDY_THRESHOLD studies or
    LARGE_PUSH_SIZE_THRESHOLD_BYTES total size. Small/ordinary pushes skip
    straight through with no extra dialog. Returns True to proceed.

    C.3: when this dialog does fire, also mentions how many of the
    destinations being pushed to have document transfer enabled --
    purely additive to an already-existing confirmation, never a new one."""
    total_bytes, eta_seconds = _estimate_push_size_and_eta(pids)
    if len(pids) <= LARGE_PUSH_STUDY_THRESHOLD and total_bytes <= LARGE_PUSH_SIZE_THRESHOLD_BYTES:
        return True
    lines = [f"About to push {len(pids)} patient(s), {_format_bytes(total_bytes)} total"]
    if dest_count > 1:
        lines[0] += f" to {dest_count} destinations"
    lines.append(f"Estimated duration per destination: {_format_duration(eta_seconds)}")
    if dest_count > 1 and eta_seconds is not None:
        lines.append(f"Estimated total across all destinations (sequential): {_format_duration(eta_seconds * dest_count)}")
    if dests:
        doc_count = sum(1 for d in dests if d.get("doc_transfer_enabled"))
        if doc_count:
            lines.append(f"This will also send Report/History documents to "
                        f"{doc_count} of {len(dests)} selected destination(s).")
    lines.append("\nContinue?")
    return modern_confirm("Confirm Large Push", "\n".join(lines))


def do_push_selected():
    pids = _effective_selection(push_tree)
    if not pids:
        modern_showwarning("No Selection", "Select at least one patient to push.")
        return
    if push_job["running"]:
        modern_showwarning("Busy", "A push job is already running.")
        return
    anon = anon_var.get()
    if push_multi_dest_toggle_var.get():
        dests = _get_selected_multi_destinations()
        if not dests:
            modern_showwarning("No Destinations", "Check at least one destination to push to.")
            return
        if not _confirm_large_push_if_needed(pids, dest_count=len(dests), dests=dests):
            return
        threading.Thread(target=run_push_job_multi, args=(pids, dests),
                          kwargs={"anonymize": anon}, daemon=True).start()
    else:
        dest = _get_active_dest()
        if not _confirm_large_push_if_needed(pids, dests=[dest] if dest else None):
            return
        threading.Thread(target=run_push_job, args=(pids,), kwargs={"destination": dest, "anonymize": anon}, daemon=True).start()


def do_push_all_pending():
    with data_lock:
        pids = [pid for pid, d in patient_data.items()
                if d.get("status") not in (STATUS_SENT, STATUS_SENDING)]
    if not pids:
        modern_showinfo("Nothing to push", "No pending or failed patients found.")
        return
    if push_job["running"]:
        modern_showwarning("Busy", "A push job is already running.")
        return
    anon = anon_var.get()
    if push_multi_dest_toggle_var.get():
        dests = _get_selected_multi_destinations()
        if not dests:
            modern_showwarning("No Destinations", "Check at least one destination to push to.")
            return
        if not _confirm_large_push_if_needed(pids, dest_count=len(dests), dests=dests):
            return
        threading.Thread(target=run_push_job_multi, args=(pids, dests),
                          kwargs={"anonymize": anon}, daemon=True).start()
    else:
        dest = _get_active_dest()
        if not _confirm_large_push_if_needed(pids, dests=[dest] if dest else None):
            return
        threading.Thread(target=run_push_job, args=(pids,), kwargs={"destination": dest, "anonymize": anon}, daemon=True).start()


def do_echo_active_dest():
    dest = _get_active_dest()
    if not dest:
        modern_showerror("Error", "No destination configured. Add one in the Destinations tab.")
        return
    echo_btn.configure(state="disabled", text="Testing...")

    def run():
        ok, msg = dicom_echo(dest["ae"], dest["ip"], dest["port"],
                             calling_ae=dest.get("calling_ae"), dest_label=dest.get("name"))
        def on_ui():
            echo_btn.configure(state="normal", text="C-ECHO Active Dest.")
            if ok:
                modern_showinfo("C-ECHO Success", msg)
            else:
                modern_showerror("C-ECHO Failed", msg)
        app.after(0, on_ui)

    threading.Thread(target=run, daemon=True).start()


push_all_btn.configure(command=do_push_selected)
push_everything_btn.configure(command=do_push_all_pending)
push_stop_btn.configure(command=stop_push_job)
echo_btn.configure(command=do_echo_active_dest)
push_refresh_btn.configure(command=lambda: refresh_worklists(push_search_var.get()))
push_search_var.trace_add("write", lambda *_: refresh_worklists(push_search_var.get()))
push_status_filter_var.trace_add("write", lambda *_: populate_tree(push_tree, push_search_var.get(), force=True))
push_date_filter_var.trace_add("write", lambda *_: populate_tree(push_tree, push_search_var.get(), force=True))
push_report_filter_var.trace_add("write", lambda *_: populate_tree(push_tree, push_search_var.get(), force=True))
push_dest_filter_var.trace_add("write", lambda *_: populate_tree(push_tree, push_search_var.get(), force=True))
push_export_btn.configure(command=lambda: export_worklist_view_to_csv(push_tree))

# =========================================================
# RIGHT-CLICK CONTEXT MENU (both trees)
# =========================================================

def open_in_viewer(pid):
    folder = get_patient_folder(pid)
    if not os.path.isdir(folder):
        modern_showwarning("Not found", f"No local files for {pid}.")
        return
    if sys.platform.startswith("win"):
        os.startfile(folder)
    elif sys.platform.startswith("darwin"):
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])


def show_study_grouping(pid):
    """Show a breakdown of all locally stored studies for a patient."""
    studies = get_studies_for_patient(pid)
    if not studies:
        modern_showinfo("No Studies", f"No local DICOM files found for {pid}.")
        return

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title(f"Studies for {pid}")
    win.geometry("700x350")
    _safe_grab_set(win)
    win.protocol("WM_DELETE_WINDOW", lambda: (win.grab_release(), win.destroy()))

    cols = ("study_uid", "modality", "date", "count")
    hdgs = {"study_uid": "Study UID", "modality": "Modality", "date": "Study Date", "count": "#Images"}
    widths = {"study_uid": 320, "modality": 80, "date": 100, "count": 70}

    tree_f = ctk.CTkFrame(win, fg_color=THEME_SURFACE)
    tree_f.pack(fill="both", expand=True, padx=10, pady=10)
    sv = ttk.Scrollbar(tree_f, orient="vertical")
    t = ttk.Treeview(tree_f, columns=cols, show="headings", yscrollcommand=sv.set)
    sv.config(command=t.yview)
    for col in cols:
        t.heading(col, text=hdgs[col])
        t.column(col, width=widths[col], anchor="w")
    for uid, info in studies.items():
        t.insert("", "end", values=(uid, info.get("modality", ""), info.get("date", ""), info["count"]))
    sv.pack(side="right", fill="y")
    t.pack(fill="both", expand=True)


def _worklist_column_at(tree, event):
    """Returns the WL_COLUMNS name under the given click event, or None.
    IMPORTANT: tree.identify_column() returns an index relative to the
    tree's CURRENTLY DISPLAYED columns (tree["displaycolumns"]), not the
    full WL_COLUMNS tuple. If any column has been hidden via the column-
    visibility toggle (see _column_visibility / _apply_column_visibility),
    those two orderings diverge and indexing straight into WL_COLUMNS
    silently returns the wrong column name -- e.g. a click on "report"
    could get treated as a click on "documents" instead. Resolving
    against the tree's own displaycolumns keeps this correct regardless
    of which columns are currently hidden."""
    if tree.identify_region(event.x, event.y) != "cell":
        return None
    col_id = tree.identify_column(event.x)  # e.g. '#6'
    try:
        col_index = int(col_id.replace("#", "")) - 1
        displayed = tree.cget("displaycolumns")
        if not displayed or displayed == "#all" or tuple(displayed) == ("#all",):
            displayed = WL_COLUMNS
        else:
            displayed = tuple(str(c) for c in displayed)
        return displayed[col_index]
    except (ValueError, IndexError):
        return None


WEASIS_PATH_FILE = "weasis_path.txt"

WEASIS_EXECUTABLE_CANDIDATES = {
    "Windows": [
        r"C:\Program Files\Weasis\Weasis.exe",
        r"C:\Program Files (x86)\Weasis\Weasis.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Weasis\Weasis.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Weasis\Weasis.exe"),
        os.path.expandvars(r"%APPDATA%\Weasis\Weasis.exe"),
    ],
    "Darwin": [
        "/Applications/Weasis.app/Contents/MacOS/Weasis",
        os.path.expanduser("~/Applications/Weasis.app/Contents/MacOS/Weasis"),
    ],
    "Linux": [
        "/usr/bin/weasis",
        "/usr/local/bin/weasis",
        os.path.expanduser("~/weasis/weasis"),
        os.path.expanduser("~/.local/share/weasis/weasis"),
    ],
}


def _load_saved_weasis_path():
    if os.path.isfile(WEASIS_PATH_FILE):
        try:
            with open(WEASIS_PATH_FILE, "r", encoding="utf-8") as f:
                path = f.read().strip()
            if path and os.path.isfile(path):
                return path
        except Exception:
            log_exception("Failed to read weasis_path.txt")
    return None


def _save_weasis_path(path):
    try:
        atomic_write(WEASIS_PATH_FILE, path)
    except Exception:
        log_exception("Failed to save weasis_path.txt")


def find_weasis_executable():
    """Locate the Weasis launcher: a previously remembered path, then
    PATH, then the common per-OS install locations."""
    saved = _load_saved_weasis_path()
    if saved:
        return saved
    found = shutil.which("weasis") or shutil.which("Weasis")
    if found:
        return found
    for candidate in WEASIS_EXECUTABLE_CANDIDATES.get(platform.system(), []):
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


# ---- RadiAnt (mirrors the Weasis discovery block above) -----------------

RADIANT_PATH_FILE = "radiant_path.txt"

RADIANT_EXECUTABLE_CANDIDATES = {
    # RadiAnt is Windows-only; other platforms simply won't find a match
    # here and the user will be prompted to locate it (or told it's
    # unavailable, same as any not-installed viewer).
    "Windows": [
        r"C:\Program Files\RadiAnt DICOM Viewer\RadiAntViewer.exe",
        r"C:\Program Files (x86)\RadiAnt DICOM Viewer\RadiAntViewer.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\RadiAnt DICOM Viewer\RadiAntViewer.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\RadiAnt DICOM Viewer\RadiAntViewer.exe"),
    ],
    "Darwin": [],
    "Linux": [],
}


def _load_saved_radiant_path():
    if os.path.isfile(RADIANT_PATH_FILE):
        try:
            with open(RADIANT_PATH_FILE, "r", encoding="utf-8") as f:
                path = f.read().strip()
            if path and os.path.isfile(path):
                return path
        except Exception:
            log_exception("Failed to read radiant_path.txt")
    return None


def _save_radiant_path(path):
    try:
        atomic_write(RADIANT_PATH_FILE, path)
    except Exception:
        log_exception("Failed to save radiant_path.txt")


def find_radiant_executable():
    """Locate the RadiAnt launcher: a previously remembered path, then
    PATH, then the common Windows install locations."""
    saved = _load_saved_radiant_path()
    if saved:
        return saved
    found = shutil.which("RadiAntViewer") or shutil.which("radiant")
    if found:
        return found
    for candidate in RADIANT_EXECUTABLE_CANDIDATES.get(platform.system(), []):
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


VIEWER_WEASIS = "weasis"
VIEWER_RADIANT = "radiant"
VIEWER_PREF_FILE = "viewer_pref.txt"

VIEWER_REGISTRY = {
    VIEWER_WEASIS: {
        "label": "Weasis",
        "find": find_weasis_executable,
        "save_path": _save_weasis_path,
    },
    VIEWER_RADIANT: {
        "label": "RadiAnt",
        "find": find_radiant_executable,
        "save_path": _save_radiant_path,
    },
}


def _load_remembered_viewer():
    if os.path.isfile(VIEWER_PREF_FILE):
        try:
            with open(VIEWER_PREF_FILE, "r", encoding="utf-8") as f:
                choice = f.read().strip()
            if choice in VIEWER_REGISTRY:
                return choice
        except Exception:
            log_exception("Failed to read viewer_pref.txt")
    return None


def _save_remembered_viewer(choice):
    try:
        atomic_write(VIEWER_PREF_FILE, choice)
    except Exception:
        log_exception("Failed to save viewer_pref.txt")


def _patient_dcm_folder_and_files(pid):
    """Shared lookup used by every viewer launcher: returns (folder,
    dcm_files) or (None, None) after showing the appropriate error, so
    each viewer doesn't repeat the same folder/empty-folder checks."""
    folder = get_patient_folder(pid)
    if not os.path.isdir(folder):
        modern_showerror("Open in Viewer", f"No folder found for patient {pid}.")
        return None, None
    dcm_files = [
        os.path.join(folder, fname)
        for fname in os.listdir(folder)
        if fname.lower().endswith(".dcm")
    ]
    if not dcm_files:
        modern_showwarning("Open in Viewer", f"No DICOM images found for patient {pid}.")
        return None, None
    return folder, dcm_files


def open_patient_in_viewer(pid, viewer):
    """Launches the chosen viewer (Weasis or RadiAnt) directly with every
    DICOM image in this patient's received folder — no 'open with'
    prompt. viewer must be VIEWER_WEASIS or VIEWER_RADIANT."""
    entry = VIEWER_REGISTRY.get(viewer)
    if not entry:
        return
    folder, dcm_files = _patient_dcm_folder_and_files(pid)
    if not folder:
        return

    exe = entry["find"]()
    if not exe:
        modern_showinfo(
            f"Locate {entry['label']}",
            f"{entry['label']} wasn't found automatically. Please locate the "
            f"{entry['label']} application/executable once — this will be "
            f"remembered for next time.")
        chosen = filedialog.askopenfilename(title=f"Locate {entry['label']}")
        if not chosen:
            return
        exe = chosen
        entry["save_path"](exe)

    try:
        # Pass the whole patient folder rather than every individual file:
        # this scales cleanly even when a patient has thousands of images,
        # and both viewers will scan the folder for DICOM content on launch.
        subprocess.Popen([exe, folder])
    except Exception as e:
        log_exception(f"Failed to launch {entry['label']} for {pid}")
        modern_showerror(
            f"Could Not Open {entry['label']}",
            f"Failed to launch {entry['label']}: {e}\n\n"
            f"If {entry['label']} has moved, delete '{VIEWER_PATH_FILE_FOR(viewer)}' "
            f"next to this app and try again to relocate it.")


def VIEWER_PATH_FILE_FOR(viewer):
    return WEASIS_PATH_FILE if viewer == VIEWER_WEASIS else RADIANT_PATH_FILE


def open_patient_in_weasis(pid):
    """Back-compat wrapper: existing call sites (context menu, double/
    single-click handlers) can keep calling this name directly for the
    'always Weasis' path; the new picker dialog is a separate entry point."""
    open_patient_in_viewer(pid, VIEWER_WEASIS)


def open_patient_in_radiant(pid):
    open_patient_in_viewer(pid, VIEWER_RADIANT)


def _open_selected_in_viewer(tree):
    """Used by the 'Open in Viewer' toolbar button on both the Receiver
    and Pusher tabs: acts on whichever worklist row is currently
    selected, same as clicking the row's Images cell."""
    selected = tree.selection()
    if not selected:
        modern_showinfo("Open in Viewer", "Select a patient row first.")
        return
    open_patient_in_viewer_picker(selected[0])


def open_patient_in_viewer_picker(pid):
    """Shows a small dialog with two real buttons -- 'Open in Weasis' and
    'Open in RadiAnt' -- so the user picks which viewer launches this
    patient's images. An optional 'Remember my choice' checkbox skips the
    dialog on future clicks until changed via 'Change default viewer'."""
    remembered = _load_remembered_viewer()
    if remembered:
        open_patient_in_viewer(pid, remembered)
        return

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title("Open in Viewer")
    win.geometry("340x230")
    win.transient(app)
    _safe_grab_set(win)

    ctk.CTkLabel(win, text=f"Open patient {pid} in:",
                 font=get_font("caption", "bold")).pack(pady=(20, 14))

    remember_var = ctk.BooleanVar(value=False)

    def choose(viewer):
        if remember_var.get():
            _save_remembered_viewer(viewer)
        win.destroy()
        open_patient_in_viewer(pid, viewer)

    ctk.CTkButton(win, text="Open in Weasis", width=240,
                  command=lambda: choose(VIEWER_WEASIS)).pack(pady=6)
    ctk.CTkButton(win, text="Open in RadiAnt", width=240,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: choose(VIEWER_RADIANT)).pack(pady=6)

    ctk.CTkCheckBox(win, text="Remember my choice", variable=remember_var).pack(pady=(14, 4))
    ctk.CTkButton(win, text="Cancel", width=100, fg_color="transparent",
                  border_width=1, command=win.destroy).pack(pady=(6, 0))


def show_full_error_dialog(pid, message):
    """Show the full, untruncated last-error message for a patient in a
    selectable/copyable popup — the worklist column is fixed-width and
    will visually truncate long diagnostic messages (e.g. the detailed
    association-rejection reasons), so this is how you read/copy the
    whole thing without digging through the Logs tab."""
    import tkinter as tk
    win = tk.Toplevel(app)
    _keep_toplevel_small(win)
    win.title(f"Last Error — {pid}")
    win.geometry("560x260")
    win.transient(app)

    box = ctk.CTkTextbox(win, wrap="word", font=get_font("caption"))
    box.pack(fill="both", expand=True, padx=12, pady=(12, 6))
    box.insert("1.0", message)
    box.configure(state="disabled")

    def copy_to_clipboard():
        app.clipboard_clear()
        app.clipboard_append(message)

    btn_row = ctk.CTkFrame(win, fg_color="transparent")
    btn_row.pack(fill="x", padx=12, pady=(0, 12))
    ctk.CTkButton(btn_row, text="Copy", width=100, command=copy_to_clipboard).pack(side="left")
    ctk.CTkButton(btn_row, text="Close", width=100, command=win.destroy).pack(side="right")


def _on_worklist_double_click(tree, event):
    """Double-clicking the Images cell opens the viewer picker (Weasis /
    RadiAnt) with all of this patient's images. Double-clicking the Report
    cell opens/creates that patient's report. Double-clicking anywhere
    else on the row opens the viewer picker too, EXCEPT when the row has
    a failed/retrying status with an error message, in which case that
    takes priority (jumps to the full error dialog instead, which is more
    useful for a row in that state)."""
    row_id = tree.identify_row(event.y)
    if not row_id:
        return
    col = _worklist_column_at(tree, event)
    if col == "documents":
        open_patient_in_viewer_picker(row_id)
        return
    if col == "report":
        open_report_action(row_id)
        return
    if col == "history":
        open_history_action(row_id)
        return
    with data_lock:
        info = patient_data.get(row_id, {})
        last_error = info.get("last_error", "")
        status = info.get("status", "")
    if last_error and status in (STATUS_FAILED, STATUS_RETRYING):
        show_full_error_dialog(row_id, last_error)
        return
    open_patient_in_viewer_picker(row_id)


def _on_worklist_single_click(tree, event):
    """Single-clicking the Images cell opens the viewer picker (Weasis /
    RadiAnt); single-clicking the Report cell opens/creates that
    patient's report."""
    row_id = tree.identify_row(event.y)
    if not row_id:
        return
    col = _worklist_column_at(tree, event)
    if col == "documents":
        open_patient_in_viewer_picker(row_id)
    elif col == "report":
        open_report_action(row_id)
    elif col == "history":
        open_history_action(row_id)


rec_tree.bind("<Double-1>", lambda e: _on_worklist_double_click(rec_tree, e))
push_tree.bind("<Double-1>", lambda e: _on_worklist_double_click(push_tree, e))
rec_tree.bind("<Button-1>", lambda e: _on_worklist_single_click(rec_tree, e), add="+")
push_tree.bind("<Button-1>", lambda e: _on_worklist_single_click(push_tree, e), add="+")
rec_tree.bind("<Button-1>", lambda e: _on_worklist_checkbox_click(rec_tree, e), add="+")
push_tree.bind("<Button-1>", lambda e: _on_worklist_checkbox_click(push_tree, e), add="+")
rec_tree.bind("<<TreeviewSelect>>", lambda _e: _refresh_row_checkboxes(rec_tree), add="+")
push_tree.bind("<<TreeviewSelect>>", lambda _e: _refresh_row_checkboxes(push_tree), add="+")


def open_report_action(pid):
    """UI-layer wrapper around open_report(): handles the one interactive
    decision point (a corrupt existing report) with an explicit
    confirmation, and turns every other outcome into a friendly
    messagebox rather than a silent failure."""
    status, result = open_report(pid)

    if status == "ok":
        return  # opened successfully, nothing further to show

    if status == "corrupt":
        path = result
        proceed = modern_askyesno(
            "Report Appears Corrupted",
            f"The existing report for {pid} could not be read (it may be damaged or "
            f"incomplete):\n\n{path}\n\n"
            f"Your original file will be kept as a timestamped backup in the same "
            f"folder — it will NOT be deleted. A fresh report template will be created "
            f"so you can continue working.\n\nProceed?",
        )
        if not proceed:
            return
        ok, path_or_err = recover_corrupt_report(pid)
        if not ok:
            modern_showerror("Report Recovery Failed", path_or_err)
            return
        ok2, err2 = open_document(path_or_err)
        if not ok2:
            modern_showerror("Could Not Open Report", err2)
        else:
            set_fields(pid, report_last_opened=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        return

    # status == "error"
    hint = ""
    if "not find an application" in result.lower() or "could not open" in result.lower():
        hint = "\n\nIf Microsoft Word (or another .docx-capable app) isn't installed, install one and try again."
    modern_showerror("Could Not Open Report", result + hint)


def open_history_action(pid):
    """UI-layer wrapper around open_history(): creates the patient's
    Patient_History.txt on first use (with a small identifying header),
    then opens it with whatever text editor the OS has associated with
    .txt files -- every subsequent click just opens the same file, never
    overwrites it."""
    ok, result = open_history(pid)
    if ok:
        return  # opened successfully, nothing further to show

    hint = ""
    if "not find an application" in result.lower() or "could not open" in result.lower():
        hint = "\n\nIf no text editor is associated with .txt files, install/associate one and try again."
    modern_showerror("Could Not Open History", result + hint)


# =========================================================
# SOFT DELETE / UNDO  (1.2)
# =========================================================
# _bulk_delete_selected and the right-click "Delete" both used to
# shutil.rmtree() immediately after modern_confirm -- irreversible the
# instant the operator clicked. This stages deleted folders into
# OUTPUT_DIR/.trash/<pid>_<ts>/ instead, offers a short-lived in-app Undo
# toast, and only permanently purges .trash entries older than
# TRASH_RETENTION_MINUTES via a periodic sweep (same throttled-check
# pattern as rotate_logs_if_needed()).

TRASH_RETENTION_MINUTES = 10
_trash_manifest = {}       # trash_id -> {"pid", "trash_path", "data", "deleted_at"}
_trash_sweep_state = {"last": 0}
_undo_toast_state = {"widget": None}


def _trash_dir():
    d = os.path.join(OUTPUT_DIR, ".trash")
    os.makedirs(d, exist_ok=True)
    return d


def _stage_delete_to_trash(pid, patient_record):
    """Moves pid's folder into .trash instead of deleting it outright.
    Returns a trash_id that _restore_from_trash() can use to undo this
    specific deletion. Falls back to a real delete only if the move itself
    fails (e.g. cross-device edge case)."""
    folder = get_patient_folder(pid)
    trash_id = f"{_filename_safe_pid(pid)}_{int(time.time() * 1000)}"
    trash_path = os.path.join(_trash_dir(), trash_id)
    moved = False
    if os.path.isdir(folder):
        try:
            shutil.move(folder, trash_path)
            moved = True
        except Exception:
            log_exception(f"Could not stage {pid} in trash; deleting outright")
            shutil.rmtree(folder, ignore_errors=True)
    _trash_manifest[trash_id] = {
        "pid": pid,
        "trash_path": trash_path if moved else None,
        "data": patient_record,
        "deleted_at": time.time(),
    }
    return trash_id


def _restore_from_trash(trash_id):
    entry = _trash_manifest.pop(trash_id, None)
    if entry is None:
        return False
    pid = entry["pid"]
    trash_path = entry.get("trash_path")
    if trash_path and os.path.isdir(trash_path):
        dest = get_patient_folder(pid)
        try:
            if os.path.isdir(dest):
                shutil.rmtree(dest, ignore_errors=True)
            shutil.move(trash_path, dest)
        except Exception:
            log_exception(f"Failed to restore {pid} from trash")
            return False
    if entry.get("data") is not None:
        with data_lock:
            patient_data[pid] = entry["data"]
    write_audit_log("RESTORE", f"pid={pid} restored from trash (undo)")
    autosave_csv()
    _bump_data_version()
    refresh_worklists()
    return True


def sweep_trash_folder():
    """Periodic (called from periodic_refresh): permanently purge .trash
    entries older than TRASH_RETENTION_MINUTES. Self-throttled to once a
    minute so it's cheap to call on every refresh tick."""
    now = time.time()
    if now - _trash_sweep_state["last"] < 60:
        return
    _trash_sweep_state["last"] = now
    cutoff = now - TRASH_RETENTION_MINUTES * 60
    for trash_id in list(_trash_manifest.keys()):
        entry = _trash_manifest.get(trash_id)
        if entry and entry["deleted_at"] < cutoff:
            trash_path = entry.get("trash_path")
            if trash_path and os.path.isdir(trash_path):
                shutil.rmtree(trash_path, ignore_errors=True)
            _trash_manifest.pop(trash_id, None)
    # Also sweep any orphaned trash folders left over from a prior run
    # (e.g. app closed before their entry aged out of the in-memory manifest).
    try:
        td = _trash_dir()
        for name in os.listdir(td):
            full = os.path.join(td, name)
            if os.path.isdir(full) and os.path.getmtime(full) < cutoff:
                shutil.rmtree(full, ignore_errors=True)
    except Exception:
        log_exception("Failed to sweep orphaned trash folders")


def _show_undo_toast(message, on_undo, seconds=10):
    """A real in-app toast (distinct from the OS-level show_toast_threadsafe)
    with a clickable Undo button, auto-dismissing after `seconds`. Only one
    is shown at a time -- a newer delete's toast replaces the last."""
    old = _undo_toast_state.get("widget")
    if old is not None:
        try:
            old.destroy()
        except Exception:
            pass
        _undo_toast_state["widget"] = None

    toast = ctk.CTkFrame(app, fg_color=THEME_HEADING_BG, corner_radius=10)
    toast.place(relx=0.5, rely=0.94, anchor="center")
    ctk.CTkLabel(toast, text=message, font=get_font("small"), text_color=THEME_TEXT).pack(
        side="left", padx=(16, 10), pady=10)

    def _dismiss():
        if _undo_toast_state.get("widget") is toast:
            _undo_toast_state["widget"] = None
        try:
            toast.destroy()
        except Exception:
            pass

    def _do_undo():
        _dismiss()
        on_undo()

    ctk.CTkButton(toast, text="Undo", width=80, height=28, corner_radius=8,
                  fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
                  command=_do_undo).pack(side="left", padx=(0, 16), pady=10)
    _undo_toast_state["widget"] = toast
    app.after(seconds * 1000, _dismiss)


def _delete_pids_with_undo(pids):
    """Shared soft-delete path used by both the right-click 'Delete' menu
    item and the bulk toolbar's 'Delete Selected' -- stages folders into
    .trash and offers one Undo toast covering the whole batch."""
    if not pids:
        return
    with data_lock:
        records = {pid: patient_data.pop(pid, None) for pid in pids if pid in patient_data}
    if not records:
        return
    trash_ids = []
    for pid, record in records.items():
        trash_ids.append(_stage_delete_to_trash(pid, record))
        write_audit_log("DELETE", f"pid={pid} staged to trash (undo available {TRASH_RETENTION_MINUTES}m)")
    autosave_csv()
    _bump_data_version()
    refresh_worklists()

    def _undo_all():
        for trash_id in trash_ids:
            _restore_from_trash(trash_id)

    n = len(records)
    _show_undo_toast(f"Deleted {n} patient{'s' if n != 1 else ''}", _undo_all)


def build_context_menu(tree, event):
    selected = tree.selection()
    if not selected:
        return
    pid = selected[0]

    import tkinter as tk
    menu = tk.Menu(app, tearoff=0)

    if len(selected) > 1:
        # Multiple rows selected -- offer bulk actions across the whole
        # selection instead of only acting on the row under the cursor.
        pids = list(selected)
        n = len(pids)
        menu.add_command(label=f"{n} studies selected", state="disabled")
        menu.add_separator()
        menu.add_command(
            label=f"Push {n} studies to active destination",
            command=lambda: threading.Thread(
                target=run_push_job, args=(pids,),
                kwargs={"destination": _get_active_dest(), "anonymize": anon_var.get()},
                daemon=True).start())

        dests = load_destinations()
        if dests:
            dest_menu = tk.Menu(menu, tearoff=0)
            for d in dests:
                def push_to_many(dest=d, pids=pids):
                    threading.Thread(
                        target=run_push_job, args=(pids,),
                        kwargs={"destination": dest, "anonymize": anon_var.get()},
                        daemon=True).start()
                dest_menu.add_command(label=f"{d['name']} ({d['ae']}@{d['ip']}:{d['port']})", command=push_to_many)
            menu.add_cascade(label=f"Push {n} studies to...", menu=dest_menu)
        menu.add_separator()
        menu.add_command(label=f"Reset {n} studies to Pending",
                         command=lambda: [reset_patient_status(p) for p in pids])
        menu.add_separator()

        def delete_many():
            if modern_askyesno("Delete", f"Delete {n} selected studies from worklist AND disk?\nYou'll have a few seconds to undo."):
                _delete_pids_with_undo(pids)

        menu.add_command(label=f"Delete {n} studies from worklist + disk", command=delete_many)

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return

    menu.add_command(label=f"Open local folder  [{pid}]", command=lambda: open_in_viewer(pid))
    menu.add_command(label="Show study grouping", command=lambda: show_study_grouping(pid))
    menu.add_command(label="Open in Viewer...", command=lambda: open_patient_in_viewer_picker(pid))
    menu.add_command(label="    ↳  Weasis directly", command=lambda: open_patient_in_weasis(pid))
    menu.add_command(label="    ↳  RadiAnt directly", command=lambda: open_patient_in_radiant(pid))
    if _load_remembered_viewer():
        menu.add_command(
            label="    ↳  Forget remembered viewer choice",
            command=lambda: (_save_remembered_viewer(""), None))
    menu.add_command(label="Open Report", command=lambda: open_report_action(pid))
    menu.add_command(label="Open History", command=lambda: open_history_action(pid))
    menu.add_separator()
    menu.add_command(label="Push to active destination",
                     command=lambda: threading.Thread(
                         target=run_push_job, args=([pid],),
                         kwargs={"destination": _get_active_dest(), "anonymize": anon_var.get()},
                         daemon=True).start())
    menu.add_separator()

    dests = load_destinations()
    if dests:
        dest_menu = tk.Menu(menu, tearoff=0)
        for d in dests:
            def push_to(dest=d):
                # Goes straight to push_single_patient rather than
                # run_push_job -- this is a one-row action, not a job, so
                # it shouldn't be silently dropped by run_push_job's
                # "a job is already running" guard, and it bypasses the
                # global push_dest_var entirely (2.2).
                threading.Thread(
                    target=push_single_patient, args=(pid,),
                    kwargs={"destination": dest, "anonymize": anon_var.get()},
                    daemon=True).start()
            dest_menu.add_command(label=f"{d['name']} ({d['ae']}@{d['ip']}:{d['port']})", command=push_to)
        menu.add_cascade(label="Push this study to...", menu=dest_menu)
        menu.add_separator()

    if tree is push_tree:
        # C.2: manual "resend documents only" action (A.7) -- Pusher side
        # only, since the Receiver tab is the delivery target, not the
        # source of Report/History files.
        has_docs = os.path.isfile(get_report_path(pid)) or os.path.isfile(get_history_path(pid))
        menu.add_command(
            label="Resend Report/History Only",
            command=lambda: do_resend_patient_documents(pid),
            state="normal" if has_docs else "disabled",
        )
        menu.add_separator()

    last_error = ""
    with data_lock:
        last_error = patient_data.get(pid, {}).get("last_error", "")
    if last_error:
        menu.add_command(label="View full error message...",
                         command=lambda: show_full_error_dialog(pid, last_error))
        menu.add_separator()

    menu.add_command(label="Reset status to Pending",
                     command=lambda: reset_patient_status(pid))
    menu.add_separator()

    def _copy_to_clipboard(text):
        app.clipboard_clear()
        app.clipboard_append(text)

    study_uid = ""
    with data_lock:
        study_uid = patient_data.get(pid, {}).get("study_uid", "")
    menu.add_command(label="Copy Patient ID", command=lambda: _copy_to_clipboard(pid))
    menu.add_command(label="Copy Study UID", command=lambda: _copy_to_clipboard(study_uid),
                     state="normal" if study_uid else "disabled")
    menu.add_separator()

    def delete_patient():
        if modern_askyesno("Delete", f"Delete {pid} from worklist AND disk?\nYou'll have a few seconds to undo."):
            _delete_pids_with_undo([pid])

    menu.add_command(label="Delete from worklist + disk", command=delete_patient)

    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


rec_tree.bind("<Button-3>", lambda e: build_context_menu(rec_tree, e))
push_tree.bind("<Button-3>", lambda e: build_context_menu(push_tree, e))
rec_tree.bind("<Button-2>", lambda e: build_context_menu(rec_tree, e))   # macOS
push_tree.bind("<Button-2>", lambda e: build_context_menu(push_tree, e))

# Keyboard shortcuts: Ctrl+A selects every visible row, Escape clears the
# selection -- standard multi-select affordances users expect from any
# list/table view.
for _wl_tree in (rec_tree, push_tree):
    _wl_tree.bind("<Control-a>", lambda e, t=_wl_tree: (t.selection_set(t.get_children("")), "break")[1])
    _wl_tree.bind("<Control-A>", lambda e, t=_wl_tree: (t.selection_set(t.get_children("")), "break")[1])
    _wl_tree.bind("<Escape>", lambda e, t=_wl_tree: (_clear_persistent_selection(t), "break")[1])


# =========================================================
# CONTEXTUAL SELECTION TOOLBAR
# =========================================================
# A slim action bar that appears above a worklist tree only while rows
# are selected, and disappears otherwise -- the "toolbar transforms into
# contextual actions on selection" pattern, built on the same
# run_push_job / delete / reset actions the right-click menu already
# uses (no new business logic, just a faster multi-row entry point to it).

_contextual_toolbars = {}  # id(tree) -> CTkFrame


def _effective_selection(tree):
    """The true selected-PID set for `tree`: for paginated trees this is
    the cross-page persistent set (see 1.1), since tree.selection() only
    ever reflects rows currently rendered on the visible page."""
    if id(tree) in _persistent_selection:
        return list(_get_persistent_selection(tree))
    return list(tree.selection())


def _bulk_push_selected(tree):
    pids = _effective_selection(tree)
    if not pids:
        return
    threading.Thread(
        target=run_push_job, args=(pids,),
        kwargs={"destination": _get_active_dest(), "anonymize": anon_var.get()},
        daemon=True).start()


def _bulk_reset_selected(tree):
    for pid in _effective_selection(tree):
        reset_patient_status(pid)


def _bulk_copy_ids_selected(tree):
    pids = _effective_selection(tree)
    if not pids:
        return
    app.clipboard_clear()
    app.clipboard_append("\n".join(pids))


def _bulk_delete_selected(tree):
    pids = _effective_selection(tree)
    if not pids:
        return
    if not modern_confirm(
        "Delete Selected", f"Delete {len(pids)} patient(s) from worklist AND disk?\nYou'll have a few seconds to undo.",
        danger=True,
    ):
        return
    _delete_pids_with_undo(pids)
    _get_persistent_selection(tree).difference_update(pids)


def attach_contextual_toolbar(tree, wl_frame):
    """Packs a hidden-by-default action bar directly above wl_frame, and
    wires <<TreeviewSelect>> to show/hide it and keep its count current."""
    bar = ctk.CTkFrame(wl_frame.master, fg_color=THEME_HEADING_BG, corner_radius=10, height=40)
    bar.pack_propagate(False)

    count_lbl = ctk.CTkLabel(bar, text="", font=get_font("small", "bold"), text_color=THEME_TEXT)
    count_lbl.pack(side="left", padx=(14, 10))

    ctk.CTkButton(bar, text="Push Selected", width=145, height=28, corner_radius=8,
                  image=get_icon("send", size=14, color="#ffffff"), compound="left",
                  fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
                  command=lambda: _bulk_push_selected(tree)).pack(side="left", padx=4)
    ctk.CTkButton(bar, text="Reset to Pending", width=170, height=28, corner_radius=8,
                  image=get_icon("rotate-ccw", size=14, color=THEME_TEXT), compound="left",
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: _bulk_reset_selected(tree)).pack(side="left", padx=4)
    ctk.CTkButton(bar, text="Copy IDs", width=115, height=28, corner_radius=8,
                  image=get_icon("layers", size=14, color=THEME_TEXT), compound="left",
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: _bulk_copy_ids_selected(tree)).pack(side="left", padx=4)
    ctk.CTkButton(bar, text="Delete Selected", width=160, height=28, corner_radius=8,
                  image=get_icon("trash-2", size=14, color="#ffffff"), compound="left",
                  fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER,
                  command=lambda: _bulk_delete_selected(tree)).pack(side="left", padx=4)

    def _set_count(n):
        if n > 0:
            count_lbl.configure(text=f"{n} selected")
            if not bar.winfo_ismapped():
                bar.pack(fill="x", padx=10, pady=(0, 8), before=wl_frame)
        else:
            bar.pack_forget()

    def _on_select(_evt=None):
        _set_count(len(_effective_selection(tree)))

    tree.bind("<<TreeviewSelect>>", _on_select, add="+")
    _selection_count_hooks[id(tree)] = _set_count
    _contextual_toolbars[id(tree)] = bar


attach_contextual_toolbar(rec_tree, rec_wl_frame)
attach_contextual_toolbar(push_tree, push_wl_frame)

# =========================================================
# RIGHT INSPECTOR PANEL
# =========================================================
# Slides in from the right edge of content_row when a worklist row is
# selected, showing Patient/Study/Transfer info without opening a popup.
# This is the closest practical equivalent to a full IDE-style inspector
# dock in CustomTkinter/Tkinter: it's a real panel (not a dialog), it
# animates open/closed, and its content refreshes live -- it just isn't
# independently draggable/floatable the way a native docking framework
# would allow. Reuses existing read-only accessors (get_studies_for_patient,
# get_checkpoint_info, get_report_path, get_history_path) -- no new state.

inspector_frame = ctk.CTkFrame(content_row, fg_color=THEME_SURFACE, corner_radius=12, width=0)
inspector_frame.pack(side="right", fill="y", padx=(10, 0))
inspector_frame.pack_propagate(False)

inspector_header = ctk.CTkFrame(inspector_frame, fg_color="transparent")
inspector_header.pack(fill="x", padx=16, pady=(14, 6))
ctk.CTkLabel(inspector_header, text="Inspector", font=get_font("section", "bold"),
             text_color=THEME_TEXT).pack(side="left")
inspector_close_btn = ctk.CTkButton(
    inspector_header, text="", width=26, height=26, corner_radius=8,
    image=get_icon("x", size=14, color=THEME_TEXT_MUTED),
    fg_color="transparent", hover_color=THEME_HEADING_BG, text_color=THEME_TEXT_MUTED,
)
inspector_close_btn.pack(side="right")

inspector_scroll = ctk.CTkScrollableFrame(inspector_frame, fg_color="transparent")
inspector_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 12))

inspector_current_pid = {"value": None}


def _inspector_row(parent, label, value):
    row = ctk.CTkFrame(parent, fg_color="transparent")
    row.pack(fill="x", pady=2)
    ctk.CTkLabel(row, text=label, font=get_font("micro", "bold"), text_color=THEME_TEXT_MUTED,
                 width=100, anchor="w").pack(side="left")
    ctk.CTkLabel(row, text=str(value) if value not in (None, "") else "—", font=get_font("small"),
                 text_color=THEME_TEXT, anchor="w", justify="left", wraplength=170).pack(side="left", fill="x", expand=True)


def _inspector_section(title):
    card = make_card(inspector_scroll, title=title)
    card.pack(fill="x", pady=(0, 10))
    return card.body


def _animate_inspector(target_width, step=0):
    # Enterprise UI requirement: no animations/transitions anywhere.
    # Snap directly to the target width instead of tweening.
    try:
        inspector_frame.configure(width=max(target_width, 0))
    except Exception:
        return
    inspector_state["animating"] = False


def open_inspector():
    if not inspector_state["visible"]:
        inspector_state["visible"] = True
        inspector_state["animating"] = True
        _animate_inspector(INSPECTOR_WIDTH)


def close_inspector():
    inspector_current_pid["value"] = None
    if inspector_state["visible"]:
        inspector_state["visible"] = False
        inspector_state["animating"] = True
        _animate_inspector(0)


inspector_close_btn.configure(command=close_inspector)


def _human_size(n):
    n = float(n or 0)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def _finish_add_attachments(pid, paths, batch_label):
    ok_count, errors = add_patient_attachments(pid, list(paths))
    if inspector_current_pid["value"] == pid:
        refresh_inspector(pid)
    if errors:
        detail = "\n".join(f"{os.path.basename(p)}: {e}" for p, e in errors[:8])
        modern_showwarning(
            "Some Attachments Failed",
            f"{ok_count} of {len(paths)} {batch_label} added for {pid}.\n\n{detail}")
    elif ok_count:
        ui_event_queue.put(("toast", ("Attachments Added", f"{ok_count} file(s) added for {pid}.")))


def do_add_attachments(pid):
    """Opens a native multi-select file picker (unlimited files in one
    go) and copies everything chosen into this patient's Attachments
    folder. Reports partial failures instead of silently dropping files
    from a large batch."""
    paths = filedialog.askopenfilenames(title=f"Add attachments for {pid}")
    if not paths:
        return
    _finish_add_attachments(pid, paths, "file(s)")


def do_add_attachment_folder(pid):
    """Opens a native folder picker and copies the ENTIRE folder's
    contents into Attachments, preserving its internal subfolder
    structure (see add_patient_attachments) -- the button-driven
    equivalent of dragging a folder onto the drop zone, for when
    drag-and-drop isn't available or isn't the user's preference."""
    folder = filedialog.askdirectory(title=f"Add a folder of attachments for {pid}")
    if not folder:
        return
    _finish_add_attachments(pid, [folder], "folder(s)")


def do_remove_attachment(pid, relative_path, display_name):
    if not modern_askyesno("Remove Attachment", f"Remove \"{display_name}\" from {pid}? This cannot be undone."):
        return
    ok, err = remove_patient_attachment(pid, relative_path)
    if not ok:
        modern_showerror("Remove Failed", err)
        return
    if inspector_current_pid["value"] == pid:
        refresh_inspector(pid)


def do_open_attachment(abspath):
    ok, err = open_document(abspath)
    if not ok:
        modern_showerror("Could Not Open File", err)


def _on_attachment_drop(event, pid):
    """<<Drop>> handler for the Inspector's Attachments drop zone (only
    ever bound when TKINTERDND2_AVAILABLE). app.tk.splitlist is Tk's own
    Tcl-list parser -- it's what correctly handles the brace-quoting
    tkdnd uses for paths containing spaces, unlike naively stripping
    '{'/'}' characters out of event.data."""
    try:
        paths = [p for p in app.tk.splitlist(event.data) if os.path.isfile(p)]
    except Exception:
        paths = []
    if not paths:
        return
    ok_count, errors = add_patient_attachments(pid, paths)
    if inspector_current_pid["value"] == pid:
        refresh_inspector(pid)
    if errors:
        detail = "\n".join(f"{os.path.basename(p)}: {e}" for p, e in errors[:8])
        modern_showwarning(
            "Some Attachments Failed",
            f"{ok_count} of {len(paths)} file(s) added for {pid}.\n\n{detail}")
    elif ok_count:
        ui_event_queue.put(("toast", ("Attachments Added", f"{ok_count} file(s) added for {pid} (drag & drop).")))


def refresh_inspector(pid):
    """Rebuilds the inspector body for the given patient id. Safe to call
    repeatedly (e.g. from periodic_refresh) -- cheap relative to a full
    tree rebuild, and only runs while the panel is actually open."""
    for child in inspector_scroll.winfo_children():
        child.destroy()

    with data_lock:
        d = dict(patient_data.get(pid, {}))
    if not d:
        ctk.CTkLabel(inspector_scroll, text="This record is no longer in the worklist.",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED,
                     wraplength=260, justify="left").pack(pady=20, padx=10)
        return

    patient_body = _inspector_section(f"{pid}")
    _inspector_row(patient_body, "Name", d.get("patient_name"))
    _inspector_row(patient_body, "Institution", d.get("institution"))
    _inspector_row(patient_body, "Modality", d.get("modality"))
    _inspector_row(patient_body, "Source", d.get("source"))

    transfer_body = _inspector_section("Status & Transfer")
    _inspector_row(transfer_body, "Status", d.get("status"))
    _inspector_row(transfer_body, "Received", d.get("time"))
    _inspector_row(transfer_body, "Sent", d.get("sent_time"))
    _inspector_row(transfer_body, "Pushed To", d.get("push_target"))
    if d.get("last_error"):
        _inspector_row(transfer_body, "Last Error", d.get("last_error"))

    ckpt = get_checkpoint_info(pid)
    if ckpt:
        ckpt_body = _inspector_section("↻  Resume Checkpoint")
        _inspector_row(ckpt_body, "Destination", ckpt.get("destination"))
        _inspector_row(ckpt_body, "Progress", f"{len(ckpt.get('sent_sop_uids', []) or [])} sent")
        if ckpt.get("failure_reason"):
            _inspector_row(ckpt_body, "Interrupted", ckpt.get("failure_reason"))

    try:
        studies = get_studies_for_patient(pid)
    except Exception:
        studies = {}
    studies_body = _inspector_section(f"Studies ({len(studies)})")
    if not studies:
        ctk.CTkLabel(studies_body, text="No local files for this patient.",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w")
    else:
        for study_uid, info in list(studies.items())[:6]:
            short_uid = (study_uid[:24] + "…") if len(study_uid) > 24 else study_uid
            _inspector_row(studies_body, info.get("modality") or "—",
                            f"{short_uid}  ·  {info.get('count', 0)} img")

    try:
        attachments = list_patient_attachments(pid)
    except Exception:
        attachments = []
    attach_body = _inspector_section(f"Attachments ({len(attachments)})")
    if not attachments:
        ctk.CTkLabel(attach_body, text="No attachments yet.",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(anchor="w")
    else:
        for relative_path, abspath, size in attachments[:12]:
            # Display name drops the "Attachments/" prefix, keeping any
            # real subfolder structure visible (e.g. "Consent/signed.pdf").
            display_name = relative_path[len(ATTACHMENTS_SUBDIR) + 1:] if relative_path.startswith(ATTACHMENTS_SUBDIR + "/") else relative_path
            row = ctk.CTkFrame(attach_body, fg_color="transparent")
            row.pack(fill="x", pady=1)
            ctk.CTkLabel(row, text="", image=get_icon("file-text", size=13, color=THEME_TEXT_MUTED)).pack(side="left")
            name_lbl = ctk.CTkLabel(row, text=f"{display_name}  ·  {_human_size(size)}",
                                    font=get_font("small"), text_color=THEME_TEXT, anchor="w",
                                    wraplength=155, justify="left", cursor="hand2")
            name_lbl.pack(side="left", fill="x", expand=True, padx=(4, 0))
            name_lbl.bind("<Button-1>", lambda _e, p=abspath: do_open_attachment(p))
            ctk.CTkButton(row, text="", image=get_icon("x", size=11, color=THEME_TEXT_MUTED),
                         width=20, height=20, corner_radius=6, fg_color="transparent",
                         hover_color=THEME_NEUTRAL_BTN_HOVER,
                         command=lambda rp=relative_path, dn=display_name: do_remove_attachment(pid, rp, dn)
                         ).pack(side="right")
        if len(attachments) > 12:
            ctk.CTkLabel(attach_body, text=f"+ {len(attachments) - 12} more (open the patient folder to see all)",
                        font=get_font("micro"), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 0))
    add_row = ctk.CTkFrame(attach_body, fg_color="transparent")
    add_row.pack(fill="x", pady=(6, 0))
    ctk.CTkButton(add_row, text="Add File(s)…", height=28, corner_radius=8,
                 image=get_icon("plus", size=13, color=THEME_TEXT),
                 fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                 command=lambda: do_add_attachments(pid)).pack(side="left", fill="x", expand=True)
    ctk.CTkButton(add_row, text="Add Folder…", height=28, corner_radius=8,
                 image=get_icon("folder", size=13, color=THEME_TEXT),
                 fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                 command=lambda: do_add_attachment_folder(pid)).pack(side="left", fill="x", expand=True, padx=(4, 0))
    if TKINTERDND2_AVAILABLE:
        ctk.CTkLabel(attach_body, text="or drag files here", font=get_font("micro"),
                    text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 0))
        try:
            attach_body.drop_target_register(DND_FILES)
            attach_body.dnd_bind("<<Drop>>", lambda e, p=pid: _on_attachment_drop(e, p))
        except Exception:
            pass  # DnD is purely additive -- the Add button above still works either way

    actions_body = _inspector_section("Quick Actions")
    ctk.CTkButton(actions_body, text="Open / Create Report", height=30, corner_radius=8,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: open_report_action(pid)).pack(fill="x", pady=2)
    ctk.CTkButton(actions_body, text="Open / Create History", height=30, corner_radius=8,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: open_history_action(pid)).pack(fill="x", pady=2)
    ctk.CTkButton(actions_body, text="Open in Viewer…", height=30, corner_radius=8,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: open_patient_in_viewer_picker(pid)).pack(fill="x", pady=2)
    ctk.CTkButton(actions_body, text="Open Local Folder", height=30, corner_radius=8,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=lambda: open_in_viewer(pid)).pack(fill="x", pady=2)


def _on_inspector_selection(tree):
    sel = tree.selection()
    if not sel:
        close_inspector()


def _on_worklist_name_click(tree, event):
    """Opens the Inspector only when the click landed on the Patient Name
    cell -- previously any selection change (including box-select, ctrl/
    shift-click, or 'Select All') popped the Inspector open, which fought
    with normal multi-row selection. Clicking the name is now the one
    dedicated way to open it."""
    row_id = tree.identify_row(event.y)
    if not row_id:
        return
    col = _worklist_column_at(tree, event)
    if col != "patient_name":
        return
    tree.selection_set(row_id)
    tree.focus(row_id)
    inspector_current_pid["value"] = row_id
    open_inspector()
    refresh_inspector(row_id)


rec_tree.bind("<<TreeviewSelect>>", lambda _e: _on_inspector_selection(rec_tree), add="+")
push_tree.bind("<<TreeviewSelect>>", lambda _e: _on_inspector_selection(push_tree), add="+")
rec_tree.bind("<Button-1>", lambda e: _on_worklist_name_click(rec_tree, e), add="+")
push_tree.bind("<Button-1>", lambda e: _on_worklist_name_click(push_tree, e), add="+")

# =========================================================
# TABLE VIEW OPTIONS: column visibility + density
# =========================================================
# ttk.Treeview has no native "hide column" call, but `displaycolumns`
# lets you show a subset of `columns` in their original order without
# touching the underlying data model -- that's what drives visibility
# here. Density adjusts the shared Treeview rowheight style, so it
# applies to both tables together (one Treeview style backs both).

_column_visibility = {id(rec_tree): {c: True for c in WL_COLUMNS},
                      id(push_tree): {c: True for c in WL_COLUMNS}}
_density_state = {"mode": "Comfortable"}  # Comfortable=28px rows, Compact=22px

DEFAULT_ALWAYS_VISIBLE = {"sel", "patient_id", "patient_name", "status"}


def _load_view_options():
    """Restores saved column visibility + density -- the 'saved layout'
    part of the table-experience spec. Column *order*/drag-reorder isn't
    persisted because ttk.Treeview doesn't support reordering at all
    (see the View Options section further up); only what's genuinely
    save-able (show/hide + density) is."""
    try:
        if not os.path.isfile(VIEW_OPTIONS_FILE):
            return
        with open(VIEW_OPTIONS_FILE, "r", encoding="utf-8") as f:
            saved = json.load(f)
        _density_state["mode"] = saved.get("density", "Comfortable")
        for key, tree in (("rec", rec_tree), ("push", push_tree)):
            saved_cols = saved.get(f"{key}_columns")
            if isinstance(saved_cols, dict):
                vis = _column_visibility[id(tree)]
                for col in WL_COLUMNS:
                    if col in DEFAULT_ALWAYS_VISIBLE:
                        continue
                    if col in saved_cols:
                        vis[col] = bool(saved_cols[col])
    except Exception:
        log_exception("Failed to load view_options.json")


def _save_view_options():
    try:
        payload = {
            "density": _density_state["mode"],
            "rec_columns": _column_visibility[id(rec_tree)],
            "push_columns": _column_visibility[id(push_tree)],
        }
        atomic_write(VIEW_OPTIONS_FILE, json.dumps(payload, indent=2))
    except Exception:
        log_exception("Failed to save view_options.json")


_load_view_options()


def _apply_column_visibility(tree):
    vis = _column_visibility[id(tree)]
    tree.configure(displaycolumns=[c for c in WL_COLUMNS if vis.get(c, True)])
    # Hiding a column used to just leave its space empty on the right --
    # reclaim it for the columns still showing.
    _stretch_worklist_columns_to_fill(tree)


def _apply_density():
    style = ttk.Style()
    rowheight = 22 if _density_state["mode"] == "Compact" else 28
    style.configure("Treeview", rowheight=rowheight)
    # rec_tree/push_tree use their own style (see build_worklist_tree) --
    # keep it in sync with the density toggle too, with the same "+6"
    # bigger-worklist bump applied elsewhere.
    style.configure(WORKLIST_TREE_STYLE, rowheight=rowheight + 6)


_apply_column_visibility(rec_tree)
_apply_column_visibility(push_tree)
_apply_density()

# Growing the window (or the Receiver/Pusher pane within it) should also
# hand the newly-available width to the visible columns, not just leave
# it blank -- debounced so a drag-resize doesn't reflow on every pixel.
_stretch_resize_job = {"rec_tree": None, "push_tree": None}


def _debounced_stretch(tree, key):
    job = _stretch_resize_job.get(key)
    if job:
        try:
            app.after_cancel(job)
        except Exception:
            pass
    _stretch_resize_job[key] = app.after(120, lambda: _stretch_worklist_columns_to_fill(tree))


rec_tree.bind("<Configure>", lambda _e: _debounced_stretch(rec_tree, "rec_tree"), add="+")
push_tree.bind("<Configure>", lambda _e: _debounced_stretch(push_tree, "push_tree"), add="+")


def _close_popover(win):
    if win is not None and win.winfo_exists():
        win.destroy()


def _clamp_popover_geometry(x, y, width, height, margin=4):
    """Keeps a borderless CTkToplevel popover fully on-screen. Anchoring a
    popover's left edge to its trigger widget's left edge (the naive
    approach) pushes it off the right edge of the screen whenever the
    trigger sits near the right side of the window -- e.g. the Receiver
    tab's 'View' button. Every overrideredirect(True) popover should run
    its computed x/y through this before calling .geometry()."""
    try:
        screen_w = app.winfo_screenwidth()
        screen_h = app.winfo_screenheight()
    except Exception:
        return x, y
    x = max(margin, min(x, screen_w - width - margin))
    y = max(margin, min(y, screen_h - height - margin))
    return x, y


def open_view_options_popover(tree, anchor_btn):
    """A small floating panel (checkboxes + density radio) anchored under
    the 'View' button -- the practical equivalent of a column-picker
    menu, since CTk has no native dropdown-with-checkboxes widget."""
    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.overrideredirect(True)
    win.attributes("-topmost", True)
    win.configure(fg_color=THEME_SURFACE)
    app.update_idletasks()
    win_w = 220
    win_h = 92 + 26 * (len(WL_COLUMNS) - 1)  # a bit taller to fit the new title/close bar
    # Right-align under the button rather than left-align: this popover is
    # opened from a button that's often near the right edge of its toolbar
    # (e.g. Receiver tab), so left-aligning it pushed it off-screen.
    x = anchor_btn.winfo_rootx() + anchor_btn.winfo_width() - win_w
    y = anchor_btn.winfo_rooty() + anchor_btn.winfo_height() + 4
    x, y = _clamp_popover_geometry(x, y, win_w, win_h)
    win.geometry(f"{win_w}x{win_h}+{x}+{y}")

    shell = ctk.CTkFrame(win, fg_color=THEME_SURFACE, corner_radius=10,
                          border_width=1, border_color=THEME_HEADING_BG)
    shell.pack(fill="both", expand=True, padx=1, pady=1)

    title_row = ctk.CTkFrame(shell, fg_color="transparent")
    title_row.pack(fill="x", padx=(12, 6), pady=(8, 0))
    ctk.CTkLabel(title_row, text="View Options", font=get_font("small", "bold"),
                 text_color=THEME_TEXT).pack(side="left")
    ctk.CTkButton(title_row, text="", width=24, height=24, corner_radius=6,
                  image=get_icon("x", size=13, color=THEME_TEXT_MUTED),
                  fg_color="transparent", hover_color=THEME_NEUTRAL_BTN_HOVER,
                  text_color=THEME_TEXT_MUTED,
                  command=lambda: _close_popover(win)).pack(side="right")

    ctk.CTkLabel(shell, text="Density", font=get_font("small", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=12, pady=(6, 2))
    density_var = ctk.StringVar(value=_density_state["mode"])

    def _set_density(mode):
        _density_state["mode"] = mode
        _apply_density()
        _save_view_options()

    dens_row = ctk.CTkFrame(shell, fg_color="transparent")
    dens_row.pack(fill="x", padx=12, pady=(0, 8))
    for mode in ("Comfortable", "Compact"):
        ctk.CTkRadioButton(dens_row, text=mode, value=mode, variable=density_var,
                            font=get_font("small"), command=lambda m=mode: _set_density(m)).pack(side="left", padx=(0, 10))

    ctk.CTkLabel(shell, text="Columns", font=get_font("small", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=12, pady=(4, 2))
    vis = _column_visibility[id(tree)]
    for col in WL_COLUMNS:
        if col == "sel":
            continue
        var = ctk.BooleanVar(value=vis.get(col, True))

        def _toggle(c=col, v=var):
            if c in DEFAULT_ALWAYS_VISIBLE:
                v.set(True)  # core identity/status columns can't be hidden
                return
            vis[c] = v.get()
            _apply_column_visibility(tree)
            _save_view_options()

        cb = ctk.CTkCheckBox(shell, text=WL_HEADINGS[col], variable=var, font=get_font("small"),
                              command=_toggle)
        if col in DEFAULT_ALWAYS_VISIBLE:
            cb.configure(state="disabled")
        cb.pack(anchor="w", padx=14, pady=1)

    win.bind("<FocusOut>", lambda _e: app.after(120, lambda: _close_popover(win)))
    win.focus_set()


push_view_options_btn.configure(command=lambda: open_view_options_popover(push_tree, push_view_options_btn))

# =========================================================
# MODERN CONFIRM / ALERT DIALOGS
# =========================================================
# Themed, rounded, blocking dialogs matching the app's palette. Every
# messagebox.showinfo/showerror/showwarning/askyesno call site in the app
# (~48 of them) has been mechanically swept to modern_showinfo /
# modern_showerror / modern_showwarning / modern_askyesno below, which
# match the originals' (title, message, **kwargs) call signature exactly
# -- same blocking behavior, same return value for askyesno -- so no
# call site needed to change beyond its name.

def modern_confirm(title, message, danger=False):
    result = {"value": False}
    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title(title)
    win.attributes("-topmost", True)
    win.transient(app)
    _safe_grab_set(win)
    win.configure(fg_color=THEME_SURFACE)
    # Height grows with message length instead of a fixed 180px, so longer
    # confirmation text doesn't get clipped at the bottom of the dialog.
    est_lines = max(1, len(message) // 46 + message.count("\n") + 1)
    win_height = min(420, 150 + est_lines * 18)
    win.geometry(f"380x{win_height}")
    app.update_idletasks()
    x = app.winfo_rootx() + (app.winfo_width() - 380) // 2
    y = app.winfo_rooty() + (app.winfo_height() - win_height) // 2
    win.geometry(f"+{x}+{y}")
    win.resizable(False, False)

    ctk.CTkLabel(win, text=title, font=get_font("section", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=20, pady=(20, 4))
    make_wrapped_label(win, message, 340, font=get_font("body"),
                        text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=20, pady=(0, 16))

    btn_row = ctk.CTkFrame(win, fg_color="transparent")
    btn_row.pack(fill="x", padx=20, pady=(0, 18), side="bottom")

    def _confirm():
        result["value"] = True
        win.destroy()

    def _cancel():
        result["value"] = False
        win.destroy()

    ctk.CTkButton(btn_row, text="Cancel", width=110, height=32, corner_radius=8,
                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                  command=_cancel).pack(side="right", padx=(8, 0))
    ctk.CTkButton(btn_row, text="Delete" if danger else "Confirm", width=110, height=32, corner_radius=8,
                  fg_color=THEME_DANGER if danger else THEME_ACCENT,
                  hover_color=THEME_DANGER_HOVER if danger else THEME_ACCENT_HOVER,
                  command=_confirm).pack(side="right")

    win.bind("<Escape>", lambda _e: _cancel())
    win.bind("<Return>", lambda _e: _confirm())
    _fade_in_window(win)
    win.wait_window()
    return result["value"]


def modern_askyesno(title, message, **_ignored_kwargs):
    """Drop-in replacement for modern_askyesno(title, message, ...) --
    extra kwargs some call sites pass (parent=, icon=) are accepted and
    ignored, same as the rest of this dialog family below."""
    return modern_confirm(title, message, danger=False)


def modern_alert(title, message, kind="info", parent=None):
    """Themed single-button alert replacing showinfo/showerror/showwarning.
    Blocking (grab_set + wait_window) exactly like the tkinter originals,
    so call sites that run code immediately after showing one keep the
    same "user has dismissed it" ordering guarantee."""
    icon_name = {"error": "circle-x", "warning": "triangle-alert"}.get(kind)
    color = {"error": THEME_DANGER, "warning": THEME_WARNING, "info": THEME_ACCENT}.get(kind, THEME_ACCENT)

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title(title)
    win.attributes("-topmost", True)
    win.transient(app)
    _safe_grab_set(win)
    win.configure(fg_color=THEME_SURFACE)
    win.resizable(False, False)

    header = ctk.CTkFrame(win, fg_color="transparent")
    header.pack(fill="x", padx=20, pady=(20, 4))
    ctk.CTkLabel(header, text="", image=get_icon(icon_name, size=20, color=color)).pack(side="left", padx=(0, 8))
    ctk.CTkLabel(header, text=title, font=get_font("section", "bold"), text_color=color).pack(side="left")

    make_wrapped_label(win, message, 340, font=get_font("body"),
                        text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=20, pady=(0, 16))

    btn_row = ctk.CTkFrame(win, fg_color="transparent")
    btn_row.pack(fill="x", padx=20, pady=(0, 18), side="bottom")
    ctk.CTkButton(btn_row, text="OK", width=110, height=32, corner_radius=8,
                  fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
                  command=win.destroy).pack(side="right")

    win.bind("<Escape>", lambda _e: win.destroy())
    win.bind("<Return>", lambda _e: win.destroy())

    win.update_idletasks()
    w, h = max(380, win.winfo_reqwidth() + 20), win.winfo_reqheight() + 10
    x = app.winfo_rootx() + (app.winfo_width() - w) // 2
    y = app.winfo_rooty() + (app.winfo_height() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    _fade_in_window(win)
    win.wait_window()


def modern_showinfo(title, message, **kw):
    modern_alert(title, message, kind="info", **kw)


def modern_showerror(title, message, **kw):
    modern_alert(title, message, kind="error", **kw)


def modern_showwarning(title, message, **kw):
    modern_alert(title, message, kind="warning", **kw)

# =========================================================
# DESTINATIONS TAB CALLBACKS
# =========================================================

def _live_validate_dest_field(key):
    """9.3 -- runs the real validators as the user types, instead of only
    on Save. Blank fields don't get flagged live (that's still enforced
    at Save time via _mark_dest_field_invalid in do_add_update_dest) so
    the field doesn't turn red before the user has even started typing."""
    entry = dest_fields.get(key)
    if entry is None:
        return
    value = entry.get().strip()
    if not value:
        _clear_dest_field_invalid(key)
        return
    if key == "ae":
        ok, _msg = validate_ae_title(value, "Remote AE Title")
    elif key == "calling_ae":
        ok, _msg = validate_ae_title(value, "Calling AE Title")
    elif key == "port":
        ok, _msg = validate_port(value, "Port")
    elif key == "doc_transfer_port":
        ok, _msg = validate_port(value, "Document Transfer Port")
    else:
        ok = True
    (_clear_dest_field_invalid if ok else _mark_dest_field_invalid)(key)


def _mark_dest_field_invalid(key):
    entry = dest_fields.get(key)
    if entry is not None:
        entry.configure(border_color=THEME_DANGER, border_width=2)


def _clear_dest_field_invalid(key):
    entry = dest_fields.get(key)
    if entry is not None:
        entry.configure(border_color=dest_field_default_border.get(key, THEME_HEADING_BG), border_width=1)


def _clear_all_dest_field_invalid():
    for key in dest_fields:
        _clear_dest_field_invalid(key)


def do_add_update_dest():
    _clear_all_dest_field_invalid()
    name = dest_fields["name"].get().strip()
    ae = dest_fields["ae"].get().strip()
    calling_ae = dest_fields["calling_ae"].get().strip()  # optional
    ip = dest_fields["ip"].get().strip()
    port = dest_fields["port"].get().strip()
    if not all([name, ae, ip, port]):
        for key, val in (("name", name), ("ae", ae), ("ip", ip), ("port", port)):
            if not val:
                _mark_dest_field_invalid(key)
        modern_showerror("Error", "Profile name, Remote AE, IP, and Port are required.")
        return

    ok, ae_clean_or_err = validate_ae_title(ae, "Remote AE Title")
    if not ok:
        _mark_dest_field_invalid("ae")
        modern_showerror("Invalid Remote AE Title", ae_clean_or_err)
        return

    if calling_ae:
        ok, calling_clean_or_err = validate_ae_title(calling_ae, "Calling AE Title")
        if not ok:
            _mark_dest_field_invalid("calling_ae")
            modern_showerror("Invalid Calling AE Title", calling_clean_or_err)
            return

    ok, port_val_or_err = validate_port(port, "Port")
    if not ok:
        _mark_dest_field_invalid("port")
        modern_showerror("Invalid Port", port_val_or_err)
        return

    dests = load_destinations()
    existing = next((d for d in dests if d["name"] == name), None)

    # Document transfer fields are independently optional -- a destination
    # with doc_transfer_enabled off is fully unaffected by whatever else is
    # (or isn't) filled into doc_transfer_port/ip/auth_key.
    doc_transfer_port = dest_fields["doc_transfer_port"].get().strip() if "doc_transfer_port" in dest_fields else ""
    doc_transfer_ip = dest_fields["doc_transfer_ip"].get().strip() if "doc_transfer_ip" in dest_fields else ""
    doc_transfer_auth_key = dest_fields["doc_transfer_auth_key"].get().strip() if "doc_transfer_auth_key" in dest_fields else ""
    doc_transfer_enabled = dest_doc_transfer_enabled_var.get() if "dest_doc_transfer_enabled_var" in globals() else False
    doc_transfer_use_dicom_host = dest_doc_transfer_use_dicom_host_var.get() if "dest_doc_transfer_use_dicom_host_var" in globals() else True

    if doc_transfer_enabled:
        ok, port_or_err = validate_port(doc_transfer_port, "Document Transfer Port")
        if not ok:
            modern_showerror("Invalid Document Transfer Port", port_or_err)
            return
        if not doc_transfer_use_dicom_host and not doc_transfer_ip:
            modern_showerror("Error", "Document Transfer IP is required when not using the DICOM host.")
            return

    new_entry = {
        "name": name,
        "ae": ae,
        "calling_ae": calling_ae,  # empty string = use default "RAPPS_PUSH"
        "ip": ip,
        "port": port,
        "default": dest_default_var.get(),
        "doc_transfer_enabled": doc_transfer_enabled,
        "doc_transfer_port": doc_transfer_port,
        "doc_transfer_use_dicom_host": doc_transfer_use_dicom_host,
        "doc_transfer_ip": doc_transfer_ip,
        "doc_transfer_auth_key": doc_transfer_auth_key,
    }

    if dest_default_var.get():
        for d in dests:
            d["default"] = False

    if existing:
        idx = dests.index(existing)
        dests[idx] = new_entry
    else:
        dests.append(new_entry)

    save_destinations(dests)
    refresh_destinations_ui()
    dest_status_lbl.configure(text=f"{'Updated' if existing else 'Added'}: {name}", text_color=THEME_SUCCESS)


def do_del_dest():
    name = dest_select_var.get()
    if not name or name == "(none)":
        return
    dests = [d for d in load_destinations() if d["name"] != name]
    save_destinations(dests)
    refresh_destinations_ui()
    dest_status_lbl.configure(text=f"Deleted: {name}", text_color=THEME_DANGER)


def update_dest_trust_badge():
    if "dest_trust_auth_badge" not in globals() or "doc_transfer_auth_key" not in dest_fields:
        return
    has_key = bool(dest_fields["doc_transfer_auth_key"].get().strip())
    icon = "lock" if has_key else "lock-open"
    color = THEME_SUCCESS if has_key else THEME_TEXT_MUTED
    text = "  Authenticated" if has_key else "  Unauthenticated"
    dest_trust_auth_badge.configure(text=text, image=get_icon(icon, size=13, color=color),
                                    compound="left", text_color=color)


def do_load_dest_into_form():
    name = dest_select_var.get()
    d = get_destination_by_name(name)
    if not d:
        return
    for key, entry in dest_fields.items():
        entry.delete(0, "end")
        entry.insert(0, d.get(key, ""))
    dest_default_var.set(d.get("default", False))
    # calling_ae may not exist in migrated legacy configs — leave blank (= default)
    if "calling_ae" in dest_fields:
        dest_fields["calling_ae"].delete(0, "end")
        dest_fields["calling_ae"].insert(0, d.get("calling_ae", ""))

    # Document transfer fields may not exist on destinations saved before
    # this feature -- .get(key, default) means those load as fully disabled.
    for key in ("doc_transfer_port", "doc_transfer_ip", "doc_transfer_auth_key"):
        if key in dest_fields:
            dest_fields[key].delete(0, "end")
            dest_fields[key].insert(0, d.get(key, ""))
    if "dest_doc_transfer_enabled_var" in globals():
        dest_doc_transfer_enabled_var.set(d.get("doc_transfer_enabled", False))
    if "dest_doc_transfer_use_dicom_host_var" in globals():
        dest_doc_transfer_use_dicom_host_var.set(d.get("doc_transfer_use_dicom_host", True))
    update_dest_trust_badge()


def do_echo_dest():
    name = dest_fields["name"].get().strip() or dest_select_var.get()
    ae = dest_fields["ae"].get().strip()
    ip = dest_fields["ip"].get().strip()
    port = dest_fields["port"].get().strip()
    if not ae or not ip or not port.isdigit():
        modern_showerror("Error", "Fill in the destination fields first.")
        return
    calling_ae_field = dest_fields.get("calling_ae")
    calling_ae_str = calling_ae_field.get().strip() if calling_ae_field else None
    dest_echo_btn.configure(state="disabled", text="Testing...")

    def run():
        ok, msg = dicom_echo(ae, ip, port, calling_ae=calling_ae_str)
        def on_ui():
            dest_echo_btn.configure(state="normal", text="C-ECHO Test")
            dest_status_lbl.configure(text=msg, text_color=THEME_SUCCESS if ok else THEME_DANGER)
        app.after(0, on_ui)
    threading.Thread(target=run, daemon=True).start()

# NOTE: dest_add_btn/dest_del_btn/dest_echo_btn/dest_select_var wiring now
# lives inside build_admin_only_tabs() so it re-runs correctly every time
# these widgets are rebuilt on an Admin<->User mode switch.

# =========================================================
# ROUTING RULES CALLBACKS
# =========================================================

def do_add_routing_rule():
    rule = {
        "modality": routing_fields["modality"].get().strip(),
        "institution": routing_fields["institution"].get().strip(),
        "source_ae": routing_fields["source_ae"].get().strip(),
        "destination": routing_dest_var.get(),
    }
    if rule["destination"] in ("(none)", ""):
        modern_showerror("Error", "Select a destination for this rule.")
        return
    rules = load_routing_rules()
    rules.append(rule)
    save_routing_rules(rules)
    refresh_routing_ui()
    for e in routing_fields.values():
        e.delete(0, "end")


def do_test_routing_rule():
    """5.2 -- runs the form's current (not-yet-saved) field values through
    the exact matching logic resolve_destination_for_patient() uses per
    rule, against every patient currently in the worklist, and reports
    how many would match -- without saving anything."""
    draft_rule = {
        "modality": routing_fields["modality"].get().strip(),
        "institution": routing_fields["institution"].get().strip(),
        "source_ae": routing_fields["source_ae"].get().strip(),
    }
    with data_lock:
        rows = list(patient_data.values())
    matched = sum(
        1 for d in rows
        if _rule_matches_patient_fields(draft_rule, d.get("modality", ""), d.get("institution", ""), d.get("source", ""))
    )
    modern_showinfo("Test Rule", f"This rule would currently match {matched} of {len(rows)} studies in the worklist.")


def do_del_routing_rule():
    sel = routing_tree.selection()
    if not sel:
        return
    idx = routing_tree.index(sel[0])
    rules = load_routing_rules()
    if 0 <= idx < len(rules):
        del rules[idx]
        save_routing_rules(rules)
        refresh_routing_ui()


def do_move_rule(direction):
    sel = routing_tree.selection()
    if not sel:
        return
    idx = routing_tree.index(sel[0])
    rules = load_routing_rules()
    new_idx = idx + direction
    if 0 <= new_idx < len(rules):
        rules[idx], rules[new_idx] = rules[new_idx], rules[idx]
        save_routing_rules(rules)
        refresh_routing_ui()

# NOTE: routing_*_btn wiring now lives inside build_admin_only_tabs().

# =========================================================
# SOP CLASSES TAB CALLBACKS
# =========================================================

def do_save_sop():
    if sop_search_var.get().strip():
        # A filter is active -- the boxes only show matching lines right
        # now, so reading them directly would silently save a truncated
        # config. Clearing the filter restores the full text synchronously
        # via the trace on sop_search_var before we read it below.
        sop_search_var.set("")
    raw_classes = sop_classes_box.get("1.0", "end").strip()
    raw_ts = sop_ts_box.get("1.0", "end").strip()
    sop_dict, ts_dict = {}, {}
    for line in raw_classes.splitlines():
        if "=" in line:
            k, _, v = line.partition("=")
            sop_dict[k.strip()] = v.strip()
    for line in raw_ts.splitlines():
        if "=" in line:
            k, _, v = line.partition("=")
            ts_dict[k.strip()] = v.strip()
    try:
        save_sop_ini(sop_dict, ts_dict)
        sop_status_lbl.configure(text="Saved. Restart receiver to apply.", text_color=THEME_SUCCESS)
    except Exception as e:
        sop_status_lbl.configure(text=f"Save failed: {e}", text_color=THEME_DANGER)


# NOTE: Clear-Log functionality has been permanently removed -- logs are
# immutable audit records. See rotate_logs_if_needed() / archive_log_file()
# / do_archive_logs_now() / do_export_logs() for the retention-based
# replacement. sop_save_btn/sop_load_btn/log_*_btn/log_file_var wiring
# lives inside build_admin_only_tabs().

# =========================================================
# QUERY/RETRIEVE CALLBACKS
# =========================================================

def do_qr_find():
    rae = qr_ae_entry.get().strip()
    rip = qr_ip_entry.get().strip()
    rport = qr_port_entry.get().strip()
    if not rae or not rip or not rport.isdigit():
        modern_showerror("Error", "Remote AE, IP, and Port are required.")
        return

    qr_find_btn.configure(state="disabled", text="Querying...")
    qr_status_label.configure(text="Sending C-FIND...")

    filter_pid = qr_filter_entries["pid"].get().strip()
    filter_pname = qr_filter_entries["pname"].get().strip()
    filter_date = qr_filter_entries["date"].get().strip()
    filter_mod = qr_filter_entries["mod"].get().strip()

    def run():
        ok, results, msg = query_remote_studies(
            rae, rip, rport,
            patient_id=filter_pid,
            patient_name=filter_pname,
            study_date=filter_date,
            modality=filter_mod,
        )

        def on_ui():
            qr_find_btn.configure(state="normal", text="C-FIND Query")
            qr_status_label.configure(text=msg)
            qr_tree.delete(*qr_tree.get_children())
            for r in results:
                qr_tree.insert("", "end", values=(
                    r["patient_id"], r["patient_name"], r["study_date"],
                    r["modality"], r["instances"], r["description"], r["study_uid"],
                ))

        app.after(0, on_ui)

    threading.Thread(target=run, daemon=True).start()


def do_qr_retrieve():
    config = decrypt_and_load(RECEIVER_CONFIG)
    if not config:
        modern_showerror("Error", "Receiver config missing. The receiver must be configured and running to accept retrieved studies.")
        return
    our_ae, _port = config.split("|")

    if not receiver_state.get("running"):
        if not modern_askyesno("Warning",
                "The receiver is not currently running.\n"
                "Retrieved studies will fail unless the remote PACS can reach this node.\n\n"
                "Continue anyway?"):
            return

    rae = qr_ae_entry.get().strip()
    rip = qr_ip_entry.get().strip()
    rport = qr_port_entry.get().strip()

    selected = qr_tree.selection()
    if not selected:
        modern_showwarning("No Selection", "Select at least one study to retrieve.")
        return

    study_uids = [qr_tree.set(iid, "study_uid") for iid in selected]
    qr_retrieve_btn.configure(state="disabled", text="Retrieving...")
    qr_status_label.configure(text=f"Retrieving {len(study_uids)} study/studies...")

    def run():
        errors = []
        for uid in study_uids:
            ok, msg = retrieve_study(rae, rip, rport, uid, our_ae)
            if not ok:
                errors.append(f"{uid[:20]}...: {msg}")

        def on_ui():
            qr_retrieve_btn.configure(state="normal", text="C-MOVE Retrieve Selected")
            if errors:
                qr_status_label.configure(text=f"Errors: {len(errors)}")
                modern_showerror("Retrieve Errors", "\n".join(errors))
            else:
                qr_status_label.configure(text="Retrieve complete — check worklist.")

        app.after(0, on_ui)

    threading.Thread(target=run, daemon=True).start()

# NOTE: qr_find_btn/qr_retrieve_btn wiring now lives inside build_admin_only_tabs().

# =========================================================
# EVENT PUMP (main-thread UI update loop)
# =========================================================

_low_disk_warned_at = [0.0]   # track last warn time to avoid spam

# Drives the "Live Receive Progress" card (rec_live_progress / labels).
# There's no upfront file-count for an incoming association, so a
# receive "session" is just tracked by activity: it opens on the first
# C-STORE and auto-closes after RECV_IDLE_TIMEOUT seconds of silence.
RECV_IDLE_TIMEOUT = 2.5
_recv_session = {"active": False, "count": 0, "bytes": 0, "start": None, "last_update": None}


def _dispatch_ui_event(event, payload):
    """Handles a single dequeued UI event. Split out from pump_events()
    so a handler that raises (e.g. it still targets a widget from an
    Admin-only tab that the idle-timeout auto-lock tore down mid-session)
    can't abort the whole batch and silently drop every OTHER
    already-dequeued event for this tick -- see pump_events()."""
    if event == "refresh":
        populate_tree(rec_tree, rec_search_var.get())
        populate_tree(push_tree, push_search_var.get())

    elif event == "receiver_status":
        running = payload
        if running:
            rec_status_label.configure(text="● Running", text_color=THEME_SUCCESS)
        else:
            rec_status_label.configure(text="● Stopped", text_color=THEME_DANGER)
        refresh_receiver_monitoring()

    elif event == "doc_transfer_status":
        doc_running, doc_port = payload
        if "rec_doc_transfer_status_label" in globals():
            if doc_running:
                # §3.6 fix: an empty doc_transfer_receiver_auth_key
                # means every connection is accepted unauthenticated
                # (intentional, for backward compatibility -- see
                # _handle_doc_transfer_connection) but that state
                # previously had no visible indicator anywhere.
                # Surface it right on the status label the operator
                # is already looking at, in the warning color.
                auth_key = str(load_app_settings().get("doc_transfer_receiver_auth_key", "") or "")
                if auth_key:
                    rec_doc_transfer_status_label.configure(
                        text=f"● Doc transfer: listening on {doc_port}", text_color=THEME_SUCCESS)
                else:
                    rec_doc_transfer_status_label.configure(
                        text=f"● Doc transfer: listening on {doc_port} — ⚠ UNAUTHENTICATED (no auth key set; "
                             f"anyone on the network can send/receive documents)",
                        text_color=THEME_WARNING)
            else:
                rec_doc_transfer_status_label.configure(
                    text="○ Doc transfer: off", text_color=THEME_TEXT_MUTED)

    elif event == "push_started":
        overall_progress.set(0)
        push_progress_status_label.configure(text="Preparing to push selected studies…")
        overall_progress_label.configure(text="Starting push...")
        throughput_label.configure(text="")
        refresh_pusher_monitoring()

    elif event == "push_progress":
        sent = push_job["sent_images"]
        attempted = push_job["attempted_images"]
        total = push_job["total_images"]
        frac = attempted / max(total, 1)
        overall_progress.set(frac)
        cur_pid = push_job.get("current_pid") or ""
        cur_dest = push_job.get("current_dest") or ""
        if cur_pid and cur_dest:
            push_progress_status_label.configure(
                text=f"Sending patient {cur_pid} to {cur_dest}…")
        overall_progress_label.configure(text=f"{sent} / {total} images sent successfully")
        rate, eta = get_push_throughput_eta()
        if rate > 0:
            eta_str = f"ETA: {int(eta)}s" if eta is not None else "ETA: —"
            throughput_label.configure(text=f"{rate:.1f} img/s  {eta_str}")
        populate_tree(push_tree, push_search_var.get())
        refresh_pusher_monitoring()

    elif event == "push_finished":
        sent = push_job["sent_images"]
        total = push_job["total_images"]
        overall_progress.set(1.0)
        push_progress_status_label.configure(text="Push finished.")
        overall_progress_label.configure(text=f"Done — {sent} / {total} images sent successfully")
        throughput_label.configure(text="")
        populate_tree(rec_tree, rec_search_var.get())
        populate_tree(push_tree, push_search_var.get())
        refresh_pusher_monitoring()

    elif event == "push_error":
        modern_showerror("Push Error", payload)

    elif event == "import_progress":
        done, total = payload
        frac = done / max(total, 1)
        import_progress.set(frac)
        import_progress_label.configure(text=f"Importing: {done} / {total}")
        if done >= total:
            import_progress_label.configure(text=f"Done: {total} files processed")
            import_btn.configure(state="normal")

    elif event == "import_done":
        imported, failed, total = payload
        import_progress.set(1.0)
        msg = f"Import complete: {imported} imported, {failed} failed"
        import_progress_label.configure(text=msg)
        import_btn.configure(state="normal")
        populate_tree(rec_tree, rec_search_var.get())
        populate_tree(push_tree, push_search_var.get())

    elif event == "recv_progress":
        # A file just landed via C-STORE. DICOM storage doesn't
        # announce an upfront total for an association, so we
        # can't show a determinate percentage -- instead this
        # keeps an accurate running count/size/rate visible for
        # as long as files keep arriving, and idles itself out
        # (see the watchdog below) once nothing has arrived for
        # a couple of seconds.
        size_bytes, calling_ae, pid, modality = payload
        now = time.time()
        if not _recv_session["active"]:
            _recv_session["active"] = True
            _recv_session["count"] = 0
            _recv_session["bytes"] = 0
            _recv_session["start"] = now
            rec_live_progress.configure(mode="indeterminate")
            rec_live_progress.start()
        _recv_session["count"] += 1
        _recv_session["bytes"] += max(size_bytes or 0, 0)
        _recv_session["last_update"] = now
        elapsed = max(now - _recv_session["start"], 0.001)
        rate = _recv_session["count"] / elapsed
        mb = _recv_session["bytes"] / (1024 * 1024)
        mod_txt = f"{modality} " if modality else ""
        rec_live_progress_status_label.configure(
            text=f"Receiving {mod_txt}image for patient {pid} from {calling_ae}…")
        rec_live_progress_label.configure(
            text=f"Receiving… {_recv_session['count']} file(s), {mb:.1f} MB")
        rec_live_throughput_label.configure(text=f"{rate:.1f} files/s")

    elif event == "toast":
        title, message = payload
        _show_toast(title, message)
        if APP_SETTINGS.get("notification_sounds_enabled", True):
            try:
                app.bell()
            except Exception:
                pass

    elif event == "notification_added":
        _refresh_notification_center_ui()

    elif event == "low_disk_warning":
        now = time.time()
        if now - _low_disk_warned_at[0] > 300:  # throttle to once per 5 min
            _low_disk_warned_at[0] = now
            free_gb = payload
            modern_showwarning(
                "Low Disk Space",
                f"Free disk space is critically low: {free_gb:.2f} GB remaining.\n"
                f"The receiver may fail to save new images."
            )


def pump_events():
    """Drain the UI event queue and apply all pending updates on the
    main thread. This makes every background thread's state visible to
    the GUI without requiring explicit app.after() calls in each worker.
    Each event is dispatched -- and exception-isolated -- individually by
    _dispatch_ui_event() so one broken event can never swallow the rest
    of the same tick's batch."""
    while True:
        try:
            event, payload = ui_event_queue.get_nowait()
        except queue.Empty:
            break
        except Exception:
            log_exception("Error draining the UI event queue")
            break
        try:
            _dispatch_ui_event(event, payload)
        except Exception:
            log_exception(f"Error handling UI event {event!r}")


    # Idle watchdog for the Live Receive Progress card: once no C-STORE
    # has landed for RECV_IDLE_TIMEOUT seconds, stop the indeterminate
    # bar, show a brief "Done" summary, then relax back to the idle state.
    if _recv_session["active"] and _recv_session["last_update"] is not None:
        if time.time() - _recv_session["last_update"] > RECV_IDLE_TIMEOUT:
            _recv_session["active"] = False
            count = _recv_session["count"]
            mb = _recv_session["bytes"] / (1024 * 1024)
            elapsed = max(time.time() - _recv_session["start"], 0.001)
            try:
                rec_live_progress.stop()
                rec_live_progress.configure(mode="determinate")
                rec_live_progress.set(1.0)
                rec_live_progress_status_label.configure(text="Receive complete.")
                rec_live_progress_label.configure(
                    text=f"Done — {count} file(s) received, {mb:.1f} MB in {elapsed:.1f}s")
                rec_live_throughput_label.configure(text="")

                def _relax_recv_bar():
                    if not _recv_session["active"]:
                        rec_live_progress.set(0)
                        rec_live_progress_status_label.configure(text="")
                        rec_live_progress_label.configure(text="Idle — waiting for incoming studies")

                app.after(4000, _relax_recv_bar)
            except Exception:
                pass

    # Auto-tail active log file while Logs tab is selected (Admin-only tab)
    if admin_tabs_active["value"] and tabview.get().startswith("Logs") and log_tail_var.get():
        if log_display_mode_var.get() == "Structured":
            refresh_structured_log_view()
        else:
            refresh_log_view()

    app.after(get_refresh_interval_ms(), pump_events)


# Periodic background tree refresh (fallback for events that might be missed)
_refresh_counter = [0]


_KNOWN_TREE_VAR_NAMES = [
    "rec_tree", "push_tree", "routing_tree", "qr_tree",
    "dash_failures_tree", "dash_dest_health_tree", "health_tree", "offlineq_tree",
    "export_files_tree", "backup_history_tree", "ldap_roster_tree",
    "dt_sent_tree", "dt_recv_tree", "log_struct_tree",
]


def _all_known_trees():
    """Every ttk.Treeview the app has built so far. Some (the admin-only
    ones) only exist once build_admin_only_tabs() has run at least once,
    so this checks the module namespace rather than assuming they exist."""
    trees = []
    for name in _KNOWN_TREE_VAR_NAMES:
        obj = globals().get(name)
        if obj is not None:
            trees.append(obj)
    return trees


def refresh_ui_theme():
    """Re-applies the current palette (DARK_THEME_PALETTE, adjusted for
    High Contrast Mode) to every already-built widget, live -- no restart
    needed. This is not a Dark/Light switch (the app has only one theme);
    it exists so the High Contrast Mode and Larger Click Targets
    accessibility settings take effect immediately instead of requiring
    the user to relaunch the app.

    What updates live:
      * Every ttk.Treeview in the app (worklists, PACS Health, Offline
        Queue, Routing Rules, Query/Retrieve, Admin Dashboard tables) --
        these read from a shared ttk.Style, which DOES apply retroactively
        to already-built widgets, so one style update re-themes all of
        them at once.
      * The window background, header bar, admin bar, and tab bar --
        these were built with explicit hex colors (not CTk's theme-aware
        color tuples), so they're recolored directly here.
      * Every registered nav button's height (Larger Click Targets)."""
    global THEME_BG, THEME_SURFACE, THEME_HEADING_BG, THEME_ACCENT, THEME_ACCENT_HOVER
    global THEME_TEXT, THEME_TEXT_MUTED, THEME_NEUTRAL_BTN, THEME_NEUTRAL_BTN_HOVER
    global THEME_DANGER, THEME_DANGER_HOVER, THEME_SUCCESS, THEME_SUCCESS_HOVER
    global THEME_WARNING, THEME_WARNING_HOVER
    global STALE_HIGHLIGHT_COLOR, SEARCH_MATCH_HIGHLIGHT_BG, THEME_SEGMENTED_HOVER, THEME_ODD_ROW

    palette = _compute_theme_palette()

    THEME_BG = palette["bg"]
    THEME_SURFACE = palette["surface"]
    THEME_HEADING_BG = palette["heading_bg"]
    THEME_ACCENT = palette["accent"]
    THEME_ACCENT_HOVER = palette["accent_hover"]
    THEME_TEXT = palette["text"]
    THEME_TEXT_MUTED = palette["text_muted"]
    THEME_NEUTRAL_BTN = palette["neutral_btn"]
    THEME_NEUTRAL_BTN_HOVER = palette["neutral_btn_hover"]
    THEME_DANGER = palette["danger"]
    THEME_DANGER_HOVER = palette["danger_hover"]
    THEME_SUCCESS = palette["success"]
    THEME_SUCCESS_HOVER = palette["success_hover"]
    THEME_WARNING = palette["warning"]
    THEME_WARNING_HOVER = palette["warning_hover"]
    STALE_HIGHLIGHT_COLOR = palette["stale"]
    SEARCH_MATCH_HIGHLIGHT_BG = palette["search_highlight"]
    THEME_SEGMENTED_HOVER = palette["segmented_hover"]
    THEME_ODD_ROW = palette["odd_row"]

    try:
        app.configure(fg_color=THEME_BG)
        header_bar.configure(fg_color=THEME_SURFACE)
        header_title_lbl.configure(text_color=THEME_TEXT)
        header_subtitle_lbl.configure(text_color=THEME_TEXT_MUTED)
        header_about_btn.configure(text_color=THEME_TEXT_MUTED, hover_color=THEME_HEADING_BG)
        header_notif_btn.configure(hover_color=THEME_HEADING_BG, text_color=THEME_TEXT,
                                    image=get_icon("bell", size=16, color=THEME_TEXT))
        nav_collapse_btn.configure(image=get_icon(
            "panel-left-open" if nav_state["collapsed"] else "panel-left-close",
            size=15, color=THEME_TEXT_MUTED))
        _sync_nav_highlight()
        admin_bar.configure(fg_color=THEME_SURFACE)
        admin_bar_status_badge.configure(fg_color=THEME_HEADING_BG)
        admin_bar_status_lbl.configure(text_color=THEME_TEXT)
        tabview.configure(
            fg_color=THEME_SURFACE, segmented_button_fg_color=THEME_HEADING_BG,
            segmented_button_selected_color=THEME_ACCENT,
            segmented_button_selected_hover_color=THEME_ACCENT_HOVER,
            segmented_button_unselected_hover_color=THEME_SEGMENTED_HOVER,
            text_color=THEME_TEXT,
        )
        nav_frame.configure(fg_color=THEME_SURFACE)
        status_bar.configure(fg_color=THEME_SURFACE)
        for _sb_lbl in (sb_receiver_lbl, sb_queue_lbl, sb_user_lbl, sb_cpu_lbl, sb_ram_lbl,
                        sb_net_lbl, sb_version_lbl, sb_clock_lbl):
            if _sb_lbl is not sb_receiver_lbl:
                _sb_lbl.configure(text_color=THEME_TEXT_MUTED)
            _sb_lbl.set_bg(THEME_SURFACE)
        nav_collapse_btn.configure(fg_color=THEME_HEADING_BG, text_color=THEME_TEXT_MUTED)
        for _header in nav_section_headers.values():
            _header.configure(text_color=THEME_TEXT_MUTED)
        for _empty_lbl in _all_empty_state_labels:
            if _empty_lbl.winfo_exists():
                _empty_lbl.configure(text_color=THEME_TEXT_MUTED)
        _sync_nav_highlight()
        for _canvas in dash_graph_canvases.values():
            _canvas.configure(bg=THEME_SURFACE)
        _nav_btn_height = get_nav_button_height()
        for _btn in nav_buttons.values():
            _btn.configure(height=_nav_btn_height)
    except Exception:
        log_exception("Failed to refresh theme on top-level chrome")

    try:
        style = ttk.Style()
        style.configure("Treeview", background=THEME_SURFACE, foreground=THEME_TEXT,
                        fieldbackground=THEME_SURFACE, rowheight=get_worklist_row_height())
        style.configure("Treeview.Heading", background=THEME_HEADING_BG, foreground=THEME_TEXT)
        style.map("Treeview.Heading", background=[("active", palette["segmented_hover"])])
        style.map("Treeview", background=[("selected", THEME_ACCENT)], foreground=[("selected", "#ffffff")])
        # Receiver/Pusher worklists now follow the normal theme background
        # like every other tree (just with text forced to white -- see
        # build_worklist_tree) -- keep both row height and background/
        # heading colors in sync with theme changes (e.g. High Contrast
        # Mode) here too.
        style.configure(WORKLIST_TREE_STYLE, rowheight=get_worklist_row_height() + 6,
                        background=THEME_SURFACE, fieldbackground=THEME_SURFACE)
        style.configure(f"{WORKLIST_TREE_STYLE}.Heading", background=THEME_HEADING_BG)
        style.map(f"{WORKLIST_TREE_STYLE}.Heading", background=[("active", palette["segmented_hover"])])
        style.configure("Vertical.TScrollbar", background=THEME_SURFACE, troughcolor=THEME_BG,
                        bordercolor=THEME_SURFACE, arrowcolor=THEME_TEXT_MUTED)
        style.configure("Horizontal.TScrollbar", background=THEME_SURFACE, troughcolor=THEME_BG,
                        bordercolor=THEME_SURFACE, arrowcolor=THEME_TEXT_MUTED)

        # rec_tree/push_tree now use the same dark background + palette-
        # driven row tags as every other tree -- text stays pinned to
        # white via the WORKLIST_TREE_STYLE foreground set above, so
        # there's no longer any reason to exclude them from this loop.
        for tree in _all_known_trees():
            tree.tag_configure(STALE_HIGHLIGHT_TAG, foreground=STALE_HIGHLIGHT_COLOR)
            tree.tag_configure(SEARCH_MATCH_HIGHLIGHT_TAG, background=SEARCH_MATCH_HIGHLIGHT_BG)
            tree.tag_configure("even_row", background=THEME_SURFACE)
            tree.tag_configure("odd_row", background=palette["odd_row"])
    except Exception:
        log_exception("Failed to refresh theme on treeviews")


# =========================================================
# UNIVERSAL SEARCH (Phase 12)
# =========================================================
# Reuses existing data end-to-end: patient_data/_row_search_haystack
# (same fields the worklist search already matches), get_structured_log_records
# (Phase 6b), load_destinations/load_routing_rules (Destinations/Routing
# tabs), and a small curated index for Settings screens that have no
# single "record" to search (there's no settings database to query --
# these are just labeled jump-points into the tabs that already exist).

SETTINGS_SEARCH_INDEX = [
    ("bandwidth throttle limit rate", "Bandwidth Throttling", "Bandwidth"),
    ("ldap active directory login sso authentication", "LDAP / Active Directory", "LDAP / AD"),
    ("backup schedule restore history", "Backup & Restore", "Backup"),
    ("sop class transfer syntax", "SOP Classes", "SOP Classes"),
    ("routing rule autoroute modality institution", "Routing Rules", "Routing Rules"),
    ("destination pacs ae title port", "Destinations", "Destinations"),
    ("export zip password protected", "Export", "Export"),
    ("report pdf daily weekly monthly", "Reports", "Reports"),
    ("performance cpu ram disk throughput", "Performance", "Performance"),
    ("health echo c-echo online offline", "PACS Health", "PACS Health"),
    ("admin dashboard overview", "Admin Dashboard", "Admin Dashboard"),
    ("query retrieve c-find c-move", "Query/Retrieve", "Query/Retrieve"),
    ("logs audit trail search filter", "Logs", "Logs"),
    ("receiver ae title port storage scp", "Receiver", "Receiver"),
    ("pusher queue retry throughput push", "Pusher", "Pusher"),
]


def _universal_search_patients(query, limit=8):
    results = []
    with data_lock:
        rows = list(patient_data.items())
    q = query.lower()
    for pid, d in rows:
        haystack = _row_search_haystack(pid, d)
        if any(q in field.lower() for field in haystack):
            sub = f"{d.get('patient_name', '')} · {d.get('institution', '')} · {d.get('status', '')}"
            results.append({
                "label": f"{pid}",
                "sublabel": sub,
                "action": lambda p=pid: _jump_to_patient(p),
            })
            if len(results) >= limit:
                break
    return results


def _universal_search_reports(query, limit=5):
    results = []
    with data_lock:
        rows = list(patient_data.items())
    q = query.lower()
    for pid, d in rows:
        haystack = _row_search_haystack(pid, d)
        if not any(q in field.lower() for field in haystack):
            continue
        rpath = get_report_path(pid)
        if os.path.isfile(rpath):
            results.append({
                "label": f"Report — {pid}",
                "sublabel": d.get("patient_name", ""),
                "action": lambda p=rpath: open_document(p),
            })
            if len(results) >= limit:
                break
    return results


def _universal_search_logs(query, limit=5):
    if not admin_tabs_active["value"]:
        return []  # jsonl logs are an admin-only surface; don't imply access in User mode
    results = []
    for module in (RECEIVER_LOG, PUSH_LOG, AUDIT_LOG, APP_LOG):
        try:
            recs = get_structured_log_records(module, search_text=query, range_option="Last 7 Days", max_records=3)
        except Exception:
            recs = []
        for rec in recs:
            results.append({
                "label": f"{module} — {rec.get('event', 'entry')}",
                "sublabel": (rec.get("error_message") or rec.get("details") or rec.get("patient_id", ""))[:80],
                "action": lambda m=module: _jump_to_logs(m),
            })
            if len(results) >= limit:
                return results
    return results


def _universal_search_destinations(query, limit=6):
    results = []
    q = query.lower()
    for dest in load_destinations():
        haystack = f"{dest.get('name', '')} {dest.get('ae', '')} {dest.get('ip', '')} {dest.get('calling_ae', '')}".lower()
        if q in haystack:
            results.append({
                "label": dest.get("name", "(unnamed)"),
                "sublabel": f"{dest.get('ae', '')} @ {dest.get('ip', '')}:{dest.get('port', '')}",
                "action": lambda: _jump_to_tab("Destinations"),
            })
            if len(results) >= limit:
                break
    return results


def _universal_search_routing_rules(query, limit=6):
    results = []
    q = query.lower()
    for rule in load_routing_rules():
        haystack = " ".join(str(rule.get(k, "")) for k in ("modality", "institution", "source_ae", "destination")).lower()
        if q in haystack:
            label = " / ".join(filter(None, [rule.get("modality"), rule.get("institution"), rule.get("source_ae")])) or "(any)"
            results.append({
                "label": f"Rule: {label}",
                "sublabel": f"→ {rule.get('destination', '')}",
                "action": lambda: _jump_to_tab("Routing Rules"),
            })
            if len(results) >= limit:
                break
    return results


def _universal_search_settings(query, limit=6):
    results = []
    q = query.lower()
    for keywords, label, tab_name in SETTINGS_SEARCH_INDEX:
        if q in keywords or q in label.lower():
            # Admin-only settings tabs shouldn't appear as jumpable results
            # in User mode -- they don't exist to jump to yet.
            if tab_name in ADMIN_ONLY_TAB_NAMES and not admin_tabs_active["value"]:
                continue
            results.append({"label": label, "sublabel": "Settings", "action": lambda t=tab_name: _jump_to_tab(t)})
            if len(results) >= limit:
                break
    return results


def run_universal_search(query):
    query = query.strip()
    if len(query) < 2:
        return {}
    return {
        "Patients & Studies": _universal_search_patients(query),
        "Reports": _universal_search_reports(query),
        "Logs": _universal_search_logs(query),
        "Destinations": _universal_search_destinations(query),
        "Routing Rules": _universal_search_routing_rules(query),
        "Settings": _universal_search_settings(query),
    }


def _jump_to_tab(tab_name):
    try:
        tabview.set(tab_name)
        _sync_nav_highlight()
    except Exception:
        log_exception(f"Universal search: failed to jump to {tab_name}")


def _jump_to_patient(pid):
    """Jumps to Receiver (or Pusher, if that's where the study currently
    lives) and reuses that tab's own search box to home in on the patient
    -- no separate highlighting mechanism invented."""
    with data_lock:
        d = patient_data.get(pid, {})
    target_tab = "Pusher" if d.get("status") in (STATUS_SENDING, STATUS_SENT, STATUS_RETRYING, STATUS_QUEUED) else "Receiver"
    _jump_to_tab(target_tab)
    if target_tab == "Pusher":
        push_search_var.set(pid)
    else:
        rec_search_var.set(pid)


def _jump_to_logs(module):
    _jump_to_tab("Logs")
    try:
        log_file_var.set(module)
        log_display_mode_var.set("Structured")
    except Exception:
        pass


universal_search_state = {"window": None, "after_id": None}


def _on_universal_search_keyrelease():
    if universal_search_state["after_id"] is not None:
        app.after_cancel(universal_search_state["after_id"])
    universal_search_state["after_id"] = app.after(250, _run_and_show_universal_search)


def _close_universal_search_panel():
    win = universal_search_state.get("window")
    if win is not None and win.winfo_exists():
        win.destroy()
    universal_search_state["window"] = None


def _maybe_close_universal_search_panel():
    # Give a result-row click a chance to register before we tear the
    # panel down on focus loss -- also don't close if focus currently
    # sits anywhere inside the results panel itself (e.g. a slower
    # click still between button-press and button-release on a row;
    # CTkButton's command fires on release, not press).
    try:
        focused = app.focus_get()
        if focused is header_search_entry:
            return
        win = universal_search_state.get("window")
        if win is not None and win.winfo_exists() and focused is not None:
            w = focused
            while w is not None:
                if w == win:
                    return
                w = w.master
    except Exception:
        pass
    _close_universal_search_panel()


def _run_and_show_universal_search():
    query = header_search_var.get()
    grouped = run_universal_search(query)
    total = sum(len(v) for v in grouped.values())

    if not query.strip() or len(query.strip()) < 2:
        _close_universal_search_panel()
        return

    win = universal_search_state.get("window")
    if win is None or not win.winfo_exists():
        win = ctk.CTkToplevel(app)
        _keep_toplevel_small(win)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(fg_color=THEME_SURFACE)
        x = header_search_entry.winfo_rootx()
        y = header_search_entry.winfo_rooty() + header_search_entry.winfo_height() + 4
        x, y = _clamp_popover_geometry(x, y, 420, 420)
        win.geometry(f"420x420+{x}+{y}")
        scroll = ctk.CTkScrollableFrame(win, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)
        win.list_frame = scroll
        universal_search_state["window"] = win

    for child in win.list_frame.winfo_children():
        child.destroy()

    if total == 0:
        ctk.CTkLabel(win.list_frame, text=f"No results for “{query}”.",
                     font=get_font("small"), text_color=THEME_TEXT_MUTED).pack(pady=20)
        return

    for category, results in grouped.items():
        if not results:
            continue
        ctk.CTkLabel(win.list_frame, text=category.upper(), font=get_font("micro", "bold"),
                     text_color=THEME_TEXT_MUTED, anchor="w").pack(fill="x", padx=8, pady=(10, 2))
        for res in results:
            def _on_click(action=res["action"]):
                _close_universal_search_panel()
                action()

            row = ctk.CTkButton(
                win.list_frame, text=f"{res['label']}\n{res['sublabel']}" if res.get("sublabel") else res["label"],
                anchor="w", fg_color=THEME_HEADING_BG, hover_color=THEME_ACCENT_HOVER,
                text_color=THEME_TEXT, corner_radius=8, height=42,
                font=get_font("small"), command=_on_click,
            )
            row.pack(fill="x", padx=8, pady=2)


_ui_perf_stats = {"last_tick_time": None, "last_update_ms": 0.0, "avg_update_ms": 0.0, "tick_count": 0}


def periodic_refresh():
    _perf_tick_start = time.time()
    if _ui_perf_stats["last_tick_time"] is not None:
        pass  # actual UI-refresh-frequency is derived from get_refresh_interval_ms() + this timestamp
    _ui_perf_stats["last_tick_time"] = _perf_tick_start
    _refresh_counter[0] += 1
    _sync_nav_highlight()
    try:
        refresh_status_bar()
    except Exception:
        log_exception("Status bar refresh failed")
    try:
        refresh_header_status_chips()
    except Exception:
        log_exception("Header status chips refresh failed")
    try:
        if inspector_state["visible"] and inspector_current_pid["value"]:
            refresh_inspector(inspector_current_pid["value"])
    except Exception:
        log_exception("Inspector panel refresh failed")
    populate_tree(rec_tree, rec_search_var.get())
    populate_tree(push_tree, push_search_var.get())
    refresh_push_badge()
    if current_view["role"] == "user":
        refresh_receiver_badge()
    try:
        refresh_home_dashboard()
    except Exception:
        log_exception("Dashboard refresh failed")
    try:
        refresh_receiver_monitoring()
    except Exception:
        log_exception("Receiver monitoring refresh failed")
    try:
        refresh_pusher_monitoring()
    except Exception:
        log_exception("Pusher monitoring refresh failed")
    # Only do the dashboard's data_lock scan + disk_usage() syscall when
    # someone is actually looking at it, not on every tick regardless.
    if admin_tabs_active["value"] and tabview.get().startswith("Admin Dashboard"):
        refresh_admin_dashboard()
    # Same visibility gating -- computing this involves a log scan +
    # several psutil calls, so it only runs while someone's looking at it.
    if admin_tabs_active["value"] and tabview.get().startswith("Performance"):
        try:
            refresh_performance_tab()
        except Exception:
            log_exception("Performance tab refresh failed")
    if admin_tabs_active["value"] and tabview.get().startswith("Backup"):
        try:
            refresh_backup_history_ui()
        except Exception:
            log_exception("Backup history refresh failed")
    if admin_tabs_active["value"] and tabview.get().startswith("LDAP / AD"):
        try:
            refresh_ldap_roster_ui()
        except Exception:
            log_exception("LDAP roster refresh failed")
    if admin_tabs_active["value"] and tabview.get() == "Settings":
        try:
            refresh_settings_diagnostics()
        except Exception:
            log_exception("Settings diagnostics refresh failed")
    if admin_tabs_active["value"] and tabview.get() == "Doc Transfer":
        try:
            refresh_doc_transfer_tab()
        except Exception:
            log_exception("Doc Transfer tab refresh failed")
    # Cheap background guard for immutable log rotation/archival -- this
    # itself only actually stats the log files once every
    # _ROTATION_CHECK_INTERVAL_SEC seconds, so it's safe to call on every
    # tick without adding real overhead.
    try:
        rotate_logs_if_needed()
    except Exception:
        log_exception("Background log rotation check failed")
    try:
        sweep_trash_folder()
    except Exception:
        log_exception("Trash sweep failed")
    try:
        refresh_health_tree()
    except Exception:
        log_exception("PACS health tree refresh failed")
    try:
        refresh_next_backup_indicator()
    except Exception:
        log_exception("Next-backup indicator refresh failed")
    try:
        refresh_offline_queue_ui()
    except Exception:
        log_exception("Offline queue UI refresh failed")
    try:
        refresh_destination_filter_options()
    except Exception:
        log_exception("Destination filter refresh failed")
    try:
        refresh_export_patient_options()
    except Exception:
        log_exception("Export patient options refresh failed")
    try:
        check_admin_idle_timeout()
    except Exception:
        log_exception("Admin idle timeout check failed")
    elapsed_ms = (time.time() - _perf_tick_start) * 1000.0
    _ui_perf_stats["last_update_ms"] = elapsed_ms
    n = _ui_perf_stats["tick_count"] = _ui_perf_stats["tick_count"] + 1
    prev_avg = _ui_perf_stats["avg_update_ms"]
    _ui_perf_stats["avg_update_ms"] = prev_avg + (elapsed_ms - prev_avg) / min(n, 50)
    app.after(get_refresh_interval_ms(), periodic_refresh)

# =========================================================
# ADMIN MODE MANAGEMENT
# =========================================================

def apply_receiver_mode_visibility():
    """Show/hide the Receiver-tab widgets that are Admin-only per spec:
    AE Title field, Port field, Start/Stop Receiver Server button, and the
    Autoroute checkbox. Everything else on the Receiver tab (worklist,
    search/filter/sort, Import Folder, Document/Report Manager access,
    Reset-to-Pending, View full error) stays available in both modes."""
    is_admin = current_view["role"] == "admin"
    if is_admin:
        for w in _receiver_admin_only_grid_widgets:
            w.grid()
        for w in _receiver_admin_only_pack_widgets:
            w.pack(pady=4)
    else:
        for w in _receiver_admin_only_grid_widgets:
            w.grid_remove()
        for w in _receiver_admin_only_pack_widgets:
            w.pack_forget()
    rec_badge_label.pack_forget()
    if not is_admin:
        rec_badge_label.pack(anchor="w", padx=12, pady=(0, 4))
        refresh_receiver_badge()


def update_admin_bar_ui():
    """Refreshes the always-visible admin bar to match admin_session /
    current_view state."""
    if not admin_session["unlocked"]:
        admin_bar_status_lbl.configure(text="Mode: USER")
        admin_bar_status_dot.configure(text_color=THEME_TEXT_MUTED)
        admin_login_btn.pack(side="left", padx=4, pady=6)
        admin_switch_view_btn.pack_forget()
        admin_change_pin_btn.pack_forget()
        admin_lock_btn.pack_forget()
        return

    admin_login_btn.pack_forget()
    admin_change_pin_btn.pack(side="left", padx=4, pady=6)
    admin_lock_btn.pack(side="left", padx=4, pady=6)

    if current_view["role"] == "admin":
        admin_bar_status_lbl.configure(text="Mode: ADMIN")
        admin_bar_status_dot.configure(text_color=THEME_SUCCESS)
        admin_switch_view_btn.configure(text="Switch to User View")
        admin_switch_view_btn.pack(side="left", padx=4, pady=6, before=admin_change_pin_btn)
    else:
        admin_bar_status_lbl.configure(text="Mode: USER (Admin preview)")
        admin_bar_status_dot.configure(text_color=THEME_ACCENT)
        admin_switch_view_btn.configure(text="Switch to Admin View")
        admin_switch_view_btn.pack(side="left", padx=4, pady=6, before=admin_change_pin_btn)


def set_view(role):
    """(Re)builds the tab set for `role` ("admin" or "user") and updates all
    the mode-dependent UI. This is the single place that adds/removes the
    Admin-only tabs, per spec ("rebuild/show-hide tabs rather than
    duplicating tab-construction code")."""
    current_view["role"] = role
    if role == "admin" and not admin_tabs_active["value"]:
        build_admin_only_tabs()
        tabview.set("Admin Dashboard")
    elif role == "user" and admin_tabs_active["value"]:
        tear_down_admin_only_tabs()
        tabview.set("Receiver")
    apply_receiver_mode_visibility()
    update_admin_bar_ui()


_admin_login_dialog_ref = {"win": None}  # §fix: see do_admin_login()


def do_admin_login():
    # §fix (admin login instability): re-clicking "Admin Login" while a
    # login dialog is already open used to always create a brand new
    # Toplevel. Normally the existing dialog's grab_set() blocks that
    # click -- but grab_set() can silently fail to actually take the grab
    # (see _safe_grab_set()), and a slow/unmapped window can leave a
    # moment where the background button is still clickable. A second,
    # independent Toplevel each with its own grab then left the app in a
    # confusing state that looked like "admin login stops working" over
    # an extended session. If a login dialog is already open, just bring
    # it to the front instead of stacking another one.
    existing = _admin_login_dialog_ref["win"]
    if existing is not None:
        try:
            if existing.winfo_exists():
                existing.lift()
                existing.focus_force()
                return
        except Exception:
            pass
        _admin_login_dialog_ref["win"] = None

    ldap_cfg = load_ldap_config()
    show_ldap_tab = ldap_cfg.get("enabled", False) and LDAP3_AVAILABLE

    win = ctk.CTkToplevel(app)
    _admin_login_dialog_ref["win"] = win
    _keep_toplevel_small(win)
    win.title("Admin Login")
    win.geometry("380x300" if show_ldap_tab else "340x180")
    win.transient(app)
    _safe_grab_set(win)

    def _do_pin_login(pin_entry, err_lbl):
        pin = pin_entry.get().strip()
        if verify_admin_pin(pin):
            write_audit_log("ADMIN-LOGIN-SUCCESS", "Admin PIN accepted")
            current_identity.update(username=None, display_name=None, source="local", ldap_role=None)
            admin_session["unlocked"] = True
            _record_ui_activity()  # D.3: don't start the idle clock already expired
            win.destroy()
            set_view("admin")
            show_toast_threadsafe("Admin Unlocked", "You are now viewing the app in Admin mode.")
        else:
            write_audit_log("ADMIN-LOGIN-FAILED", "Incorrect Admin PIN entered")
            notify_event("authentication_failures", "Authentication Failure", "Incorrect Admin PIN entered")
            err_lbl.configure(text="Incorrect PIN.")
            pin_entry.delete(0, "end")

    def _do_domain_login(user_entry, pw_entry, err_lbl):
        username = user_entry.get().strip()
        password = pw_entry.get()
        if not username or not password:
            err_lbl.configure(text="Enter both username and password.")
            return
        ok, message, info = ldap_authenticate(username, password)
        if not ok:
            err_lbl.configure(text=message)
            pw_entry.delete(0, "end")
            return
        current_identity.update(
            username=info["username"], display_name=info["display_name"],
            source="ldap", ldap_role=info["role"],
        )
        win.destroy()
        if info["role"] == "Administrators":
            admin_session["unlocked"] = True
            _record_ui_activity()  # D.3: don't start the idle clock already expired
            set_view("admin")
            show_toast_threadsafe("Admin Unlocked", f"{message} You are now viewing the app in Admin mode.")
        else:
            show_toast_threadsafe(
                "Domain Login Successful",
                f"{message} Your role ({info['role']}) uses standard User mode in this app.",
            )

    if not show_ldap_tab:
        ctk.CTkLabel(win, text="Enter Admin PIN", font=get_font("section", "bold")).pack(pady=(16, 8))
        pin_entry = ctk.CTkEntry(win, width=200, show="•")
        pin_entry.pack(pady=4)
        pin_entry.focus_set()
        err_lbl = ctk.CTkLabel(win, text="", text_color=THEME_DANGER)
        err_lbl.pack(pady=4)
        pin_entry.bind("<Return>", lambda e: _do_pin_login(pin_entry, err_lbl))
        ctk.CTkButton(win, text="Unlock", command=lambda: _do_pin_login(pin_entry, err_lbl)).pack(pady=10)
        return

    # ---- LDAP is enabled: offer both Local Admin PIN and Domain Login ----
    login_tabs = ctk.CTkTabview(win, width=340, height=250)
    login_tabs.pack(padx=16, pady=16, fill="both", expand=True)
    tab_pin = login_tabs.add("Local Admin PIN")
    tab_domain = login_tabs.add("Domain Login")

    ctk.CTkLabel(tab_pin, text="Enter Admin PIN", font=get_font("caption", "bold")).pack(pady=(14, 8))
    pin_entry = ctk.CTkEntry(tab_pin, width=200, show="•")
    pin_entry.pack(pady=4)
    pin_err_lbl = ctk.CTkLabel(tab_pin, text="", text_color=THEME_DANGER)
    pin_err_lbl.pack(pady=4)
    pin_entry.bind("<Return>", lambda e: _do_pin_login(pin_entry, pin_err_lbl))
    ctk.CTkButton(tab_pin, text="Unlock", command=lambda: _do_pin_login(pin_entry, pin_err_lbl)).pack(pady=10)

    ctk.CTkLabel(tab_domain, text="Domain Username", font=get_font("body")).pack(pady=(14, 2))
    domain_user_entry = ctk.CTkEntry(tab_domain, width=220)
    domain_user_entry.pack(pady=2)
    ctk.CTkLabel(tab_domain, text="Password", font=get_font("body")).pack(pady=(8, 2))
    domain_pw_entry = ctk.CTkEntry(tab_domain, width=220, show="•")
    domain_pw_entry.pack(pady=2)
    domain_err_lbl = ctk.CTkLabel(tab_domain, text="", text_color=THEME_DANGER, wraplength=280)
    domain_err_lbl.pack(pady=6)
    domain_pw_entry.bind("<Return>", lambda e: _do_domain_login(domain_user_entry, domain_pw_entry, domain_err_lbl))
    ctk.CTkButton(tab_domain, text="Log In",
                 command=lambda: _do_domain_login(domain_user_entry, domain_pw_entry, domain_err_lbl)).pack(pady=8)
    domain_user_entry.focus_set()


def do_switch_view_toggle():
    """Admin-only preview toggle: switch the tab set between Admin and User
    without re-entering the PIN (single click back, per spec)."""
    if not admin_session["unlocked"]:
        return
    set_view("user" if current_view["role"] == "admin" else "admin")


def do_admin_lock():
    admin_session["unlocked"] = False
    write_audit_log("ADMIN-LOGOUT", "Admin session locked")
    current_identity.update(username=None, display_name=None, source="local", ldap_role=None)
    set_view("user")


def do_change_admin_pin():
    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title("Change Admin PIN")
    win.geometry("360x320")
    win.transient(app)
    _safe_grab_set(win)

    ctk.CTkLabel(win, text="Change Admin PIN", font=get_font("section", "bold")).pack(pady=(16, 10))
    ctk.CTkLabel(win, text="Current PIN").pack(anchor="w", padx=30)
    cur_entry = ctk.CTkEntry(win, width=240, show="•")
    cur_entry.pack(padx=30, pady=4)
    ctk.CTkLabel(win, text=f"New PIN ({ADMIN_PIN_MIN_DIGITS}+ digits)").pack(anchor="w", padx=30, pady=(10, 0))
    new_entry = ctk.CTkEntry(win, width=240, show="•")
    new_entry.pack(padx=30, pady=4)
    ctk.CTkLabel(win, text="Confirm New PIN").pack(anchor="w", padx=30, pady=(10, 0))
    confirm_entry = ctk.CTkEntry(win, width=240, show="•")
    confirm_entry.pack(padx=30, pady=4)
    err_lbl = ctk.CTkLabel(win, text="", text_color=THEME_DANGER, wraplength=300)
    err_lbl.pack(pady=8)

    def submit():
        if not verify_admin_pin(cur_entry.get().strip()):
            write_audit_log("ADMIN-PIN-CHANGE-FAILED", "Current PIN did not match")
            err_lbl.configure(text="Current PIN is incorrect.")
            return
        new_pin = new_entry.get().strip()
        ok, msg = validate_pin_format(new_pin)
        if not ok:
            err_lbl.configure(text=msg)
            return
        if new_pin != confirm_entry.get().strip():
            err_lbl.configure(text="New PIN entries do not match.")
            return
        save_admin_pin(new_pin)
        write_audit_log("ADMIN-PIN-CHANGED", "Admin PIN changed successfully")
        win.destroy()
        modern_showinfo("PIN Changed", "The Admin PIN has been updated.")

    ctk.CTkButton(win, text="Save New PIN", command=submit).pack(pady=10)


def show_set_admin_pin_dialog():
    """One-time, blocking 'Set Admin PIN' dialog shown before the main
    window opens (app is withdrawn until this returns). No PIN can be set
    that doesn't meet the numeric/6-digit-minimum policy, and the dialog
    cannot be dismissed without setting one."""
    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title("Set Admin PIN")
    win.geometry("380x340")
    win.protocol("WM_DELETE_WINDOW", lambda: None)  # must complete setup first
    win.lift()
    win.attributes("-topmost", True)
    win.after(200, lambda: win.attributes("-topmost", False))
    win.focus_force()
    _safe_grab_set(win)

    ctk.CTkLabel(win, text="Welcome — set an Admin PIN",
                 font=get_font("section_lg", "bold")).pack(pady=(20, 4))
    ctk.CTkLabel(win, text="This PIN protects Admin-only settings.\n"
                           f"Numeric, at least {ADMIN_PIN_MIN_DIGITS} digits.",
                 font=get_font("small"), text_color=THEME_TEXT_MUTED, justify="center").pack(pady=(0, 14))

    ctk.CTkLabel(win, text="New PIN").pack(anchor="w", padx=40)
    pin1 = ctk.CTkEntry(win, width=260, show="•")
    pin1.pack(padx=40, pady=4)
    ctk.CTkLabel(win, text="Confirm PIN").pack(anchor="w", padx=40, pady=(10, 0))
    pin2 = ctk.CTkEntry(win, width=260, show="•")
    pin2.pack(padx=40, pady=4)
    err_lbl = ctk.CTkLabel(win, text="", text_color=THEME_DANGER, wraplength=300)
    err_lbl.pack(pady=10)

    def submit():
        p1, p2 = pin1.get().strip(), pin2.get().strip()
        ok, msg = validate_pin_format(p1)
        if not ok:
            err_lbl.configure(text=msg)
            return
        if p1 != p2:
            err_lbl.configure(text="PIN entries do not match.")
            return
        save_admin_pin(p1)
        write_audit_log("ADMIN-PIN-SET", "Initial Admin PIN configured on first run")
        win.destroy()

    ctk.CTkButton(win, text="Set PIN", command=submit).pack(pady=12)
    win.wait_window()  # block startup() until this completes


def ensure_admin_pin_configured():
    if not admin_pin_is_configured():
        show_set_admin_pin_dialog()


admin_login_btn.configure(command=do_admin_login)
admin_switch_view_btn.configure(command=do_switch_view_toggle)
admin_change_pin_btn.configure(command=do_change_admin_pin)
admin_lock_btn.configure(command=do_admin_lock)


# ---- Part 4: quick-filter buttons, keyboard shortcuts ----
rec_select_failed_btn.configure(command=lambda: select_all_by_status(rec_tree, {STATUS_FAILED}))
rec_select_pending_btn.configure(command=lambda: select_all_by_status(rec_tree, {STATUS_PENDING}))
push_select_failed_btn.configure(command=lambda: select_all_by_status(push_tree, {STATUS_FAILED}))
push_select_pending_btn.configure(command=lambda: select_all_by_status(push_tree, {STATUS_PENDING}))


def _focus_active_search(event=None):
    """Ctrl+F: focus the search box for whichever tab (Receiver/Pusher) is
    currently active."""
    current = tabview.get()
    if current.startswith("Pusher"):
        push_search_entry.focus_set()
    else:
        rec_search_entry.focus_set()
    return "break"


def _manual_refresh_shortcut(event=None):
    """F5: trigger the same manual worklist refresh the Refresh buttons do."""
    refresh_worklists(rec_search_var.get())
    if admin_tabs_active["value"]:
        refresh_admin_dashboard()
    return "break"


app.bind_all("<Control-f>", _focus_active_search)
app.bind_all("<F5>", _manual_refresh_shortcut)
# D.3: any mouse/keyboard activity resets the admin idle-timeout clock.
app.bind_all("<Motion>", _record_ui_activity, add="+")
app.bind_all("<KeyPress>", _record_ui_activity, add="+")
app.bind_all("<Button>", _record_ui_activity, add="+")

# =========================================================
# STARTUP
# =========================================================

def startup():
    """Run all initialisation that needs the GUI to exist first."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    generate_key()

    try:
        apply_logging_level()  # Settings > Logging & Diagnostics > Logging Level
    except Exception:
        pass

    # §fix: opt-in only -- see the note near the pynetdicom imports for why
    # this is no longer called unconditionally at module import time.
    if APP_SETTINGS.get("verbose_dicom_protocol_logging"):
        try:
            debug_logger()
            app_logger.info(
                "Verbose DICOM protocol logging (pynetdicom debug_logger) is ON -- "
                "full PDU/DIMSE traces will be printed for every association. "
                "Turn off in Settings when done diagnosing."
            )
        except Exception:
            log_exception("Failed to enable verbose DICOM protocol logging")

    if not DECODE_CODECS_AVAILABLE:
        app_logger.warning(
            "No pixel-data DECODER installed (missing: %s, and python-gdcm not found "
            "either). Receiving or pushing JPEG2000/JPEG-LS/RLE-encoded files will fail "
            "outright wherever transcoding is needed. Install with: pip install %s "
            "(or: pip install python-gdcm)",
            ", ".join(MISSING_CODEC_PLUGINS), " ".join(MISSING_CODEC_PLUGINS),
        )
    elif not ENCODE_CODECS_AVAILABLE:
        app_logger.warning(
            "No pixel-data ENCODER installed (missing: %s). python-gdcm can decode "
            "compressed files but cannot re-compress them, so pushing an uncompressed "
            "file to a destination that only accepts a compressed transfer syntax will "
            "fail to transcode. Install with: pip install %s",
            ", ".join(MISSING_CODEC_PLUGINS), " ".join(MISSING_CODEC_PLUGINS),
        )

    # The real data loading happens here, BEFORE the window is revealed,
    # so the splash screen (still the only visible thing) reflects actual
    # work instead of being dismissed the instant the (empty) window
    # appears. Each _splash_step() only fires once its preceding call has
    # actually returned -- a slow step (e.g. a large worklist CSV) will
    # visibly hold the bar at that step's label rather than faking progress.
    try:
        _splash_step(0.85, "Loading saved studies…")
        load_csv()
        _migrate_legacy_push_config()

        _splash_step(0.90, "Loading configuration…")
        _load_receiver_config()
        refresh_destinations_ui()
        refresh_routing_ui()
        load_sop_into_editor()

        _splash_step(0.95, "Populating worklists…")
        populate_tree(rec_tree)
        populate_tree(push_tree)
        refresh_log_view()

        # App always launches in User mode (admin_session starts locked).
        apply_receiver_mode_visibility()
        update_admin_bar_ui()

        _splash_step(1.0, "Ready.")
    except Exception:
        log_exception("Startup failed before window was shown")
        # Fall through and show the window anyway -- an error dialog with
        # no visible parent is worse than a window that opens with some
        # data missing (the exception is already logged either way).

    _close_splash_screen()

    # Show the main window. A Toplevel dialog parented to a still-
    # withdrawn root can fail to display or grab focus on some platforms,
    # which left the app stuck invisibly waiting on a PIN dialog no one
    # could see (looked like "running but nothing opens"). Deiconifying
    # before showing the PIN dialog fixes that.
    app.deiconify()
    if APP_SETTINGS.get("launch_maximized", False):
        try:
            app.state("zoomed")
        except Exception:
            try:
                app.attributes("-zoomed", True)
            except Exception:
                pass
    elif APP_SETTINGS.get("remember_window_geometry", True):
        try:
            geom = APP_SETTINGS.get("last_window_geometry")
            if geom:
                app.geometry(geom)
        except Exception:
            pass
    app.lift()
    app.attributes("-topmost", True)
    app.after(200, lambda: app.attributes("-topmost", False))
    app.update()

    # One-time Admin PIN setup (now shown on top of a visible main window).
    ensure_admin_pin_configured()

    try:
        if APP_SETTINGS.get("auto_start_receiver_on_launch", False):
            try:
                do_start_receiver()
            except Exception:
                log_exception("Failed to auto-start receiver on launch")

        # Check disk on startup
        check_disk_space_and_warn()

        # Continuous PACS health monitoring runs regardless of admin/user
        # mode -- it just populates destination_health_cache in the
        # background; the PACS Health tab (admin-only) and the landing
        # Dashboard both read from that same cache.
        start_pacs_health_monitor_thread()
        start_offline_queue_worker_thread()
        start_backup_scheduler_thread()
        start_report_email_scheduler_thread()
        # Self-heals any image/document counts in the loaded worklist CSV
        # that were inflated by duplicate transfers under the old counting
        # logic -- see reconcile_patient_image_counts(). Runs off the UI
        # thread so a large worklist doesn't delay showing the window.
        start_count_reconciliation_thread()

        write_audit_log("APP_START", f"pid={os.getpid()}")
        app_logger.info("Application started (pid=%d)", os.getpid())
    except Exception:
        log_exception("Startup failed after window was shown")
        modern_showerror(
            "Startup error",
            "The app window is open, but part of startup failed.\n"
            "Check app.log for details.\n\n" + traceback.format_exc()[-800:]
        )


# =========================================================
# ENTRYPOINT
# =========================================================

def _focus_universal_search(_event=None):
    header_search_entry.focus_set()
    header_search_entry.select_range(0, "end")


def _global_escape_handler(_event=None):
    """Escape closes whichever overlay panel is currently open, from
    anywhere in the app -- not just while the search box itself has
    focus (see header_search_entry's own <Escape> binding above)."""
    _close_universal_search_panel()
    win = notification_panel_state.get("window")
    if win is not None and win.winfo_exists():
        win.destroy()
        notification_panel_state["window"] = None


app.bind_all("<Control-f>", _focus_universal_search)
app.bind_all("<Control-F>", _focus_universal_search)
app.bind("<Escape>", _global_escape_handler)  # main window only -- won't interfere with other dialogs' own Escape handling

if __name__ == "__main__":
    try:
        startup()
    except Exception:
        log_exception("Fatal error during startup")
        try:
            _close_splash_screen()
        except Exception:
            pass
        try:
            app.deiconify()
        except Exception:
            pass
        modern_showerror(
            "Startup error",
            "A fatal error occurred during startup. Check app.log for details.\n\n"
            + traceback.format_exc()[-800:]
        )
    app.after(500, pump_events)
    app.after(get_refresh_interval_ms(), periodic_refresh)
    app.mainloop()
