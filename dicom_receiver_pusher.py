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

try:
    import io as _icon_io
    import base64 as _icon_b64
    from lucide_icons_data import LUCIDE_ICONS_B64
    LUCIDE_ICONS_AVAILABLE = True
except Exception:
    LUCIDE_ICONS_AVAILABLE = False
    LUCIDE_ICONS_B64 = {}
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

from pynetdicom import AE, evt, build_role
from pynetdicom.sop_class import (
    PatientRootQueryRetrieveInformationModelFind,
    StudyRootQueryRetrieveInformationModelFind,
    PatientRootQueryRetrieveInformationModelMove,
    StudyRootQueryRetrieveInformationModelMove,
)

from pydicom.uid import UID
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
THEME_CONFIG_FILE = "theme_config.json"  # not secret -- Dark/Light/System preference
VIEW_OPTIONS_FILE = "view_options.json"  # not secret -- column visibility + table density
BACKUP_DIR = "backups"  # not secret -- holds full application backup ZIPs
BACKUP_HISTORY_FILE = "backup_history.json"  # not secret -- backup run metadata
BACKUP_SCHEDULE_CONFIG_FILE = "backup_schedule_config.json"  # not secret -- scheduled-backup settings
LDAP_CONFIG_FILE = "ldap_config.enc"  # encrypted: holds the service-account bind password
LDAP_USERS_FILE = "ldap_users.json"  # not secret -- imported directory roster (no passwords)
ROUTING_RULES_FILE = "routing_rules.json"  # not secret -- just routing logic
TRANSFER_CHECKPOINTS_FILE = "transfer_checkpoints.json"  # resume-from-failure state (not secret)
OFFLINE_QUEUE_FILE = "offline_queue.json"  # persistent offline push queue (not secret)
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
}

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


def _roll_daily_stats_if_needed():
    today = datetime.date.today().isoformat()
    with _dashboard_stats_lock:
        if daily_stats["date"] != today:
            daily_stats["date"] = today
            daily_stats["studies_received_uids"] = set()
            daily_stats["images_received"] = 0
            daily_stats["documents_received"] = 0
            daily_stats["studies_sent_uids"] = set()
            daily_stats["images_sent"] = 0
            daily_stats["failed_transfers"] = 0


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

app_shutdown_event = threading.Event()
DEFAULT_SOP_CLASSES = {
    # Curated down to exactly 128 entries -- DICOM's hard per-association
    # presentation-context limit (see MAX_PRESENTATION_CONTEXTS below). At
    # 128 entries, every SOP Class in this default set is always offered;
    # cap_sop_list() never has to silently truncate anything. Only the
    # least-used entry (ImplantTemplateGroupStorage, a niche implant-
    # planning SOP) was dropped to make room -- everything commonly seen
    # in general radiology/PACS traffic (CT/MR/US/XA/NM/PET/RT/SR/
    # waveforms/presentation states/secondary capture/etc.) is kept. Add
    # it back via the in-GUI SOP editor if your site actually needs it.
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
    "HighThroughputJPEG2000": "1.2.840.10008.1.2.4.203",
}


def create_sop_ini():
    """Generate a full sopclass.ini (SOP Classes + Transfer Syntaxes) the
    first time the app runs. Both the Receiver and the Pusher read this
    SAME file, so negotiated contexts always match on both sides. Users
    may hand-edit this file later to add/remove entries."""
    if os.path.exists(SOP_INI):
        return
    config = configparser.ConfigParser()
    config["SOP_CLASSES"] = DEFAULT_SOP_CLASSES
    config["TRANSFER_SYNTAXES"] = DEFAULT_TRANSFER_SYNTAXES
    with open(SOP_INI, "w") as f:
        config.write(f)


def load_extra_sops():
    """Load every SOP Class UID listed under [SOP_CLASSES] in sopclass.ini."""
    create_sop_ini()
    config = configparser.ConfigParser()
    config.read(SOP_INI)
    sops = []
    if "SOP_CLASSES" in config:
        for _, uid in config["SOP_CLASSES"].items():
            try:
                sops.append(UID(uid.strip()))
            except Exception:
                pass
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
        for _, uid in config["TRANSFER_SYNTAXES"].items():
            try:
                ts.append(UID(uid.strip()))
            except Exception:
                pass
    if not ts:
        ts = [UID("1.2.840.10008.1.2"), UID("1.2.840.10008.1.2.1")]
    return ts


def save_sop_ini(sop_classes: dict, transfer_syntaxes: dict):
    """Persist an edited SOP Class / Transfer Syntax set back to sopclass.ini.
    Used by the in-GUI SOP editor."""
    config = configparser.ConfigParser()
    config["SOP_CLASSES"] = sop_classes
    config["TRANSFER_SYNTAXES"] = transfer_syntaxes
    with open(SOP_INI, "w") as f:
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
        with open(KEY_FILE, "wb") as f:
            f.write(Fernet.generate_key())


def load_key():
    with open(KEY_FILE, "rb") as f:
        return f.read()


def encrypt_and_save(file_name, data):
    fernet = Fernet(load_key())
    with open(file_name, "wb") as f:
        f.write(fernet.encrypt(data.encode()))


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
    with open(ADMIN_AUTH_FILE, "w", encoding="utf-8") as f:
        json.dump(record, f, indent=2)


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
    with open(TLS_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


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


def build_ssl_context_for_client(tls_cfg, port=None):
    """Build an ssl.SSLContext for outbound (pusher) associations. Returns
    None if TLS isn't needed (plaintext, default behaviour for everything
    except known TLS ports) or if context creation fails.
    TLS is used when EITHER tls_cfg['enabled'] is True (manual override,
    applies to every destination) OR `port` equals DICOM_TLS_PORT (2762 --
    auto-detected per connection so destinations living on the standard
    DICOM-TLS port work without touching config, while everything else on
    a normal port is untouched and stays plaintext)."""
    use_tls = bool(tls_cfg.get("enabled")) or (port is not None and int(port) == DICOM_TLS_PORT)
    if not use_tls:
        return None
    try:
        import ssl
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        if tls_cfg.get("ca_cert") and os.path.exists(tls_cfg["ca_cert"]):
            ctx.load_verify_locations(cafile=tls_cfg["ca_cert"])
        else:
            ctx.check_hostname = False
            ctx.verify_mode = __import__("ssl").CERT_NONE
        if tls_cfg.get("require_mutual_tls") and tls_cfg.get("cert") and tls_cfg.get("key"):
            ctx.load_cert_chain(certfile=tls_cfg["cert"], keyfile=tls_cfg["key"])
        return ctx
    except Exception:
        log_exception("Failed to build client TLS context; starting in plaintext mode.")
        return None

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
]

CONFIG_CHANGE_AUDIT_EVENTS = {
    "DESTINATIONS-CHANGED", "ROUTING-RULES-CHANGED", "SOP-CONFIG-CHANGED",
    "LOG-RETENTION-CHANGED", "NOTIFICATIONS-CONFIG-CHANGED",
    "ADMIN-PIN-CHANGED", "ADMIN-PIN-SET", "TLS-CONFIG-CHANGED",
    "THEME-CHANGED", "LDAP-CONFIG-CHANGED", "BANDWIDTH-CONFIG-CHANGED", "BACKUP-SCHEDULE-CHANGED",
}

DEFAULT_NOTIFICATIONS_CONFIG = {
    "channels": {"desktop": True, "email": False, "webhook": False},
    "events": {key: (key != "configuration_changes") for key, _ in NOTIFICATION_EVENT_TYPES},
    "email": {
        "smtp_host": "", "smtp_port": 587, "smtp_user": "", "smtp_password": "",
        "from_addr": "", "to_addrs": "", "use_tls": True,
    },
    "webhook": {"url": ""},
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
            for section in ("channels", "events", "email", "webhook"):
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
    with open(ROUTING_RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(rules, f, indent=2)
    write_audit_log("ROUTING-RULES-CHANGED", f"count={len(rules)}")


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
        if rule.get("modality") and rule["modality"].upper() != modality.upper():
            continue
        if rule.get("institution") and rule["institution"].lower() != institution.lower():
            continue
        if rule.get("source_ae") and rule["source_ae"].upper() != source_ae.upper():
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
            with open(TRANSFER_CHECKPOINTS_FILE, "w", encoding="utf-8") as f:
                json.dump(_checkpoint_cache["value"] or {}, f, indent=2)
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
            with open(OFFLINE_QUEUE_FILE, "w", encoding="utf-8") as f:
                json.dump(_offline_queue_cache["value"] or [], f, indent=2)
    except Exception:
        log_exception("Failed to save offline_queue.json")


def enqueue_offline(pid, destination_name, reason):
    """Adds (or updates) a patient in the persistent offline queue.
    Preserves original queue position/queued_at if the patient was
    already queued -- re-failing doesn't push it to the back of the line."""
    with _offline_queue_lock:
        items = load_offline_queue()
        for item in items:
            if item["pid"] == pid:
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
        write_audit_log("QUEUE-CHANGED", f"pid={pid} action=enqueued dest={destination_name}")


def dequeue_offline(pid):
    with _offline_queue_lock:
        items = load_offline_queue()
        removed = next((i for i in items if i["pid"] == pid), None)
        remaining = [i for i in items if i["pid"] != pid]
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
                        try:
                            sent, total, ok = push_single_patient(item["pid"], destination=dest)
                        except Exception:
                            log_exception(f"Offline queue retry failed for {item['pid']}")
                            ok = False
                        if ok:
                            dequeue_offline(item["pid"])
                            notify_event("push_complete", "Offline Queue: Push Succeeded",
                                        f"{item['pid']} delivered to {item['destination_name']} after being queued.")
                        else:
                            with _offline_queue_lock:
                                items = load_offline_queue()
                                for it in items:
                                    if it["pid"] == item["pid"]:
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
        with open(BANDWIDTH_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        _bandwidth_config_cache["value"] = cfg
        bandwidth_limiter.reload()
        write_audit_log("BANDWIDTH-CONFIG-CHANGED",
                        f"preset={cfg.get('preset')} custom_mbps={cfg.get('custom_mbps')} "
                        f"effective_mbps={get_effective_bandwidth_mbps(cfg)}")
    except Exception:
        log_exception("Failed to save bandwidth_config.json")
        raise


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
# BUILT-IN ZIP EXPORT
# =========================================================
# Supports exporting an entire patient, a single study, a single series,
# an explicit list of files, or the entire worklist, each optionally as
# a password-protected (AES-256, via pyzipper) ZIP, and optionally
# bundling the patient's report, a filtered log excerpt, a metadata
# summary, and/or a real DICOMDIR (via pydicom's FileSet). Runs entirely
# in a background thread with progress callbacks -- never blocks the UI.

def get_patient_folder(pid):
    return os.path.join(OUTPUT_DIR, pid)


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
                arcname = f"{pid}/{os.path.basename(fpath)}"
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
                        zf_write_file(f"{pid}/Reports/{os.path.basename(rpath)}", rpath)

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
                                severity="All", range_option="All Time", max_records=500):
    """Structured (parsed-jsonl) equivalent of the Logs tab's raw text
    tail. Reuses _iter_jsonl_records_in_range exactly as the PDF Report
    feature already does (live file + rotated archives), so this is a
    second consumer of existing infrastructure, not a new log source."""
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
    images_received = len(ok_receives)

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

    top_modalities = Counter(r.get("modality") for r in ok_receives if r.get("modality")).most_common(5)
    top_institutions = Counter(r.get("institution") for r in ok_receives if r.get("institution")).most_common(5)

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
        dest_rows = [["Destination", "Sent", "Failed"]] + [
            [n, str(stats["destination_stats"][n]["sent"]), str(stats["destination_stats"][n]["failed"])]
            for n in dest_names
        ]
    else:
        dest_rows = [["Destination", "Sent", "Failed"], ["(no push activity in this date range)", "0", "0"]]
    dest_table = Table(dest_rows, colWidths=[3 * inch, 1.5 * inch, 1.5 * inch])
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
                          THEME_CONFIG_FILE, LOG_RETENTION_CONFIG_FILE, ADMIN_AUTH_FILE],
        "Encryption Keys": [KEY_FILE],
        "Routing Rules": [ROUTING_RULES_FILE],
        "Destination Profiles": [DESTINATIONS_CONFIG, PUSH_CONFIG, RECEIVER_CONFIG],
        "TLS Configuration": [TLS_CONFIG_FILE],
        "CSV Database": [CSV_FILE],
        "Logs": [RECEIVER_LOG, PUSH_LOG, APP_LOG, RECEIVER_LOG_JSONL, PUSH_LOG_JSONL, APP_LOG_JSONL],
        "Audit Logs": [AUDIT_LOG, AUDIT_LOG_JSONL],
        "User Settings": [THEME_CONFIG_FILE],
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
        with open(BACKUP_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)
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
    with open(BACKUP_SCHEDULE_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
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
        with open(LDAP_USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=2)
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
    users[username] = {
        "username": username, "display_name": display_name, "email": email,
        "role": role, "last_login": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    _save_ldap_users(users)


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



# Two pieces: the palette + persistence defined here (early, since the GUI
# construction below reads the saved preference before the window is even
# built), and apply_theme() / the "Theme:" selector callback defined later
# once the actual widgets (app, header_bar, admin_bar, tabview, and every
# ttk Treeview) exist to be recolored.

THEME_PALETTES = {
    "dark": {
        "bg": "#0e1015", "surface": "#181b21", "heading_bg": "#232733",
        "accent": "#2f8eff", "accent_hover": "#1f6fd6", "text": "#e8eaed",
        "text_muted": "#8b93a7", "neutral_btn": "#3a3f4c", "neutral_btn_hover": "#4b5263",
        "danger": "#f04747", "danger_hover": "#c0392b", "success": "#2ecc71", "success_hover": "#27ae60",
        "odd_row": "#1d212a", "stale": "#ff5555", "search_highlight": "#3a3510",
        "segmented_hover": "#2c3140",
    },
    "light": {
        "bg": "#f4f5f7", "surface": "#ffffff", "heading_bg": "#e7e9ee",
        "accent": "#2f6fe0", "accent_hover": "#2558bd", "text": "#1b1f27",
        "text_muted": "#5b6472", "neutral_btn": "#e2e5ea", "neutral_btn_hover": "#d3d7de",
        "danger": "#e74c3c", "danger_hover": "#c0392b", "success": "#27ae60", "success_hover": "#219150",
        "odd_row": "#f0f1f4", "stale": "#c0392b", "search_highlight": "#fff6cc",
        "segmented_hover": "#d8dbe2",
    },
}


def load_theme_preference():
    """Returns the saved preference exactly as the UI stores it:
    'Dark', 'Light', or 'System'. Defaults to 'Dark' (this app's
    original, unchanged look) if nothing has been saved yet."""
    try:
        if os.path.exists(THEME_CONFIG_FILE):
            with open(THEME_CONFIG_FILE, "r", encoding="utf-8") as f:
                mode = json.load(f).get("mode", "Dark")
            if mode in ("Dark", "Light", "System"):
                return mode
    except Exception:
        log_exception("Failed to load theme_config.json")
    return "Dark"


def save_theme_preference(mode):
    try:
        with open(THEME_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"mode": mode}, f, indent=2)
        write_audit_log("THEME-CHANGED", f"mode={mode}")
    except Exception:
        log_exception("Failed to save theme_config.json")


def _resolve_system_theme():
    """Best-effort OS dark/light detection. Falls back to 'dark' (this
    app's original look) when detection isn't available -- there's no
    universal stdlib way to ask the OS, and adding a hard dependency on a
    platform-detection package just for this one preference isn't worth
    it, so 'System' degrades gracefully rather than guessing wrong."""
    try:
        import darkdetect
        return "dark" if darkdetect.isDark() else "light"
    except Exception:
        return "dark"






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
    Safe to call even if some tags are absent (best-effort)."""
    pseudonym = _pseudonym_for(original_pid)
    try:
        ds.PatientID = pseudonym
        ds.PatientName = pseudonym
    except Exception:
        pass

    for tag in ANONYMIZE_TAGS_BLANK:
        if hasattr(ds, tag):
            try:
                setattr(ds, tag, "")
            except Exception:
                pass

    for tag in ANONYMIZE_TAGS_REMOVE_IF_PRESENT:
        if hasattr(ds, tag):
            try:
                delattr(ds, tag)
            except Exception:
                pass

    # Strip private tags and curves/overlays which can carry burned-in PHI
    try:
        ds.remove_private_tags()
    except Exception:
        pass

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
            with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
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
            for row in reader:
                pid = row["Patient ID"]
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
        with open(LOG_RETENTION_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"retention_days": days}, f, indent=2)
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
                if os.path.getsize(log_path) >= LOG_ROTATION_MAX_BYTES:
                    needs_rotation = True
                elif _log_file_age_days(log_path) >= retention_days:
                    needs_rotation = True
            if needs_rotation:
                reason = "size-threshold" if os.path.exists(log_path) and os.path.getsize(log_path) >= LOG_ROTATION_MAX_BYTES else "retention-window"
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

def upsert_patient(pid, pname, institution, study_uid, modality, source, status=None):
    """Create or update a worklist entry. Thread-safe.
    `source` is either the calling AE Title (network receive), the literal
    string "Imported" (folder import), or any caller-supplied label."""
    with data_lock:
        if pid in patient_data:
            patient_data[pid]["count"] += 1
            patient_data[pid]["time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            patient_data[pid]["source"] = source
            # New images arriving for an already-Sent study should NOT silently
            # stay marked Sent -- that hides the fact there's new unsent data.
            if patient_data[pid].get("status") == STATUS_SENT:
                patient_data[pid]["status"] = STATUS_PENDING
        else:
            patient_data[pid] = {
                "patient_name": pname,
                "institution": institution,
                "study_uid": study_uid,
                "modality": modality,
                "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "count": 1,
                "doc_count": 0,
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
    folder = os.path.join(OUTPUT_DIR, pid)
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


# =========================================================
# Shared filesystem helpers (still used by report generation, etc.)
# =========================================================

def get_patient_folder(pid):
    return os.path.join(OUTPUT_DIR, pid)


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
REPORT_FILENAME = "Radiology_Report.docx"

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
    return os.path.join(get_reports_folder(pid), REPORT_FILENAME)


def ensure_reports_folder(pid):
    """Creates the patient + Reports folders if missing. Returns
    (ok, folder_path_or_error_message)."""
    def _make():
        folder = get_reports_folder(pid)
        os.makedirs(folder, exist_ok=True)
        return folder
    return _safe_document_op("Create Reports folder", _make)


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

HISTORY_FILENAME = "Patient_History.txt"


def get_history_path(pid):
    return os.path.join(get_reports_folder(pid), HISTORY_FILENAME)


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
        with open(path, "w", encoding="utf-8") as f:
            f.write(header)
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

            dest_folder = os.path.join(OUTPUT_DIR, pid)
            os.makedirs(dest_folder, exist_ok=True)
            dest_path = os.path.join(dest_folder, f"{sop_uid}.dcm")

            if not os.path.exists(dest_path):
                shutil.copy2(path, dest_path)

            upsert_patient(pid, pname, institution, study_uid, modality, source="Imported")
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

        folder = os.path.join(OUTPUT_DIR, pid)
        os.makedirs(folder, exist_ok=True)
        file_path = os.path.join(folder, f"{sop_uid}.dcm")

        is_duplicate = os.path.exists(file_path)
        ds.save_as(file_path, write_like_original=False)

        try:
            file_size = os.path.getsize(file_path)
        except Exception:
            file_size = None
        duration_sec = round(time.time() - receive_started, 3)

        upsert_patient(pid, pname, institution, study_uid, modality,
                       source=calling_ae, status=STATUS_RECEIVED)
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
            threading.Thread(target=push_single_patient, args=(pid,), daemon=True).start()

        ui_event_queue.put(("toast", ("Image Received", f"{pid} ({modality}) from {calling_ae}")))

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
    """pynetdicom/DICOM allows at most 128 presentation contexts per
    association. If sopclass.ini lists more SOP Classes than that, trim
    and log a warning rather than letting association building fail."""
    if len(sops) > MAX_PRESENTATION_CONTEXTS:
        write_receiver_log(
            "", "", "",
            f"sopclass.ini lists {len(sops)} SOP Classes; only the first "
            f"{MAX_PRESENTATION_CONTEXTS} are usable per DICOM association limits. "
            f"Trim sopclass.ini's [SOP_CLASSES] section if you need specific ones."
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
    pynetdicom ValueError surfacing later at association time."""
    title = (title or "").strip()
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
    """Validate a TCP port string. Returns (ok, int_value_or_error_message)."""
    port_str = (port_str or "").strip()
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

    sops = cap_sop_list(load_extra_sops())
    transfer_syntaxes = load_transfer_syntaxes()

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


def stop_receiver_server():
    ae = receiver_state.get("server_ae")
    if ae:
        try:
            ae.shutdown()
        except Exception:
            log_exception("Error shutting down receiver AE")
    receiver_state["running"] = False
    ui_event_queue.put(("receiver_status", False))

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


def dicom_echo(remote_ae, remote_ip, remote_port, timeout=5, calling_ae=None):
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
        ssl_context = build_ssl_context_for_client(tls_cfg, port=port_val)

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


def _send_one_file(assoc, fpath, anonymize, pid):
    """Send a single file over an existing association. Optionally
    anonymizes the dataset in memory before sending (does NOT modify the
    file on disk). Returns (ok: bool, error_message: str)."""
    try:
        try:
            bandwidth_limiter.throttle(os.path.getsize(fpath))
        except OSError:
            pass
        ds = pydicom.dcmread(fpath, force=True)
        if anonymize:
            ds = anonymize_dataset(ds, pid)
        status = assoc.send_c_store(ds)
        if status and getattr(status, "Status", 1) == 0x0000:
            return True, ""
        return False, f"C-STORE non-success status for {fpath}: {status}"
    except Exception as e:
        return False, f"{fpath}: {e}"


def push_single_patient(pid, on_progress=None, destination=None, anonymize=False):
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

    folder = os.path.join(OUTPUT_DIR, pid)
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
    ssl_context = build_ssl_context_for_client(tls_cfg, port=port_val)

    ae = AE(ae_title=calling_clean)
    ae.acse_timeout = PUSH_ASSOC_TIMEOUT_SEC
    ae.dimse_timeout = PUSH_ASSOC_TIMEOUT_SEC
    ae.network_timeout = PUSH_ASSOC_TIMEOUT_SEC

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
    for fpath in files:
        try:
            ds = pydicom.dcmread(fpath, stop_before_pixels=True, force=True)
            sop_uid = str(getattr(ds, "SOPClassUID", "")).strip()
            if sop_uid:
                file_sop_classes.add(sop_uid)
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

    # Always include the three baseline transfer syntaxes so servers that
    # only support Implicit VR or Explicit VR Big/Little are covered.
    BASELINE_TS = [
        UID("1.2.840.10008.1.2"),    # Implicit VR Little Endian
        UID("1.2.840.10008.1.2.1"),  # Explicit VR Little Endian
        UID("1.2.840.10008.1.2.2"),  # Explicit VR Big Endian (retired but still seen)
    ]
    configured_ts = load_transfer_syntaxes()
    combined_ts = list(dict.fromkeys(configured_ts + BASELINE_TS))  # preserve order, deduplicate

    for sop in all_sops:
        try:
            ae.add_requested_context(sop, combined_ts)
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
        remaining_files = [f for f in files
                           if os.path.splitext(os.path.basename(f))[0] not in already_sent_uids]
        sent = total - len(remaining_files)  # credit for what a prior attempt already delivered
        if sent:
            with data_lock:
                push_job["attempted_images"] += sent
                push_job["sent_images"] += sent
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
        set_fields(
            pid, status=STATUS_SENT,
            sent_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            push_target=f"{remote_ae_clean}@{cfg['ip']}:{port_val}",
            last_error="", retry_count=0,
        )
        clear_transfer_checkpoint(pid)
        write_audit_log("PUSH-OK", f"pid={pid} sent={sent}/{total} dest={remote_ae_clean}@{cfg['ip']}:{port_val} (resumed, nothing left to send)")
        record_push_stat(next(iter(file_study_uids), None), sent, 0)
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
                    record_checkpoint_progress(pid, dest_name_for_checkpoint, sop_uid_sent)
                else:
                    last_error = err
                    write_push_log(pid, "", err)

                with data_lock:
                    push_job["attempted_images"] += 1
                    if ok:
                        push_job["sent_images"] += 1
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
        return sent, total, False

    delay = RETRY_BASE_DELAY_SEC * (2 ** retry_count)
    record_checkpoint_interruption(pid, dest_name_for_checkpoint, last_error or f"{failed_count} image(s) failed")
    set_fields(pid, status=STATUS_RETRYING, last_error=last_error, retry_count=retry_count + 1)
    write_audit_log("PUSH-RETRY-SCHEDULED", f"pid={pid} attempt={retry_count + 1} delay={delay}s")

    for _ in range(int(delay * 10)):
        if push_job["stop_flag"] or app_shutdown_event.is_set():
            set_fields(pid, status=STATUS_FAILED, last_error="Stopped before retry")
            return sent, total, False
        time.sleep(0.1)

    return push_single_patient(pid, on_progress=on_progress, destination=destination, anonymize=anonymize)


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

    threads = worker_threads or DEFAULT_PUSH_WORKER_THREADS
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
    app.deiconify()


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
    drops it. Never raises — toasts are best-effort."""
    try:
        from plyer import notification
        notification.notify(title=title, message=message, app_name="R-Apps DICOM", timeout=4)
        return
    except Exception:
        pass
    # Fallback: tiny floating label in the corner of the app window
    try:
        import tkinter as tk
        toast = tk.Toplevel()
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        sw, sh = toast.winfo_screenwidth(), toast.winfo_screenheight()
        toast.geometry(f"340x55+{sw - 360}+16")
        toast.configure(bg="#181b21")
        tk.Label(
            toast, text=f"{title}\n{message}", bg="#181b21", fg="white",
            font=("Helvetica", 10), wraplength=320, justify="left",
            padx=8, pady=6,
        ).pack(fill="both", expand=True)
        toast.after(4000, toast.destroy)
    except Exception:
        pass


def show_toast_threadsafe(title, message):
    """Queue a toast that will be shown on the main thread via the event
    pump. Using this from background threads avoids Tk thread-safety issues."""
    ui_event_queue.put(("toast", (title, message)))

# =========================================================
# MAIN GUI
# =========================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# THEME PALETTE -- single source of truth for the GUI polish pass.
# Cool slate/navy palette (clinical-workstation tone) rather than a
# high-contrast "developer dark mode" black.
THEME_BG = "#12161c"
THEME_SURFACE = "#1a212b"
THEME_HEADING_BG = "#232c38"
THEME_ACCENT = "#3d7fb8"
THEME_ACCENT_HOVER = "#336a99"
THEME_TEXT = "#e4e8ee"
THEME_TEXT_MUTED = "#8b97a8"
THEME_NEUTRAL_BTN = "#333d4a"
THEME_NEUTRAL_BTN_HOVER = "#414e5e"
THEME_DANGER = "#c94f4f"
THEME_DANGER_HOVER = "#a83f3f"
THEME_SUCCESS = "#4d9e77"
THEME_SUCCESS_HOVER = "#3f8563"
FONT_FAMILY = "Segoe UI" if platform.system() == "Windows" else "Helvetica"

APP_VERSION = "3.0.0"

# =========================================================
# DESIGN SYSTEM -- shared primitives (Phase 1)
# =========================================================
# Single source of truth for typography scale and reusable widget
# builders (cards, status badges) so every page/tab styles itself the
# same way. These read the live THEME_* globals at call time (not at
# import time), so anything built with them stays correct across
# apply_theme() switches -- same pattern the treeview styling already
# uses. Nothing here touches DICOM/business logic; it only standardizes
# how existing pages present it.

FONT_SCALE = {
    "title": 17,      # page/app title
    "subtitle": 12,    # secondary header text
    "section": 14,     # card/section headings
    "body": 12,        # normal text
    "small": 11,       # dense table / caption text
    "micro": 10,       # timestamps, footnotes
}


def get_font(scale="body", weight="normal"):
    """Returns a CTkFont at one of the standard scale steps above."""
    size = FONT_SCALE.get(scale, FONT_SCALE["body"])
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
        "warning": ("#e6a23c", "#e6a23c"),
        "retrying": ("#e6a23c", "#e6a23c"),
        "pending": (THEME_TEXT_MUTED, THEME_TEXT_MUTED),
        "sending": (THEME_ACCENT, THEME_ACCENT),
        "receiving": (THEME_ACCENT, THEME_ACCENT),
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
    frame = ctk.CTkFrame(parent, fg_color=THEME_HEADING_BG, corner_radius=999, height=24)
    frame.pack_propagate(False)
    dot = ctk.CTkLabel(frame, text="●", font=ctk.CTkFont(size=11), text_color=dot_color, width=12)
    dot.pack(side="left", padx=(10, 2), pady=2)
    lbl = ctk.CTkLabel(frame, text=text, font=get_font("small", "bold"), text_color=color)
    lbl.pack(side="left", padx=(0, 12), pady=2)

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
    _SPEED_PX = 1           # pixels moved per animation tick
    _TICK_MS = 30           # animation tick interval
    _PAUSE_MS = 1200        # pause at the start of each loop
    _GAP_PX = 40            # gap between the end of the text and its repeat

    def __init__(self, parent, text="", width=140, height=20, font=None,
                 text_color=None, anchor="w", fg_color="transparent", canvas_bg=None, **kwargs):
        super().__init__(parent, width=width, height=height, fg_color=fg_color, **kwargs)
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
            state["after_id"] = btn.after(160, tick)

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


def make_card(parent, title=None, subtitle=None, fg_color=None):
    """Standard enterprise 'card' container: rounded surface, optional
    title/subtitle header, and a `.body` frame for page content to pack
    into. Used as the common building block for Dashboard tiles, Receiver/
    Pusher panels, Reports, Settings groups, etc.

    Filled with a solid tone distinct from the page background rather than
    relying on a thin border for definition -- a 1px outline in a color
    close to the fill reads as a faint, incomplete-looking line rather
    than a solid box."""
    card = ctk.CTkFrame(parent, fg_color=fg_color or THEME_HEADING_BG, corner_radius=12)
    if title:
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 0 if subtitle else 8))
        ctk.CTkLabel(header, text=title, font=get_font("section", "bold"),
                     text_color=THEME_TEXT).pack(side="left")
        if subtitle:
            ctk.CTkLabel(card, text=subtitle, font=get_font("micro"),
                         text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=16, pady=(2, 8))
    body = ctk.CTkFrame(card, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
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
    row.pack(fill="x", padx=6, pady=pady)
    row._card_col = 0
    return row


def add_card_to_row(row, card, padx=5, pady=5):
    """Places `card` as the next equal-width column in `row` (see make_card_row)."""
    col = row._card_col
    row._card_col += 1
    row.grid_columnconfigure(col, weight=1, uniform="card_row")
    card.grid(row=0, column=col, padx=padx, pady=pady, sticky="nsew")



    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        pass  # Older Windows without shcore, or running under Wine -- degrade gracefully

app = ctk.CTk()
app.title("R-Apps DICOM")
app.geometry("1400x860")
app.minsize(1150, 720)
app.configure(fg_color=THEME_BG)
try:
    _logo_for_icon = _load_logo_pil()
    if _logo_for_icon is not None:
        app.iconphoto(True, ImageTk.PhotoImage(_logo_for_icon))
except Exception:
    pass
app.withdraw()  # stay hidden until the one-time Admin PIN setup (if needed) completes


def on_window_close():
    if modern_askyesno("Exit", "Exit the application? The receiver will stop."):
        graceful_shutdown()


app.protocol("WM_DELETE_WINDOW", on_window_close)


def _keep_toplevel_small(win):
    """Forces a floating Toplevel/CTkToplevel panel or dialog to open as
    a small windowed popup, never full-screen/maximized.

    On several platforms (most notably Windows), a Toplevel created while
    its master is maximized or in a fullscreen/"zoomed" state can itself
    inherit that state, which made small panels like the Notification
    Center balloon out to fill the whole screen. Explicitly forcing
    'normal' state -- both immediately and again shortly after the window
    manager finishes mapping the window -- keeps these panels at their
    intended size regardless of whether the main app window is
    maximized, fullscreen, or windowed."""
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
    font=ctk.CTkFont(family=FONT_FAMILY, size=17, weight="bold"),
    text_color=THEME_TEXT,
)
header_title_lbl.pack(side="left", padx=(8 if _header_logo_img is not None else 18, 8))

header_subtitle_lbl = ctk.CTkLabel(
    header_bar, text="Receiver & AutoRouter",
    font=ctk.CTkFont(family=FONT_FAMILY, size=11),
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


def _pacs_chip_tooltip_text():
    if not destination_health_cache:
        return "No push destinations configured yet."
    lines = []
    for name, rec in destination_health_cache.items():
        state = "online" if rec.get("ok") else "unreachable"
        lines.append(f"{name}: {state}")
    return "\n".join(lines) or "No destinations checked yet."


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
            any_online = any(rec.get("ok") for rec in destination_health_cache.values())
            all_online = all(rec.get("ok") for rec in destination_health_cache.values())
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
        "warning": "#e6a23c", "info": THEME_ACCENT,
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
        dot = ctk.CTkLabel(row, text="●", font=ctk.CTkFont(size=13),
                            text_color=_notification_kind_color(rec["kind"]), width=18)
        dot.pack(side="left", padx=(8, 0), pady=8)
        text_col = ctk.CTkFrame(row, fg_color="transparent")
        text_col.pack(side="left", fill="x", expand=True, padx=(4, 8), pady=6)
        ctk.CTkLabel(text_col, text=rec["title"], font=get_font("small", "bold"),
                     text_color=THEME_TEXT, anchor="w", justify="left").pack(fill="x")
        ctk.CTkLabel(text_col, text=rec["message"], font=get_font("micro"),
                     text_color=THEME_TEXT_MUTED, anchor="w", justify="left",
                     wraplength=340).pack(fill="x")
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
    try:
        win.grab_set()
    except Exception:
        pass

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

admin_bar = ctk.CTkFrame(app, fg_color=THEME_SURFACE, corner_radius=12)
admin_bar.pack(fill="x", padx=14, pady=(0, 8))

admin_bar_status_badge = make_status_badge(admin_bar, "Mode: USER", kind="neutral")
admin_bar_status_badge.pack(side="left", padx=(14, 15), pady=8)

# Back-compat shim: existing code elsewhere configures a text_color on
# `admin_bar_status_lbl` directly (see apply_theme / role-switch logic).
# Point that name at the badge's inner label so those call sites keep
# working unmodified.
admin_bar_status_dot = admin_bar_status_badge.winfo_children()[0]
admin_bar_status_lbl = admin_bar_status_badge.winfo_children()[1]

admin_login_btn = ctk.CTkButton(admin_bar, text="Admin Login", width=150, height=32,
                                 image=get_icon("lock", size=15, color="#ffffff"), compound="left",
                                 corner_radius=8, fg_color=THEME_ACCENT, hover_color=THEME_ACCENT_HOVER,
                                 font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"))
admin_login_btn.pack(side="left", padx=4, pady=6)

admin_switch_view_btn = ctk.CTkButton(admin_bar, text="Switch to User View", width=180, height=32,
                                       image=get_icon("user", size=15, color=THEME_TEXT), compound="left",
                                       corner_radius=8,
                                       fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                                       font=ctk.CTkFont(family=FONT_FAMILY, size=12))

admin_change_pin_btn = ctk.CTkButton(admin_bar, text="Change Admin PIN", width=170, height=32,
                                      image=get_icon("key-round", size=15, color=THEME_TEXT), compound="left",
                                      corner_radius=8,
                                      fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER,
                                      font=ctk.CTkFont(family=FONT_FAMILY, size=12))

admin_lock_btn = ctk.CTkButton(admin_bar, text="Lock", width=110, height=32,
                                image=get_icon("lock", size=15, color="#ffffff"), compound="left",
                                corner_radius=8,
                                fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER,
                                font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"))



# =========================================================
# SHARED WORKLIST TREEVIEW BUILDER
# =========================================================

WL_COLUMNS = (
    "patient_id", "patient_name", "institution", "modality",
    "count", "documents", "report", "history", "source", "status", "time", "sent_time", "push_target", "last_error",
)

WL_HEADINGS = {
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
    "patient_id": 170, "patient_name": 160, "institution": 190,
    "modality": 80, "count": 60, "documents": 150, "report": 150, "history": 150, "source": 150, "status": 110,
    "time": 175, "sent_time": 175, "push_target": 210, "last_error": 260,
}


def build_worklist_tree(parent):
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background=THEME_SURFACE, foreground=THEME_TEXT,
                    fieldbackground=THEME_SURFACE, rowheight=28, borderwidth=0,
                    font=(FONT_FAMILY, 11))
    style.configure("Treeview.Heading", background=THEME_HEADING_BG, foreground=THEME_TEXT,
                    font=(FONT_FAMILY, 11, "bold"), relief="flat", borderwidth=0)
    style.map("Treeview.Heading", background=[("active", "#2c3140")])
    style.map("Treeview", background=[("selected", THEME_ACCENT)],
              foreground=[("selected", "#ffffff")])
    style.configure("Stale.Treeview", foreground=STALE_HIGHLIGHT_COLOR)
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # drop default border frame

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
        selectmode="extended",
    )
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    for col in WL_COLUMNS:
        tree.heading(col, text=WL_HEADINGS[col],
                     command=lambda c=col: sort_tree(tree, c, False))
        tree.column(col, width=WL_WIDTHS[col], minwidth=40, anchor="w")

    tree.tag_configure(STALE_HIGHLIGHT_TAG, foreground=STALE_HIGHLIGHT_COLOR)
    tree.tag_configure(SEARCH_MATCH_HIGHLIGHT_TAG, background=SEARCH_MATCH_HIGHLIGHT_BG)
    tree.tag_configure("even_row", background=THEME_SURFACE)
    tree.tag_configure("odd_row", background="#1d212a")
    # Status-column color coding (STATUS_COLORS existed but was never
    # wired to the tree before -- this is the actual "plain text status ->
    # visual indicator" fix for the app's main table).
    for _status_val, _status_color in STATUS_COLORS.items():
        tree.tag_configure(f"status_{_status_val}", foreground=_status_color)

    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)
    return frame, tree


def sort_tree(tree, col, descending):
    data = [(tree.set(k, col), k) for k in tree.get_children("")]
    data.sort(reverse=descending, key=lambda x: x[0].lower() if isinstance(x[0], str) else x[0])
    for index, (_val, k) in enumerate(data):
        tree.move(k, "", index)
    tree.heading(col, command=lambda: sort_tree(tree, col, not descending))


_tree_render_cache = {}  # id(tree) -> (data_version, search) last rendered


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

    with data_lock:
        rows = list(patient_data.items())

    visible_row_index = 0
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

        zebra_tag = "even_row" if visible_row_index % 2 == 0 else "odd_row"
        visible_row_index += 1
        tags = []
        if stale:
            tags.append(STALE_HIGHLIGHT_TAG)
        elif search_lower and search_matched:
            # Only highlight for the free-text search (not the dropdown
            # filters, which already remove non-matches entirely) --
            # and don't fight with the stale-row highlight color.
            tags.append(SEARCH_MATCH_HIGHLIGHT_TAG)
        elif status in STATUS_COLORS:
            tags.append(f"status_{status}")
        tags.append(zebra_tag)
        tags = tuple(tags)

        report_cell = "Open Report" if os.path.isfile(get_report_path(pid)) else "Create Report"
        history_cell = "Open History" if os.path.isfile(get_history_path(pid)) else "Create History"

        tree.insert(
            "", "end", iid=pid,
            values=(
                pid,
                d.get("patient_name", ""),
                d.get("institution", ""),
                d.get("modality", ""),
                d.get("count", 0),
                "Open in Viewer" if d.get("count", 0) else "No Images",
                report_cell,
                history_cell,
                d.get("source", ""),
                status,
                d.get("time", ""),
                d.get("sent_time", ""),
                d.get("push_target", ""),
                d.get("last_error", ""),
            ),
            tags=tags,
        )

    # Restore selections that still exist
    for pid in selected_before:
        if tree.exists(pid):
            tree.selection_add(pid)

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
        text_color=THEME_TEXT_MUTED, corner_radius=8, height=36,
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

    shell = ctk.CTkFrame(win, fg_color=THEME_SURFACE, corner_radius=14,
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
    segmented_button_unselected_hover_color="#2c3140",
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

dash_home_scroll = ctk.CTkScrollableFrame(tab_dashboard_home, fg_color="transparent")
dash_home_scroll.pack(fill="both", expand=True, padx=4, pady=4)


def _home_section_label(parent, text):
    ctk.CTkLabel(parent, text=text, font=get_font("section", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=10, pady=(18, 6))


def _home_cards_row(parent):
    return make_card_row(parent)


def _home_stat_card(parent, title, value="—"):
    card = make_card(parent)
    add_card_to_row(parent, card)
    ctk.CTkLabel(card.body, text=title, font=get_font("micro"),
                 text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 0))
    # width=1 + fill="x": the card's own layout (grid, uniform column) decides
    # the real width; this just tracks it and scrolls only if the value
    # doesn't fit in whatever room the card ends up with.
    val_lbl = MarqueeLabel(card.body, text=value, width=1, height=26,
                            font=get_font("title", "bold"), text_color=THEME_TEXT,
                            canvas_bg=THEME_HEADING_BG)
    val_lbl.pack(anchor="w", fill="x", pady=(2, 0))
    return val_lbl


# ---- Receiver stats ----
_home_section_label(dash_home_scroll, "Receiver")
home_recv_row1 = _home_cards_row(dash_home_scroll)
dash_home_recv_status = _home_stat_card(home_recv_row1, "Receiver Status", "● Stopped")
dash_home_recv_ae = _home_stat_card(home_recv_row1, "AE Title")
dash_home_recv_port = _home_stat_card(home_recv_row1, "Listening Port")
dash_home_recv_ip = _home_stat_card(home_recv_row1, "Current IP")
dash_home_recv_assoc = _home_stat_card(home_recv_row1, "Active Associations", "0")

home_recv_row2 = _home_cards_row(dash_home_scroll)
dash_home_studies_recv = _home_stat_card(home_recv_row2, "Studies Received Today", "0")
dash_home_images_recv = _home_stat_card(home_recv_row2, "Images Received Today", "0")
dash_home_docs_recv = _home_stat_card(home_recv_row2, "Documents Received Today", "0")
dash_home_total_patients = _home_stat_card(home_recv_row2, "Total Patients", "0")
dash_home_recv_queue = _home_stat_card(home_recv_row2, "Current Queue Size", "0")

# ---- Pusher stats ----
_home_section_label(dash_home_scroll, "Pusher")
home_push_row1 = _home_cards_row(dash_home_scroll)
dash_home_push_status = _home_stat_card(home_push_row1, "Push Status", "Idle")
dash_home_dest_status = _home_stat_card(home_push_row1, "Destination Status", "Unknown")
dash_home_images_sent = _home_stat_card(home_push_row1, "Images Sent Today", "0")
dash_home_studies_sent = _home_stat_card(home_push_row1, "Studies Sent Today", "0")
dash_home_failed_transfers = _home_stat_card(home_push_row1, "Failed Transfers", "0")

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
        image=get_icon("triangle-alert", size=14, color="#e67e22"), compound="left",
        font=ctk.CTkFont(size=11), text_color="#e67e22",
    ).pack(anchor="w", padx=16, pady=(0, 4))

# ---- Live graphs ----
_home_section_label(dash_home_scroll, "Live Graphs (auto-refreshing)")
home_graphs_row = ctk.CTkScrollableFrame(dash_home_scroll, fg_color="transparent",
                                          orientation="horizontal", height=200)
home_graphs_row.pack(fill="x", padx=6, pady=(0, 12))

DASH_GRAPH_COLORS = {
    "studies_received": "#2f8eff",
    "studies_sent": "#2ecc71",
    "failed_transfers": "#f04747",
    "queue_size": "#f1c40f",
    "network_kbps": "#9b59b6",
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


def _draw_sparkline(canvas, values, color):
    """Lightweight dependency-free line chart drawn straight onto a
    tkinter Canvas -- avoids pulling in matplotlib just for small
    auto-refreshing trend lines."""
    try:
        canvas.delete("all")
        w = int(canvas.cget("width"))
        h = int(canvas.cget("height"))
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
tabview.set("Dashboard")
_sync_nav_highlight()

# ---- Pusher monitoring strip ----
# Purely a presentation layer over state that already exists: push_job
# (running/sent/attempted/total), get_offline_queue_summary() (retry
# queue), destination_health_cache (destination status, populated by the
# existing PACS Health / Dashboard background checker), and daily_stats
# (today's counters). No new tracking, no new background thread.

push_monitor_row = make_card_row(tab_pusher, pady=(10, 0))
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


def refresh_pusher_monitoring():
    if push_job.get("running"):
        push_monitor_status_badge.update_status("Sending", "sending")
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
offlineq_outer = ctk.CTkFrame(tab_pusher)
offlineq_outer.pack(fill="x", padx=10, pady=(10, 0))

offlineq_top = ctk.CTkFrame(offlineq_outer, fg_color="transparent")
offlineq_top.pack(fill="x", padx=8, pady=(8, 4))
ctk.CTkLabel(offlineq_top, text="Offline Queue", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
offlineq_retry_all_btn = ctk.CTkButton(offlineq_top, text="Retry Due Items Now", width=180)
offlineq_retry_all_btn.pack(side="right", padx=4)

offlineq_stats_row = make_card_row(offlineq_outer)
offlineq_stats_row.pack_configure(padx=8, pady=(0, 6))


def _offlineq_stat(parent, title):
    card = ctk.CTkFrame(parent, corner_radius=12, fg_color=THEME_HEADING_BG)
    add_card_to_row(parent, card, padx=4, pady=0)
    ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).pack(pady=(6, 0))
    val_lbl = ctk.CTkLabel(card, text="0", font=ctk.CTkFont(size=16, weight="bold"))
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
offlineq_tree_frame = ctk.CTkFrame(offlineq_outer, fg_color="#181b21")
offlineq_tree_frame.pack(fill="x", padx=8, pady=(0, 8))
offlineq_tree = ttk.Treeview(offlineq_tree_frame, columns=offlineq_cols, show="headings", height=5)
for col in offlineq_cols:
    offlineq_tree.heading(col, text=offlineq_headings[col])
    offlineq_tree.column(col, width=offlineq_widths[col], anchor="w")
offlineq_tree.pack(fill="x")
ctk.CTkLabel(offlineq_outer, text="(Double-click a row to retry that item immediately)",
             font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=10, pady=(0, 6))


ADMIN_ONLY_TAB_NAMES = [
    "Admin Dashboard", "PACS Health", "Bandwidth", "Export", "Reports",
    "Performance", "Backup", "LDAP / AD", "Query/Retrieve", "Destinations",
    "Routing Rules", "SOP Classes", "Logs",
]

# =========================================================
# RECEIVER TAB
# =========================================================

rec_monitor_row = make_card_row(tab_receiver, pady=(10, 0))
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


def refresh_receiver_monitoring():
    """Keeps the Receiver tab's own monitoring strip in sync. Reuses the
    exact same state the Dashboard already reads (receiver_state,
    daily_stats, patient_data) -- no new tracking added."""
    running = receiver_state.get("running")
    rec_monitor_status_badge.update_status("Running" if running else "Stopped",
                                            "running" if running else "stopped")
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


rec_top = ctk.CTkFrame(tab_receiver, fg_color="transparent")
rec_top.pack(fill="x", padx=10, pady=10)

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

autoroute_var = ctk.BooleanVar(value=True)
autoroute_checkbox = ctk.CTkCheckBox(
    cfg_frame, text="Enable AutoRoute (push on arrival)",
    variable=autoroute_var,
    command=lambda: receiver_state.__setitem__("autoroute", autoroute_var.get()))
autoroute_checkbox.grid(row=2, column=0, columnspan=2, padx=8, pady=8, sticky="w")

rec_status_label = ctk.CTkLabel(cfg_frame, text="● Stopped", text_color=THEME_DANGER,
                                 font=ctk.CTkFont(size=12, weight="bold"))
rec_status_label.grid(row=3, column=0, columnspan=2, padx=8, pady=4, sticky="w")

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
                          fg_color="#232733", hover_color=THEME_NEUTRAL_BTN_HOVER)
tray_btn.pack(pady=4)

rec_open_viewer_btn = ctk.CTkButton(
    btn_col, text="Open in Viewer", width=200,
    fg_color="#2f8eff", hover_color="#1f6fd6",
    command=lambda: _open_selected_in_viewer(rec_tree))
rec_open_viewer_btn.pack(pady=4)

import_col = ctk.CTkFrame(rec_top, fg_color="transparent")
import_col.pack(side="left", padx=15)

ctk.CTkLabel(import_col, text="Import DICOM Folder:",
             font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
import_btn = ctk.CTkButton(import_col, text="Import Folder...", width=200)
import_btn.pack(pady=6)
import_progress = ctk.CTkProgressBar(import_col, width=200)
import_progress.set(0)
import_progress.pack(pady=4)
import_progress_label = ctk.CTkLabel(import_col, text="No import in progress",
                                      font=ctk.CTkFont(size=11))
import_progress_label.pack()

rec_search_frame = ctk.CTkFrame(tab_receiver, fg_color="transparent")
rec_search_frame.pack(fill="x", padx=10, pady=(0, 5))
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
rec_select_failed_btn = ctk.CTkButton(rec_search_frame, text="Select all Failed", width=155,
                                       fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
rec_select_failed_btn.pack(side="left", padx=4)
rec_select_pending_btn = ctk.CTkButton(rec_search_frame, text="Select all Pending", width=165,
                                        fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
rec_select_pending_btn.pack(side="left", padx=4)
rec_export_btn = ctk.CTkButton(rec_search_frame, text="Export View to CSV", width=170,
                                fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
rec_export_btn.pack(side="left", padx=4)

rec_view_options_btn = ctk.CTkButton(rec_search_frame, text="View", width=90,
                                      image=get_icon("settings-2", size=14, color=THEME_TEXT), compound="left",
                                      fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
rec_view_options_btn.pack(side="right", padx=4)

rec_badge_label = ctk.CTkLabel(tab_receiver, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
rec_badge_label.pack(anchor="w", padx=12, pady=(0, 4))

rec_wl_frame, rec_tree = build_worklist_tree(tab_receiver)
rec_wl_frame.pack(fill="both", expand=True, padx=10, pady=(6, 10))
rec_tree_ref["tree"] = rec_tree

# =========================================================
# PUSHER TAB
# =========================================================

push_top = ctk.CTkFrame(tab_pusher, fg_color="transparent")
push_top.pack(fill="x", padx=10, pady=10)

push_dest_frame = ctk.CTkFrame(push_top)
push_dest_frame.pack(side="left", padx=(0, 15))

ctk.CTkLabel(push_dest_frame, text="Active Destination:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
push_dest_var = ctk.StringVar(value="(none configured)")
push_dest_menu = ctk.CTkOptionMenu(push_dest_frame, variable=push_dest_var, values=["(none configured)"], width=200)
push_dest_menu.grid(row=0, column=1, padx=8, pady=6)

anon_var = ctk.BooleanVar(value=False)
ctk.CTkCheckBox(push_dest_frame, text="Anonymize before push",
                variable=anon_var).grid(row=1, column=0, columnspan=2, padx=8, pady=4, sticky="w")

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
    fg_color="#2f8eff", hover_color="#1f6fd6",
    command=lambda: _open_selected_in_viewer(push_tree))
push_open_viewer_btn.pack(pady=4)
echo_btn = ctk.CTkButton(push_btn_col, text="C-ECHO Active Dest.", width=200)
echo_btn.pack(pady=4)

progress_col = ctk.CTkFrame(push_top, fg_color="transparent")
progress_col.pack(side="left", padx=20, fill="x", expand=True)

ctk.CTkLabel(progress_col, text="Overall Push Progress",
             font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
overall_progress = ctk.CTkProgressBar(progress_col, width=400)
overall_progress.set(0)
overall_progress.pack(pady=4, fill="x")
overall_progress_label = ctk.CTkLabel(progress_col, text="0 / 0 images sent",
                                       font=ctk.CTkFont(size=12))
overall_progress_label.pack(anchor="w")
throughput_label = ctk.CTkLabel(progress_col, text="",
                                 font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
throughput_label.pack(anchor="w")

push_search_frame = ctk.CTkFrame(tab_pusher, fg_color="transparent")
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
push_select_failed_btn = ctk.CTkButton(push_search_frame, text="Select all Failed", width=155,
                                        fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
push_select_failed_btn.pack(side="left", padx=4)
push_select_pending_btn = ctk.CTkButton(push_search_frame, text="Select all Pending", width=165,
                                         fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
push_select_pending_btn.pack(side="left", padx=4)
push_export_btn = ctk.CTkButton(push_search_frame, text="Export View to CSV", width=170,
                                 fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
push_export_btn.pack(side="left", padx=4)
ctk.CTkLabel(push_search_frame, text="(Right-click a row for more actions)",
             font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(side="left", padx=10)

push_view_options_btn = ctk.CTkButton(push_search_frame, text="View", width=90,
                                       image=get_icon("settings-2", size=14, color=THEME_TEXT), compound="left",
                                       fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
push_view_options_btn.pack(side="right", padx=4)

push_badge_label = ctk.CTkLabel(tab_pusher, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
push_badge_label.pack(anchor="w", padx=12, pady=(0, 4))

push_wl_frame, push_tree = build_worklist_tree(tab_pusher)
push_wl_frame.pack(fill="both", expand=True, padx=10, pady=(6, 10))
push_tree_ref["tree"] = push_tree

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
    global dest_echo_btn, dest_status_lbl
    global routing_outer, routing_list_frame, routing_cols, routing_headings, routing_widths
    global routing_vsb, routing_tree, routing_form_frame, routing_fields
    global routing_dest_var, routing_dest_menu, routing_add_btn, routing_del_btn
    global routing_up_btn, routing_down_btn
    global sop_outer, sop_split, sop_left, sop_classes_box, sop_right, sop_ts_box
    global sop_btn_row, sop_load_btn, sop_save_btn, sop_status_lbl
    global logs_top, log_file_var, log_selector, log_refresh_btn
    global log_archive_now_btn, log_export_btn
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
    global dash_received_lbl, dash_pushed_lbl, dash_failed_lbl
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
    global export_files_frame, export_patient_row
    global export_reports_var, export_logs_var, export_metadata_var, export_dicomdir_var
    global export_progress_bar, export_status_lbl, export_start_btn, export_scope_panels
    global reports_outer, reports_type_var, reports_custom_row
    global reports_from_entry, reports_to_entry, reports_generate_btn
    global reports_print_btn, reports_email_btn, reports_status_lbl
    global reports_preview_btn, reports_meta_card, reports_meta_lbl
    global perf_outer, perf_cards, perf_graph_canvases, perf_export_btn, perf_status_lbl
    global backup_outer, backup_manual_btn, backup_restore_btn, backup_status_lbl, backup_last_badge
    global backup_sched_enabled_var, backup_sched_freq_var, backup_sched_hour_var, backup_sched_save_btn
    global backup_history_tree
    global ldap_outer, ldap_enabled_var, ldap_server_entry, ldap_ssl_var, ldap_domain_entry
    global ldap_bind_dn_entry, ldap_bind_pw_entry, ldap_user_base_entry, ldap_user_filter_entry
    global ldap_bind_template_entry, ldap_group_base_entry, ldap_status_lbl
    global ldap_group_map_frame, ldap_group_map_entries, ldap_save_btn, ldap_import_btn
    global ldap_roster_tree

    # ---- Admin Dashboard tab (first tab in Admin mode) ----
    tab_dashboard = tabview.add("Admin Dashboard")

    dash_outer = ctk.CTkFrame(tab_dashboard, fg_color="transparent")
    dash_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    dash_top_row = make_card_row(dash_outer, pady=(0, 10))

    def _stat_card(parent, title):
        card = ctk.CTkFrame(parent, corner_radius=12, fg_color=THEME_HEADING_BG)
        add_card_to_row(parent, card, padx=6, pady=0)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=12)).pack(pady=(10, 0))
        val_lbl = MarqueeLabel(card, text="0", width=1, height=32,
                                font=ctk.CTkFont(size=26, weight="bold"), text_color=THEME_TEXT,
                                canvas_bg=THEME_HEADING_BG, anchor="center")
        val_lbl.pack(fill="x", padx=10, pady=(0, 10))
        return val_lbl

    dash_received_lbl = _stat_card(dash_top_row, "Received Today")
    dash_pushed_lbl = _stat_card(dash_top_row, "Pushed Successfully Today")
    dash_failed_lbl = _stat_card(dash_top_row, "Failed Today")

    dash_mid_row = ctk.CTkFrame(dash_outer, fg_color="transparent")
    dash_mid_row.pack(fill="x", pady=(0, 10))

    disk_card = make_card(dash_mid_row)
    disk_card.pack(side="left", padx=6, fill="both", expand=True)
    ctk.CTkLabel(disk_card.body, text="Disk Space", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(4, 2))
    dash_disk_bar = ctk.CTkProgressBar(disk_card.body, width=220)
    dash_disk_bar.set(0)
    dash_disk_bar.pack(pady=4, fill="x")
    dash_disk_lbl = ctk.CTkLabel(disk_card.body, text="—", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    dash_disk_lbl.pack(anchor="w")

    stale_card = make_card(dash_mid_row)
    stale_card.pack(side="left", padx=6, fill="both", expand=True)
    ctk.CTkLabel(stale_card.body, text="Stale Pending Studies", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(4, 2))
    dash_stale_lbl = ctk.CTkLabel(stale_card.body, text="0", font=ctk.CTkFont(size=20, weight="bold"), text_color=STALE_HIGHLIGHT_COLOR)
    dash_stale_lbl.pack(anchor="w")
    dash_stale_jump_btn = ctk.CTkButton(stale_card.body, text="Jump to Receiver, filtered", width=230)
    dash_stale_jump_btn.pack(anchor="w", pady=8)

    recv_card = make_card(dash_mid_row)
    recv_card.pack(side="left", padx=6, fill="both", expand=True)
    ctk.CTkLabel(recv_card.body, text="Receiver Status", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(4, 2))
    dash_receiver_status_badge = make_status_badge(recv_card.body, "Stopped", kind="stopped")
    dash_receiver_status_badge.pack(anchor="w")

    dash_bottom_row = ctk.CTkFrame(dash_outer, fg_color="transparent")
    dash_bottom_row.pack(fill="both", expand=True)

    fail_card = make_card(dash_bottom_row)
    fail_card.pack(side="left", fill="both", expand=True, padx=(0, 6))
    ctk.CTkLabel(fail_card.body, text="Recent Failures (click a row for full error)",
                 font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 4))
    dash_fail_cols = ("patient_id", "patient_name", "time", "last_error")
    dash_fail_headings = {"patient_id": "Patient ID", "patient_name": "Patient Name",
                           "time": "Received", "last_error": "Last Error"}
    dash_fail_widths = {"patient_id": 110, "patient_name": 140, "time": 140, "last_error": 260}
    dash_fail_tree_frame = ctk.CTkFrame(fail_card.body, fg_color="#181b21")
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
                 font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 4))
    dash_health_cols = ("destination", "last_echo", "push_ok", "push_failed")
    dash_health_headings = {"destination": "Destination", "last_echo": "Last C-ECHO",
                             "push_ok": "Sent", "push_failed": "Failed"}
    dash_health_widths = {"destination": 150, "last_echo": 160, "push_ok": 70, "push_failed": 70}
    dash_health_tree_frame = ctk.CTkFrame(health_card.body, fg_color="#181b21")
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
                                  calling_ae=dest.get("calling_ae") or DEFAULT_PUSH_CALLING_AE)

            def on_ui():
                try:
                    dash_dest_health_tree.set(row_id, "last_echo", ("OK" if ok else f"Failed: {msg}")[:40])
                except Exception:
                    pass  # row/tab may have been torn down (e.g. mode switch) mid-test
            app.after(0, on_ui)

        threading.Thread(target=run, daemon=True).start()

    dash_dest_health_tree.bind("<Double-1>", _on_dash_dest_echo)
    ctk.CTkLabel(health_card, text="(Double-click a row to run an on-demand C-ECHO test)",
                 font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=10, pady=(0, 8))

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
    ctk.CTkLabel(health_top, text="PACS Health Monitor", font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
    health_run_all_btn = ctk.CTkButton(health_top, text="Check All Now", width=150)
    health_run_all_btn.pack(side="right", padx=4)
    health_status_lbl = ctk.CTkLabel(health_top, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
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
    health_tree_frame = ctk.CTkFrame(health_outer, fg_color="#181b21")
    health_tree_frame.pack(fill="both", expand=True, padx=2, pady=(0, 8))
    health_tree = ttk.Treeview(health_tree_frame, columns=health_cols, show="headings", height=14)
    for col in health_cols:
        health_tree.heading(col, text=health_headings[col],
                             command=lambda c=col: sort_tree(health_tree, c, False))
        health_tree.column(col, width=health_widths[col], anchor="w")
    health_tree.pack(fill="both", expand=True)
    health_tree.tag_configure("status_online", foreground="#2ecc71")
    health_tree.tag_configure("status_offline", foreground="#f04747")
    health_tree.tag_configure("status_degraded", foreground="#f1c40f")

    ctk.CTkLabel(
        health_outer,
        text="Automatically runs a C-ECHO against every configured destination every "
             f"{PACS_HEALTH_CHECK_INTERVAL_SEC}s. Green = online, Yellow = degraded "
             "(1-2 recent failures), Red = offline (3+ consecutive failures).",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, wraplength=900, justify="left",
    ).pack(anchor="w", padx=4, pady=(0, 6))

    health_run_all_btn.configure(command=do_run_all_health_checks_now)

    # ---- Bandwidth Limiter tab ----
    tab_bandwidth = tabview.add("Bandwidth")
    bw_outer = ctk.CTkFrame(tab_bandwidth, fg_color="transparent")
    bw_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(bw_outer, text="Bandwidth Limiter", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
    ctk.CTkLabel(
        bw_outer,
        text="Throttles DICOM Push, Import, and Export so this application never saturates "
             "the network link. Applies as a shared, long-run average rate across all "
             "active transfers -- it never freezes the UI while waiting.",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, wraplength=700, justify="left",
    ).pack(anchor="w", pady=(4, 14))

    bw_card = make_card(bw_outer)
    bw_card.pack(fill="x")
    bw_preset_frame = bw_card.body
    bw_preset_frame.grid_columnconfigure(5, weight=1)

    ctk.CTkLabel(bw_preset_frame, text="Limit:", font=ctk.CTkFont(size=12)).grid(
        row=0, column=0, padx=(0, 8), pady=12, sticky="w")
    bw_preset_var = ctk.StringVar(value=load_bandwidth_config().get("preset", "Unlimited"))
    bw_preset_menu = ctk.CTkOptionMenu(
        bw_preset_frame, variable=bw_preset_var, width=160,
        values=["Unlimited", "1 Mbps", "5 Mbps", "10 Mbps", "25 Mbps", "50 Mbps", "100 Mbps", "Custom"],
    )
    bw_preset_menu.grid(row=0, column=1, padx=8, pady=12, sticky="w")

    ctk.CTkLabel(bw_preset_frame, text="Custom (Mbps):", font=ctk.CTkFont(size=12)).grid(
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

    bw_status_lbl = ctk.CTkLabel(bw_outer, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
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

    ctk.CTkLabel(export_outer, text="Built-in ZIP Export", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
    ctk.CTkLabel(
        export_outer,
        text="Export an entire patient, a single study or series, a hand-picked list of files, "
             "or the whole worklist -- as a plain ZIP or a password-protected (AES-256) ZIP, "
             "optionally bundling the report, a filtered log excerpt, a metadata summary, and/or "
             "a real DICOMDIR.",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, wraplength=760, justify="left",
    ).pack(anchor="w", pady=(4, 14))

    export_scope_var = ctk.StringVar(value="Entire Patient")
    export_scope_row = ctk.CTkFrame(export_outer, fg_color="transparent")
    export_scope_row.pack(fill="x", pady=(0, 10))
    ctk.CTkLabel(export_scope_row, text="Export:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
    export_scope_menu = ctk.CTkSegmentedButton(
        export_scope_row, variable=export_scope_var,
        values=["Entire Patient", "Study", "Series", "Selected Files", "Entire Worklist"],
    )
    export_scope_menu.pack(side="left")

    # ---- Patient / Study / Series pickers (shown/hidden per scope) ----
    export_patient_row = ctk.CTkFrame(export_outer)
    export_patient_row.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(export_patient_row, text="Patient ID:", font=ctk.CTkFont(size=12)).grid(
        row=0, column=0, padx=(12, 8), pady=10, sticky="w")
    export_pid_var = ctk.StringVar(value="")
    export_pid_menu = ctk.CTkOptionMenu(export_patient_row, variable=export_pid_var, width=200, values=["(no patients)"])
    export_pid_menu.grid(row=0, column=1, padx=8, pady=10, sticky="w")

    ctk.CTkLabel(export_patient_row, text="Study:", font=ctk.CTkFont(size=12)).grid(
        row=0, column=2, padx=(20, 8), pady=10, sticky="w")
    export_study_var = ctk.StringVar(value="")
    export_study_menu = ctk.CTkOptionMenu(export_patient_row, variable=export_study_var, width=260, values=["(select a patient)"])
    export_study_menu.grid(row=0, column=3, padx=8, pady=10, sticky="w")

    ctk.CTkLabel(export_patient_row, text="Series:", font=ctk.CTkFont(size=12)).grid(
        row=0, column=4, padx=(20, 8), pady=10, sticky="w")
    export_series_var = ctk.StringVar(value="")
    export_series_menu = ctk.CTkOptionMenu(export_patient_row, variable=export_series_var, width=260, values=["(select a patient)"])
    export_series_menu.grid(row=0, column=5, padx=8, pady=10, sticky="w")

    # ---- Selected Files picker (multi-select tree of this patient's individual instances) ----
    export_files_frame = ctk.CTkFrame(export_outer, fg_color="#181b21")
    export_files_frame.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(export_files_frame, text="Ctrl/Shift-click to select individual files:",
                 font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(anchor="w", padx=8, pady=(6, 2))
    export_files_cols = ("sop_uid", "series_uid", "modality", "size")
    export_files_tree = ttk.Treeview(export_files_frame, columns=export_files_cols, show="headings",
                                      height=6, selectmode="extended")
    for col, label, w in (("sop_uid", "SOP Instance UID", 280), ("series_uid", "Series UID", 260),
                          ("modality", "Modality", 90), ("size", "Size", 90)):
        export_files_tree.heading(col, text=label)
        export_files_tree.column(col, width=w, anchor="w")
    export_files_tree.pack(fill="x", padx=8, pady=(0, 8))

    # ---- Options ----
    export_options_frame = ctk.CTkFrame(export_outer)
    export_options_frame.pack(fill="x", pady=(4, 8))
    ctk.CTkLabel(export_options_frame, text="Options", font=ctk.CTkFont(size=12, weight="bold")).grid(
        row=0, column=0, columnspan=4, padx=12, pady=(10, 2), sticky="w")

    export_reports_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(export_options_frame, text="Include Reports", variable=export_reports_var).grid(
        row=1, column=0, padx=12, pady=6, sticky="w")
    export_logs_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(export_options_frame, text="Include Logs", variable=export_logs_var).grid(
        row=1, column=1, padx=12, pady=6, sticky="w")
    export_metadata_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(export_options_frame, text="Include Metadata", variable=export_metadata_var).grid(
        row=1, column=2, padx=12, pady=6, sticky="w")
    export_dicomdir_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(export_options_frame, text="Include DICOMDIR", variable=export_dicomdir_var).grid(
        row=1, column=3, padx=12, pady=6, sticky="w")

    ctk.CTkLabel(export_options_frame, text="Password (leave blank for a plain ZIP):",
                 font=ctk.CTkFont(size=12)).grid(row=2, column=0, columnspan=2, padx=12, pady=(6, 12), sticky="w")
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
    export_status_lbl = ctk.CTkLabel(export_outer, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    export_status_lbl.pack(anchor="w", pady=(4, 0))

    if not PYZIPPER_AVAILABLE:
        ctk.CTkLabel(
            export_outer,
            text="pyzipper is not installed -- password-protected export will be unavailable "
                 "until you run: pip install pyzipper",
            image=get_icon("triangle-alert", size=14, color="#e67e22"), compound="left",
            font=ctk.CTkFont(size=11), text_color="#e67e22",
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

    # ---- Reports tab ----
    tab_reports = tabview.add("Reports")
    reports_outer = ctk.CTkFrame(tab_reports, fg_color="transparent")
    reports_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(reports_outer, text="PDF Report Generator", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
    ctk.CTkLabel(
        reports_outer,
        text="Studies received/sent, failed studies, success rate, average transfer speed, "
             "top modalities/institutions, destination statistics, queue stats, disk usage, and "
             "uptime -- as a professional PDF with tables and charts.",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, wraplength=760, justify="left",
    ).pack(anchor="w", pady=(4, 14))

    reports_type_var = ctk.StringVar(value="Daily")
    reports_type_row = ctk.CTkFrame(reports_outer, fg_color="transparent")
    reports_type_row.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(reports_type_row, text="Report:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
    reports_type_menu = ctk.CTkSegmentedButton(
        reports_type_row, variable=reports_type_var,
        values=["Daily", "Weekly", "Monthly", "Year to Date", "Custom Date Range"],
    )
    reports_type_menu.pack(side="left")

    reports_custom_row = ctk.CTkFrame(reports_outer)
    ctk.CTkLabel(reports_custom_row, text="From (YYYY-MM-DD):", font=ctk.CTkFont(size=12)).grid(
        row=0, column=0, padx=(12, 6), pady=10, sticky="w")
    reports_from_entry = ctk.CTkEntry(reports_custom_row, width=140,
                                      placeholder_text=datetime.date.today().isoformat())
    reports_from_entry.grid(row=0, column=1, padx=6, pady=10, sticky="w")
    ctk.CTkLabel(reports_custom_row, text="To (YYYY-MM-DD):", font=ctk.CTkFont(size=12)).grid(
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

    reports_status_lbl = ctk.CTkLabel(reports_outer, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    reports_status_lbl.pack(anchor="w", pady=(8, 0))

    reports_meta_card = make_card(reports_outer, title="Last Report")
    reports_meta_card.pack(fill="x", pady=(14, 0))
    reports_meta_lbl = ctk.CTkLabel(reports_meta_card.body, text="No report generated yet this session.",
                                     font=get_font("small"), text_color=THEME_TEXT_MUTED, justify="left")
    reports_meta_lbl.pack(anchor="w")

    if not REPORTLAB_AVAILABLE:
        ctk.CTkLabel(
            reports_outer,
            text="'reportlab' and/or 'matplotlib' are not installed -- PDF report generation is "
                 "unavailable until you run: pip install reportlab matplotlib",
            image=get_icon("triangle-alert", size=14, color="#e67e22"), compound="left",
            font=ctk.CTkFont(size=11), text_color="#e67e22",
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
    ctk.CTkLabel(perf_top_row, text="Performance Metrics", font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
    perf_export_btn = ctk.CTkButton(perf_top_row, text="Export Metrics…", width=160)
    perf_export_btn.pack(side="right")
    perf_status_lbl = ctk.CTkLabel(perf_top_row, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    perf_status_lbl.pack(side="right", padx=10)

    perf_cards = {}

    def _perf_cards_row(parent):
        return make_card_row(parent)

    def _perf_card(parent, key, title):
        card = make_card(parent)
        add_card_to_row(parent, card)
        ctk.CTkLabel(card.body, text=title, font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(4, 0))
        val_lbl = MarqueeLabel(card.body, text="—", width=1, height=26,
                               font=ctk.CTkFont(size=18, weight="bold"), text_color=THEME_TEXT,
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

    if not PSUTIL_AVAILABLE:
        ctk.CTkLabel(
            perf_outer,
            text="psutil is not installed -- CPU/RAM/Disk I/O/Network metrics will show as N/A.",
            image=get_icon("triangle-alert", size=14, color="#e67e22"), compound="left",
            font=ctk.CTkFont(size=11), text_color="#e67e22",
        ).pack(anchor="w", padx=16, pady=(0, 4))

    ctk.CTkLabel(perf_outer, text="Historical Graphs (last 2 minutes, auto-refreshing)",
                 font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(14, 4))
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
        ctk.CTkLabel(card.body, text=title, font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(2, 0))
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

    ctk.CTkLabel(backup_outer, text="Backup & Restore", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
    ctk.CTkLabel(
        backup_outer,
        text="Backs up configuration, encryption keys, routing rules, destination profiles, TLS "
             "settings, reports, the CSV database, logs, and audit logs into one validated ZIP. "
             "Restoring automatically takes a safety backup of the current state first.",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, wraplength=760, justify="left",
    ).pack(anchor="w", pady=(4, 14))

    backup_run_row = ctk.CTkFrame(backup_outer, fg_color="transparent")
    backup_run_row.pack(fill="x", pady=(0, 8))
    backup_manual_btn = ctk.CTkButton(backup_run_row, text="Backup Now…", width=160,
                                      fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    backup_manual_btn.pack(side="left")
    backup_restore_btn = ctk.CTkButton(backup_run_row, text="Restore Backup…", width=170,
                                       fg_color="#e67e22", hover_color="#c9701c")
    backup_restore_btn.pack(side="left", padx=8)
    backup_last_badge = make_status_badge(backup_run_row, "No backups yet", kind="pending")
    backup_last_badge.pack(side="left", padx=16)
    backup_status_lbl = ctk.CTkLabel(backup_outer, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    backup_status_lbl.pack(anchor="w", pady=(4, 12))

    backup_sched_frame = ctk.CTkFrame(backup_outer)
    backup_sched_frame.pack(fill="x", pady=(0, 12))
    ctk.CTkLabel(backup_sched_frame, text="Scheduled Backup", font=ctk.CTkFont(size=12, weight="bold")).grid(
        row=0, column=0, columnspan=5, padx=12, pady=(10, 4), sticky="w")

    _sched_cfg = load_backup_schedule()
    backup_sched_enabled_var = ctk.BooleanVar(value=_sched_cfg.get("enabled", False))
    ctk.CTkCheckBox(backup_sched_frame, text="Enabled", variable=backup_sched_enabled_var).grid(
        row=1, column=0, padx=12, pady=(4, 12), sticky="w")

    ctk.CTkLabel(backup_sched_frame, text="Frequency:", font=ctk.CTkFont(size=12)).grid(
        row=1, column=1, padx=(12, 6), pady=(4, 12), sticky="w")
    backup_sched_freq_var = ctk.StringVar(value=_sched_cfg.get("frequency", "Daily"))
    ctk.CTkOptionMenu(backup_sched_frame, variable=backup_sched_freq_var, width=110,
                      values=["Daily", "Weekly"]).grid(row=1, column=2, padx=6, pady=(4, 12), sticky="w")

    ctk.CTkLabel(backup_sched_frame, text="At hour (0-23):", font=ctk.CTkFont(size=12)).grid(
        row=1, column=3, padx=(12, 6), pady=(4, 12), sticky="w")
    backup_sched_hour_var = ctk.StringVar(value=str(_sched_cfg.get("hour", 2)))
    ctk.CTkEntry(backup_sched_frame, textvariable=backup_sched_hour_var, width=60).grid(
        row=1, column=4, padx=6, pady=(4, 12), sticky="w")

    backup_sched_save_btn = ctk.CTkButton(backup_sched_frame, text="Save Schedule", width=130)
    backup_sched_save_btn.grid(row=1, column=5, padx=(20, 12), pady=(4, 12))

    ctk.CTkLabel(backup_outer, text="Backup History", font=ctk.CTkFont(size=13, weight="bold")).pack(
        anchor="w", pady=(4, 4))
    backup_history_cols = ("timestamp", "trigger", "size", "files", "validation", "path")
    backup_history_frame = ctk.CTkFrame(backup_outer, fg_color="#181b21")
    backup_history_frame.pack(fill="both", expand=True, pady=(0, 8))
    backup_history_tree = ttk.Treeview(backup_history_frame, columns=backup_history_cols, show="headings", height=8)
    for col, label, w in (("timestamp", "Timestamp", 150), ("trigger", "Trigger", 110),
                          ("size", "Size", 90), ("files", "Files", 60),
                          ("validation", "Validation", 200), ("path", "Path", 260)):
        backup_history_tree.heading(col, text=label,
                                     command=lambda c=col: sort_tree(backup_history_tree, c, False))
        backup_history_tree.column(col, width=w, anchor="w")
    backup_history_tree.pack(fill="both", expand=True)
    backup_history_tree.tag_configure("validation_ok", foreground="#2ecc71")
    backup_history_tree.tag_configure("validation_fail", foreground="#f04747")

    ctk.CTkLabel(backup_outer, text="(Double-click a history row to restore that backup)",
                 font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 4))

    backup_manual_btn.configure(command=do_manual_backup)
    backup_restore_btn.configure(command=do_restore_backup_dialog)
    backup_sched_save_btn.configure(command=do_save_backup_schedule)
    backup_history_tree.bind("<Double-1>", _on_backup_history_row_double_click)

    # ---- LDAP / Active Directory tab ----
    tab_ldap = tabview.add("LDAP / AD")
    ldap_outer = ctk.CTkScrollableFrame(tab_ldap, fg_color="transparent")
    ldap_outer.pack(fill="both", expand=True, padx=14, pady=(6, 14))

    ctk.CTkLabel(ldap_outer, text="LDAP / Active Directory Authentication",
                font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
    ctk.CTkLabel(
        ldap_outer,
        text="Enables Domain Login (shown as an option on the Admin login screen) using a real "
             "LDAP/AD bind. Group membership resolves to one of five roles below; only "
             "Administrators unlocks Admin mode in this app today -- Technicians/Radiologists/"
             "Support/Guest all land in the normal User mode, same as any local user. The local "
             "Admin PIN always keeps working as a fallback regardless of this setting.",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, wraplength=780, justify="left",
    ).pack(anchor="w", pady=(4, 14))

    if not LDAP3_AVAILABLE:
        ctk.CTkLabel(
            ldap_outer,
            text="'ldap3' is not installed -- LDAP/AD login is unavailable until you run: pip install ldap3",
            image=get_icon("triangle-alert", size=14, color="#e67e22"), compound="left",
            font=ctk.CTkFont(size=11), text_color="#e67e22",
        ).pack(anchor="w", pady=(0, 10))

    _ldap_cfg = load_ldap_config()

    ldap_enabled_var = ctk.BooleanVar(value=_ldap_cfg.get("enabled", False))
    ctk.CTkCheckBox(ldap_outer, text="Enable LDAP / AD Authentication", variable=ldap_enabled_var,
                    font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 10))

    ldap_form = ctk.CTkFrame(ldap_outer)
    ldap_form.pack(fill="x", pady=(0, 10))

    def _ldap_row(parent, row, label, width=260):
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=12)).grid(
            row=row, column=0, padx=(12, 8), pady=6, sticky="w")
        entry = ctk.CTkEntry(parent, width=width)
        entry.grid(row=row, column=1, padx=8, pady=6, sticky="w")
        return entry

    ldap_server_entry = _ldap_row(ldap_form, 0, "Server URI:", 320)
    ldap_server_entry.insert(0, _ldap_cfg.get("server_uri", ""))
    ctk.CTkLabel(ldap_form, text="e.g. ldap://dc1.company.local:389 or ldaps://dc1.company.local:636",
                font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).grid(row=0, column=2, padx=8, sticky="w")

    ldap_ssl_var = ctk.BooleanVar(value=_ldap_cfg.get("use_ssl", False))
    ctk.CTkCheckBox(ldap_form, text="Use SSL", variable=ldap_ssl_var).grid(row=1, column=1, padx=8, pady=6, sticky="w")

    ldap_domain_entry = _ldap_row(ldap_form, 2, "AD Domain (UPN bind):")
    ldap_domain_entry.insert(0, _ldap_cfg.get("domain", ""))
    ctk.CTkLabel(ldap_form, text="e.g. company.local -- builds username@domain for AD",
                font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).grid(row=2, column=2, padx=8, sticky="w")

    ldap_bind_template_entry = _ldap_row(ldap_form, 3, "OR Bind DN Template:", 320)
    ldap_bind_template_entry.insert(0, _ldap_cfg.get("user_bind_dn_template", ""))
    ctk.CTkLabel(ldap_form, text="Generic/OpenLDAP style, e.g. cn={username},ou=Users,dc=company,dc=local",
                font=ctk.CTkFont(size=10), text_color=THEME_TEXT_MUTED).grid(row=3, column=2, padx=8, sticky="w")

    ldap_bind_dn_entry = _ldap_row(ldap_form, 4, "Service Account (bind DN):", 320)
    ldap_bind_dn_entry.insert(0, _ldap_cfg.get("bind_dn", ""))

    ldap_bind_pw_entry = ctk.CTkEntry(ldap_form, width=260, show="•")
    ldap_bind_pw_entry.insert(0, _ldap_cfg.get("bind_password", ""))
    ctk.CTkLabel(ldap_form, text="Service Account Password:", font=ctk.CTkFont(size=12)).grid(
        row=5, column=0, padx=(12, 8), pady=6, sticky="w")
    ldap_bind_pw_entry.grid(row=5, column=1, padx=8, pady=6, sticky="w")

    ldap_user_base_entry = _ldap_row(ldap_form, 6, "User Search Base:", 320)
    ldap_user_base_entry.insert(0, _ldap_cfg.get("user_search_base", ""))

    ldap_user_filter_entry = _ldap_row(ldap_form, 7, "User Search Filter:", 320)
    ldap_user_filter_entry.insert(0, _ldap_cfg.get("user_search_filter", "(sAMAccountName={username})"))

    ldap_group_base_entry = _ldap_row(ldap_form, 8, "Group Search Base (optional):", 320)
    ldap_group_base_entry.insert(0, _ldap_cfg.get("group_search_base", ""))

    ctk.CTkLabel(ldap_outer, text="Group → Role Mapping", font=ctk.CTkFont(size=13, weight="bold")).pack(
        anchor="w", pady=(6, 4))
    ctk.CTkLabel(
        ldap_outer,
        text="Enter the LDAP/AD group CN (not the full DN) that should map to each role. Leave blank to skip a role.",
        font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED,
    ).pack(anchor="w", pady=(0, 6))

    ldap_group_map_frame = ctk.CTkFrame(ldap_outer)
    ldap_group_map_frame.pack(fill="x", pady=(0, 10))
    ldap_group_map_entries = {}
    existing_mappings = _ldap_cfg.get("group_mappings", {})
    # Invert existing cn->role mapping to role->cn for pre-filling the form
    role_to_cn = {role: cn for cn, role in existing_mappings.items()}
    for i, role in enumerate(LDAP_ROLE_NAMES):
        ctk.CTkLabel(ldap_group_map_frame, text=f"{role}:", font=ctk.CTkFont(size=12)).grid(
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
    ldap_status_lbl = ctk.CTkLabel(ldap_outer, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    ldap_status_lbl.pack(anchor="w", pady=(6, 12))

    ctk.CTkLabel(ldap_outer, text="Imported User Roster", font=ctk.CTkFont(size=13, weight="bold")).pack(
        anchor="w", pady=(4, 4))
    ldap_roster_cols = ("username", "display_name", "email", "role", "last_login")
    ldap_roster_frame = ctk.CTkFrame(ldap_outer, fg_color="#181b21")
    ldap_roster_frame.pack(fill="both", expand=True, pady=(0, 8))
    ldap_roster_tree = ttk.Treeview(ldap_roster_frame, columns=ldap_roster_cols, show="headings", height=8)
    for col, label, w in (("username", "Username", 130), ("display_name", "Display Name", 180),
                          ("email", "Email", 220), ("role", "Role", 130), ("last_login", "Last Login", 160)):
        ldap_roster_tree.heading(col, text=label)
        ldap_roster_tree.column(col, width=w, anchor="w")
    ldap_roster_tree.pack(fill="both", expand=True)

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
                                  fg_color="#3b82f6", hover_color="#174f7f")
    qr_find_btn.pack(pady=6)

    qr_retrieve_btn = ctk.CTkButton(qr_btn_col, text="C-MOVE Retrieve Selected", width=215,
                                     fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    qr_retrieve_btn.pack(pady=6)

    qr_status_label = ctk.CTkLabel(qr_btn_col, text="", font=ctk.CTkFont(size=11),
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

    qr_tree_frame = ctk.CTkFrame(tab_qr, fg_color="#181b21")
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
                 font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(8, 4))

    dest_listbox_var = ctk.StringVar(value=[])
    dest_listbox = ctk.CTkTextbox(dest_list_frame, width=220, height=350, state="disabled")
    dest_listbox.pack(padx=8, pady=4)

    dest_select_var = ctk.StringVar()
    dest_names_var = ctk.StringVar(value=[])
    dest_optmenu = ctk.CTkOptionMenu(dest_list_frame, variable=dest_select_var,
                                      values=["(none)"], width=200)
    dest_optmenu.pack(pady=4)

    dest_form_frame = ctk.CTkFrame(dest_outer)
    dest_form_frame.pack(side="left", fill="y", padx=(0, 10))

    ctk.CTkLabel(dest_form_frame, text="Destination Profile",
                 font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(8, 4))

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
        # Inline validation: clear the invalid highlight as soon as the
        # user starts fixing the field, instead of only on next submit.
        e.bind("<KeyRelease>", lambda _evt, k=key: _clear_dest_field_invalid(k))

    dest_default_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(dest_form_frame, text="Set as default destination",
                    variable=dest_default_var).pack(anchor="w", padx=10, pady=4)

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

    dest_status_lbl = ctk.CTkLabel(dest_form_frame, text="", font=ctk.CTkFont(size=11),
                                    text_color=THEME_TEXT_MUTED, wraplength=220)
    dest_status_lbl.pack(pady=4)

    # ---- Routing Rules tab (unchanged) ----
    tab_routing = tabview.add("Routing Rules")

    routing_outer = ctk.CTkFrame(tab_routing, fg_color="transparent")
    routing_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    ctk.CTkLabel(routing_outer,
                 text="Rules are evaluated top-down. The first matching rule's destination is used.\n"
                      "Leave a field blank to match any value (wildcard).",
                 font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 8))

    routing_list_frame = ctk.CTkFrame(routing_outer, fg_color="#181b21")
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
    for lbl, key in [("Modality (blank=any)", "modality"), ("Institution (blank=any)", "institution"),
                      ("Source AE (blank=any)", "source_ae")]:
        ctk.CTkLabel(routing_form_frame, text=lbl).pack(side="left", padx=(10, 2))
        e = ctk.CTkEntry(routing_form_frame, width=130)
        e.pack(side="left", padx=(0, 8))
        routing_fields[key] = e

    ctk.CTkLabel(routing_form_frame, text="→ Destination").pack(side="left", padx=(10, 2))
    routing_dest_var = ctk.StringVar(value="(none)")
    routing_dest_menu = ctk.CTkOptionMenu(routing_form_frame, variable=routing_dest_var,
                                           values=["(none)"], width=160)
    routing_dest_menu.pack(side="left", padx=(0, 8))

    routing_add_btn = ctk.CTkButton(routing_form_frame, text="Add Rule", width=120,
                                     fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    routing_add_btn.pack(side="left", padx=4)
    routing_del_btn = ctk.CTkButton(routing_form_frame, text="Delete", width=100,
                                     fg_color=THEME_DANGER, hover_color=THEME_DANGER_HOVER)
    routing_del_btn.pack(side="left", padx=4)
    routing_up_btn = ctk.CTkButton(routing_form_frame, text="▲", width=50)
    routing_up_btn.pack(side="left", padx=2)
    routing_down_btn = ctk.CTkButton(routing_form_frame, text="▼", width=50)
    routing_down_btn.pack(side="left", padx=2)

    # ---- SOP Classes tab (in-GUI editor, unchanged) ----
    tab_sop = tabview.add("SOP Classes")

    sop_outer = ctk.CTkFrame(tab_sop, fg_color="transparent")
    sop_outer.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    ctk.CTkLabel(sop_outer,
                 text="Edit SOP Classes and Transfer Syntaxes. Changes are written to sopclass.ini and take effect on next Receiver start.",
                 font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(0, 8))

    sop_split = ctk.CTkFrame(sop_outer, fg_color="transparent")
    sop_split.pack(fill="both", expand=True)

    sop_left = ctk.CTkFrame(sop_split)
    sop_left.pack(side="left", fill="both", expand=True, padx=(0, 8))
    ctk.CTkLabel(sop_left, text="[SOP_CLASSES]  name = UID",
                 font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=8, pady=4)
    sop_classes_box = ctk.CTkTextbox(sop_left, font=ctk.CTkFont(family="Courier", size=11))
    sop_classes_box.pack(fill="both", expand=True, padx=8, pady=4)

    sop_right = ctk.CTkFrame(sop_split)
    sop_right.pack(side="left", fill="both", expand=True)
    ctk.CTkLabel(sop_right, text="[TRANSFER_SYNTAXES]  name = UID",
                 font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=8, pady=4)
    sop_ts_box = ctk.CTkTextbox(sop_right, font=ctk.CTkFont(family="Courier", size=11))
    sop_ts_box.pack(fill="both", expand=True, padx=8, pady=4)

    sop_btn_row = ctk.CTkFrame(sop_outer, fg_color="transparent")
    sop_btn_row.pack(fill="x", pady=6)
    sop_load_btn = ctk.CTkButton(sop_btn_row, text="↺ Reload from file", width=180)
    sop_load_btn.pack(side="left", padx=8)
    sop_save_btn = ctk.CTkButton(sop_btn_row, text="Save sopclass.ini", width=180,
                                  fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    sop_save_btn.pack(side="left", padx=8)
    sop_status_lbl = ctk.CTkLabel(sop_btn_row, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    sop_status_lbl.pack(side="left", padx=8)

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

    log_archive_now_btn = ctk.CTkButton(logs_top, text="Archive Now", width=130)
    log_archive_now_btn.pack(side="left", padx=4)

    log_export_btn = ctk.CTkButton(logs_top, text="Export Logs…", width=140,
                                    fg_color=THEME_SUCCESS, hover_color=THEME_SUCCESS_HOVER)
    log_export_btn.pack(side="left", padx=4)

    log_tail_var = ctk.BooleanVar(value=True)
    ctk.CTkCheckBox(logs_top, text="Auto-tail", variable=log_tail_var).pack(side="left", padx=10)

    log_copy_btn = ctk.CTkButton(logs_top, text="Copy View", width=120,
                                  fg_color=THEME_NEUTRAL_BTN, hover_color=THEME_NEUTRAL_BTN_HOVER)
    log_copy_btn.pack(side="left", padx=4)

    logs_row1b = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_row1b.pack(fill="x", padx=10, pady=(0, 4))

    ctk.CTkLabel(logs_row1b, text="View:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
    log_display_mode_var = ctk.StringVar(value="Raw Text")
    log_display_mode_menu = ctk.CTkSegmentedButton(
        logs_row1b, values=["Raw Text", "Structured"], variable=log_display_mode_var,
    )
    log_display_mode_menu.pack(side="left", padx=(0, 16))

    ctk.CTkLabel(logs_row1b, text="Search:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
    log_search_var = ctk.StringVar(value="")
    log_search_entry = ctk.CTkEntry(
        logs_row1b, width=280, textvariable=log_search_var,
        placeholder_text="Patient ID/Name, institution, destination, text…")
    log_search_entry.pack(side="left")

    ctk.CTkLabel(logs_row1b, text="Severity:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(16, 6))
    log_severity_var = ctk.StringVar(value="All")
    log_severity_menu = ctk.CTkSegmentedButton(
        logs_row1b, values=["All", "Errors Only", "Success Only"], variable=log_severity_var,
    )
    log_severity_menu.pack(side="left")

    log_match_count_lbl = ctk.CTkLabel(logs_row1b, text="", font=ctk.CTkFont(size=11),
                                        text_color=THEME_TEXT_MUTED)
    log_match_count_lbl.pack(side="left", padx=12)

    logs_row2 = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_row2.pack(fill="x", padx=10, pady=(0, 8))

    ctk.CTkLabel(logs_row2, text="Retention:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
    log_retention_var = ctk.StringVar(value=str(load_log_retention_days()))
    log_retention_menu = ctk.CTkOptionMenu(
        logs_row2, variable=log_retention_var, width=90,
        values=[str(d) for d in VALID_LOG_RETENTION_DAYS],
    )
    log_retention_menu.pack(side="left")
    ctk.CTkLabel(logs_row2, text="days before archiving", font=ctk.CTkFont(size=11),
                 text_color=THEME_TEXT_MUTED).pack(side="left", padx=(4, 20))

    ctk.CTkLabel(logs_row2, text="View:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
    log_view_mode_var = ctk.StringVar(value="Live")
    log_view_mode_menu = ctk.CTkSegmentedButton(
        logs_row2, values=["Live", "Archived"], variable=log_view_mode_var,
    )
    log_view_mode_menu.pack(side="left")

    log_archive_var = ctk.StringVar(value="")
    log_archive_menu = ctk.CTkOptionMenu(logs_row2, variable=log_archive_var, width=280, values=["(no archives yet)"])
    log_archive_menu.pack(side="left", padx=(10, 0))

    log_status_lbl = ctk.CTkLabel(logs_row2, text="", font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED)
    log_status_lbl.pack(side="left", padx=12)

    logs_row3 = ctk.CTkFrame(tab_logs, fg_color="transparent")
    logs_row3.pack(fill="x", padx=10, pady=(0, 4))
    ctk.CTkLabel(logs_row3, text="Destination:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
    log_struct_dest_var = ctk.StringVar(value="All")
    log_struct_dest_menu = ctk.CTkOptionMenu(logs_row3, variable=log_struct_dest_var, width=200,
                                              values=["All"])
    log_struct_dest_menu.pack(side="left", padx=(0, 16))
    ctk.CTkLabel(logs_row3, text="Range:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
    log_struct_range_var = ctk.StringVar(value="Last 7 Days")
    log_struct_range_menu = ctk.CTkOptionMenu(logs_row3, variable=log_struct_range_var, width=140,
                                               values=LOG_STRUCT_RANGE_OPTIONS)
    log_struct_range_menu.pack(side="left")
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
                 font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED).pack(anchor="w", pady=(8, 2))
    log_detail_box = ctk.CTkTextbox(log_structured_container, height=140,
                                     font=ctk.CTkFont(family="Courier", size=11), state="disabled")
    log_detail_box.pack(fill="x")

    log_raw_container.pack(fill="both", expand=True)  # default mode = Raw Text

    # ---- Command wiring for all of the above (moved here so it re-runs
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

    sop_save_btn.configure(command=do_save_sop)
    sop_load_btn.configure(command=lambda: (load_sop_into_editor(),
                                             sop_status_lbl.configure(text="Reloaded from sopclass.ini")))

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
    log_display_mode_var.trace_add("write", _on_log_display_mode_change)
    log_struct_dest_var.trace_add("write", lambda *_: refresh_structured_log_view())
    log_struct_range_var.trace_add("write", lambda *_: refresh_structured_log_view())
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
    win.grab_set()

    def _on_close():
        win.grab_release()
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", _on_close)

    ctk.CTkLabel(win, text=summary, font=ctk.CTkFont(size=12),
                 justify="left", wraplength=380).pack(padx=15, pady=(15, 5))
    ctk.CTkLabel(win, text=f"Verification code:  {code}",
                 font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
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
    populate_tree(rec_tree, search)
    populate_tree(push_tree, search)
    refresh_receiver_badge()


def _log_line_matches_filters(line, search_lower, severity):
    if severity == "Errors Only" and "ERROR=" in line and line.rstrip().endswith("ERROR="):
        return False  # blank ERROR= means success; exclude from Errors Only
    if severity == "Success Only" and "ERROR=" in line and not line.rstrip().endswith("ERROR="):
        return False
    if search_lower and search_lower not in line.lower():
        return False
    return True


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

    if not admin_tabs_active["value"]:
        return

    display = "\n".join(
        f"{'[default] ' if d.get('default') else '           '}{d['name']:20s}  "
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


def load_sop_into_editor():
    if not admin_tabs_active["value"]:
        return
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
        color = "#f0665f" if failed else "#8b93a7"
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
            def on_fail():
                export_status_lbl.configure(text=f"Export failed: {e}", text_color=THEME_DANGER)
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
            def on_fail():
                reports_status_lbl.configure(text=f"Report generation failed: {e}", text_color=THEME_DANGER)
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
                text_color=THEME_SUCCESS if ok else "#f04747",
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
                text_color=THEME_SUCCESS if ok else "#f04747",
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
        row = (u.get("username", ""), u.get("display_name", ""), u.get("email", ""),
              u.get("role", ""), u.get("last_login", ""))
        if username in existing_ids:
            ldap_roster_tree.item(username, values=row)
        else:
            ldap_roster_tree.insert("", "end", iid=username, values=row)
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
                                                text_color="#e67e22")
                refresh_backup_history_ui()
            app.after(0, on_done)
        except Exception as e:
            log_exception("Manual backup failed")
            def on_fail():
                backup_manual_btn.configure(state="normal")
                backup_status_lbl.configure(text=f"Backup failed: {e}", text_color=THEME_DANGER)
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
            def on_fail():
                backup_restore_btn.configure(state="normal")
                backup_status_lbl.configure(text=f"Restore failed: {e}", text_color=THEME_DANGER)
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

    if previous_online is not None and previous_online != ok:
        if ok:
            notify_event("destination_online", "Destination Online", f"{name} is back online.")
        else:
            notify_event("destination_offline", "Destination Offline", f"{name} went offline: {msg}")


def check_destination_health(dest):
    """Blocking C-ECHO + timing against one destination profile dict.
    Always call this from a background thread -- it does real network
    I/O. Updates destination_health_cache via record_destination_health_result."""
    start = time.time()
    try:
        ok, msg = dicom_echo(dest["ae"], dest["ip"], int(dest["port"]),
                              calling_ae=dest.get("calling_ae") or DEFAULT_PUSH_CALLING_AE)
    except Exception as e:
        ok, msg = False, str(e)
    response_time_ms = round((time.time() - start) * 1000, 1)
    record_destination_health_result(dest["name"], ok, msg, response_time_ms)
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
        dash_home_push_status.configure(text="● Sending", text_color="#2f8eff")
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
                folder = os.path.join(OUTPUT_DIR, pid)
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

    for key, canvas in dash_graph_canvases.items():
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

    dash_received_lbl.configure(text=str(received))
    dash_pushed_lbl.configure(text=str(pushed))
    dash_failed_lbl.configure(text=str(failed))
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
            writer.writerow([WL_HEADINGS[c] for c in WL_COLUMNS])
            for iid in row_ids:
                writer.writerow(tree.item(iid, "values"))
        modern_showinfo("Export View to CSV", f"Exported {len(row_ids)} row(s) to:\n{out_path}")
    except Exception as e:
        log_exception("Failed to export worklist view to CSV")
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


def do_push_selected():
    pids = list(push_tree.selection())
    if not pids:
        modern_showwarning("No Selection", "Select at least one patient to push.")
        return
    if push_job["running"]:
        modern_showwarning("Busy", "A push job is already running.")
        return
    dest = _get_active_dest()
    anon = anon_var.get()
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
    dest = _get_active_dest()
    anon = anon_var.get()
    threading.Thread(target=run_push_job, args=(pids,), kwargs={"destination": dest, "anonymize": anon}, daemon=True).start()


def do_echo_active_dest():
    dest = _get_active_dest()
    if not dest:
        modern_showerror("Error", "No destination configured. Add one in the Destinations tab.")
        return
    echo_btn.configure(state="disabled", text="Testing...")

    def run():
        ok, msg = dicom_echo(dest["ae"], dest["ip"], dest["port"],
                             calling_ae=dest.get("calling_ae"))
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
    folder = os.path.join(OUTPUT_DIR, pid)
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
    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", lambda: (win.grab_release(), win.destroy()))

    cols = ("study_uid", "modality", "date", "count")
    hdgs = {"study_uid": "Study UID", "modality": "Modality", "date": "Study Date", "count": "#Images"}
    widths = {"study_uid": 320, "modality": 80, "date": 100, "count": 70}

    tree_f = ctk.CTkFrame(win, fg_color="#181b21")
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
    """Returns the WL_COLUMNS name under the given click event, or None."""
    if tree.identify_region(event.x, event.y) != "cell":
        return None
    col_id = tree.identify_column(event.x)  # e.g. '#6'
    try:
        col_index = int(col_id.replace("#", "")) - 1
        return WL_COLUMNS[col_index]
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
            pass
    return None


def _save_weasis_path(path):
    try:
        with open(WEASIS_PATH_FILE, "w", encoding="utf-8") as f:
            f.write(path)
    except Exception:
        pass


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
            pass
    return None


def _save_radiant_path(path):
    try:
        with open(RADIANT_PATH_FILE, "w", encoding="utf-8") as f:
            f.write(path)
    except Exception:
        pass


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
            pass
    return None


def _save_remembered_viewer(choice):
    try:
        with open(VIEWER_PREF_FILE, "w", encoding="utf-8") as f:
            f.write(choice)
    except Exception:
        pass


def _patient_dcm_folder_and_files(pid):
    """Shared lookup used by every viewer launcher: returns (folder,
    dcm_files) or (None, None) after showing the appropriate error, so
    each viewer doesn't repeat the same folder/empty-folder checks."""
    folder = os.path.join(OUTPUT_DIR, pid)
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
    win.grab_set()

    ctk.CTkLabel(win, text=f"Open patient {pid} in:",
                 font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(20, 14))

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

    box = ctk.CTkTextbox(win, wrap="word", font=ctk.CTkFont(size=13))
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


def build_context_menu(tree, event):
    selected = tree.selection()
    if not selected:
        return
    pid = selected[0]

    import tkinter as tk
    menu = tk.Menu(app, tearoff=0)
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
                threading.Thread(
                    target=run_push_job, args=([pid],),
                    kwargs={"destination": dest, "anonymize": anon_var.get()},
                    daemon=True).start()
            dest_menu.add_command(label=f"{d['name']} ({d['ae']}@{d['ip']}:{d['port']})", command=push_to)
        menu.add_cascade(label="Push to...", menu=dest_menu)
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
        if modern_askyesno("Delete", f"Delete {pid} from worklist AND disk?\nThis cannot be undone."):
            with data_lock:
                patient_data.pop(pid, None)
            autosave_csv()
            _bump_data_version()
            folder = os.path.join(OUTPUT_DIR, pid)
            if os.path.isdir(folder):
                shutil.rmtree(folder, ignore_errors=True)
            write_audit_log("DELETE", f"pid={pid} folder={folder}")
            refresh_worklists()

    menu.add_command(label="Delete from worklist + disk", command=delete_patient)

    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


rec_tree.bind("<Button-3>", lambda e: build_context_menu(rec_tree, e))
push_tree.bind("<Button-3>", lambda e: build_context_menu(push_tree, e))
rec_tree.bind("<Button-2>", lambda e: build_context_menu(rec_tree, e))   # macOS
push_tree.bind("<Button-2>", lambda e: build_context_menu(push_tree, e))

# =========================================================
# CONTEXTUAL SELECTION TOOLBAR
# =========================================================
# A slim action bar that appears above a worklist tree only while rows
# are selected, and disappears otherwise -- the "toolbar transforms into
# contextual actions on selection" pattern, built on the same
# run_push_job / delete / reset actions the right-click menu already
# uses (no new business logic, just a faster multi-row entry point to it).

_contextual_toolbars = {}  # id(tree) -> CTkFrame


def _bulk_push_selected(tree):
    pids = list(tree.selection())
    if not pids:
        return
    threading.Thread(
        target=run_push_job, args=(pids,),
        kwargs={"destination": _get_active_dest(), "anonymize": anon_var.get()},
        daemon=True).start()


def _bulk_reset_selected(tree):
    for pid in tree.selection():
        reset_patient_status(pid)


def _bulk_copy_ids_selected(tree):
    pids = list(tree.selection())
    if not pids:
        return
    app.clipboard_clear()
    app.clipboard_append("\n".join(pids))


def _bulk_delete_selected(tree):
    pids = list(tree.selection())
    if not pids:
        return
    if not modern_confirm(
        "Delete Selected", f"Delete {len(pids)} patient(s) from worklist AND disk?\nThis cannot be undone.",
        danger=True,
    ):
        return
    for pid in pids:
        with data_lock:
            patient_data.pop(pid, None)
        folder = os.path.join(OUTPUT_DIR, pid)
        if os.path.isdir(folder):
            shutil.rmtree(folder, ignore_errors=True)
        write_audit_log("DELETE", f"pid={pid} folder={folder}")
    autosave_csv()
    _bump_data_version()
    refresh_worklists()


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

    def _on_select(_evt=None):
        n = len(tree.selection())
        if n > 0:
            count_lbl.configure(text=f"{n} selected")
            if not bar.winfo_ismapped():
                bar.pack(fill="x", padx=10, pady=(0, 8), before=wl_frame)
        else:
            bar.pack_forget()

    tree.bind("<<TreeviewSelect>>", _on_select, add="+")
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
        return
    pid = sel[0]
    inspector_current_pid["value"] = pid
    open_inspector()
    refresh_inspector(pid)


rec_tree.bind("<<TreeviewSelect>>", lambda _e: _on_inspector_selection(rec_tree), add="+")
push_tree.bind("<<TreeviewSelect>>", lambda _e: _on_inspector_selection(push_tree), add="+")

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

DEFAULT_ALWAYS_VISIBLE = {"patient_id", "patient_name", "status"}


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
        with open(VIEW_OPTIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except Exception:
        log_exception("Failed to save view_options.json")


_load_view_options()


def _apply_column_visibility(tree):
    vis = _column_visibility[id(tree)]
    tree.configure(displaycolumns=[c for c in WL_COLUMNS if vis.get(c, True)])


def _apply_density():
    style = ttk.Style()
    rowheight = 22 if _density_state["mode"] == "Compact" else 28
    style.configure("Treeview", rowheight=rowheight)


_apply_column_visibility(rec_tree)
_apply_column_visibility(push_tree)
_apply_density()


def _close_popover(win):
    if win is not None and win.winfo_exists():
        win.destroy()


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
    x = anchor_btn.winfo_rootx()
    y = anchor_btn.winfo_rooty() + anchor_btn.winfo_height() + 4
    win.geometry(f"220x{60 + 26 * len(WL_COLUMNS)}+{x}+{y}")

    shell = ctk.CTkFrame(win, fg_color=THEME_SURFACE, corner_radius=10,
                          border_width=1, border_color=THEME_HEADING_BG)
    shell.pack(fill="both", expand=True, padx=1, pady=1)

    ctk.CTkLabel(shell, text="Density", font=get_font("small", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=12, pady=(10, 2))
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


rec_view_options_btn.configure(command=lambda: open_view_options_popover(rec_tree, rec_view_options_btn))
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
    win.grab_set()
    win.configure(fg_color=THEME_SURFACE)
    win.geometry("380x180")
    app.update_idletasks()
    x = app.winfo_rootx() + (app.winfo_width() - 380) // 2
    y = app.winfo_rooty() + (app.winfo_height() - 180) // 2
    win.geometry(f"+{x}+{y}")
    win.resizable(False, False)

    ctk.CTkLabel(win, text=title, font=get_font("section", "bold"),
                 text_color=THEME_TEXT).pack(anchor="w", padx=20, pady=(20, 4))
    ctk.CTkLabel(win, text=message, font=get_font("body"), text_color=THEME_TEXT_MUTED,
                 wraplength=340, justify="left").pack(anchor="w", padx=20, pady=(0, 16))

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
    color = {"error": THEME_DANGER, "warning": "#e6a23c", "info": THEME_ACCENT}.get(kind, THEME_ACCENT)

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title(title)
    win.attributes("-topmost", True)
    win.transient(app)
    win.grab_set()
    win.configure(fg_color=THEME_SURFACE)
    win.resizable(False, False)

    header = ctk.CTkFrame(win, fg_color="transparent")
    header.pack(fill="x", padx=20, pady=(20, 4))
    ctk.CTkLabel(header, text="", image=get_icon(icon_name, size=20, color=color)).pack(side="left", padx=(0, 8))
    ctk.CTkLabel(header, text=title, font=get_font("section", "bold"), text_color=color).pack(side="left")

    ctk.CTkLabel(win, text=message, font=get_font("body"), text_color=THEME_TEXT_MUTED,
                 wraplength=340, justify="left").pack(anchor="w", padx=20, pady=(0, 16))

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
    new_entry = {
        "name": name,
        "ae": ae,
        "calling_ae": calling_ae,  # empty string = use default "RAPPS_PUSH"
        "ip": ip,
        "port": port,
        "default": dest_default_var.get(),
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


def do_echo_dest():
    name = dest_fields["name"].get().strip() or dest_select_var.get()
    ae = dest_fields["ae"].get().strip()
    ip = dest_fields["ip"].get().strip()
    port = dest_fields["port"].get().strip()
    if not ae or not ip or not port.isdigit():
        modern_showerror("Error", "Fill in the destination fields first.")
        return
    dest_echo_btn.configure(state="disabled", text="Testing...")

    def run():
        calling_ae_val = dest_fields.get("calling_ae")
        calling_ae_str = calling_ae_val.get().strip() if calling_ae_val else None
        ok, msg = dicom_echo(ae, ip, port, calling_ae=calling_ae_str)
        def on_ui():
            dest_echo_btn.configure(state="normal", text="C-ECHO Test")
            dest_status_lbl.configure(text=msg, text_color=THEME_SUCCESS if ok else "#f04747")
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

    def run():
        ok, results, msg = query_remote_studies(
            rae, rip, rport,
            patient_id=qr_filter_entries["pid"].get().strip(),
            patient_name=qr_filter_entries["pname"].get().strip(),
            study_date=qr_filter_entries["date"].get().strip(),
            modality=qr_filter_entries["mod"].get().strip(),
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


def pump_events():
    """Drain the UI event queue and apply all pending updates on the
    main thread. This makes every background thread's state visible to
    the GUI without requiring explicit app.after() calls in each worker."""
    try:
        while True:
            event, payload = ui_event_queue.get_nowait()

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

            elif event == "push_started":
                overall_progress.set(0)
                overall_progress_label.configure(text="Starting push...")
                throughput_label.configure(text="")
                refresh_pusher_monitoring()

            elif event == "push_progress":
                sent = push_job["sent_images"]
                attempted = push_job["attempted_images"]
                total = push_job["total_images"]
                frac = attempted / max(total, 1)
                overall_progress.set(frac)
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

            elif event == "toast":
                title, message = payload
                _show_toast(title, message)

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

    except queue.Empty:
        pass
    except Exception:
        log_exception("Error in event pump")

    # Auto-tail active log file while Logs tab is selected (Admin-only tab)
    if admin_tabs_active["value"] and tabview.get().startswith("") and log_tail_var.get():
        if log_display_mode_var.get() == "Structured":
            refresh_structured_log_view()
        else:
            refresh_log_view()

    app.after(REFRESH_INTERVAL_MS, pump_events)


# Periodic background tree refresh (fallback for events that might be missed)
_refresh_counter = [0]


_KNOWN_TREE_VAR_NAMES = [
    "rec_tree", "push_tree", "routing_tree", "qr_tree",
    "dash_failures_tree", "dash_dest_health_tree", "health_tree", "offlineq_tree",
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


def apply_theme(mode, persist=True):
    """Switches the whole app between Dark/Light/System immediately, no
    restart required, and persists the choice.

    What updates live:
      * Every ttk.Treeview in the app (worklists, PACS Health, Offline
        Queue, Routing Rules, Query/Retrieve, Admin Dashboard tables) --
        these read from a shared ttk.Style, which DOES apply retroactively
        to already-built widgets, so one style update re-themes all of
        them at once.
      * The window background, header bar, admin bar, and tab bar --
        these were built with explicit hex colors (not CTk's theme-aware
        color tuples), so they're recolored directly here.
      * customtkinter's own appearance mode (affects any CTk widget that
        was NOT given an explicit override color, e.g. default frames)."""
    global THEME_BG, THEME_SURFACE, THEME_HEADING_BG, THEME_ACCENT, THEME_ACCENT_HOVER
    global THEME_TEXT, THEME_TEXT_MUTED, THEME_NEUTRAL_BTN, THEME_NEUTRAL_BTN_HOVER
    global THEME_DANGER, THEME_DANGER_HOVER, THEME_SUCCESS, THEME_SUCCESS_HOVER
    global STALE_HIGHLIGHT_COLOR, SEARCH_MATCH_HIGHLIGHT_BG

    resolved = _resolve_system_theme() if mode == "System" else mode.lower()
    palette = THEME_PALETTES.get(resolved, THEME_PALETTES["dark"])

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
    STALE_HIGHLIGHT_COLOR = palette["stale"]
    SEARCH_MATCH_HIGHLIGHT_BG = palette["search_highlight"]

    try:
        ctk.set_appearance_mode(resolved)
    except Exception:
        log_exception("Failed to set customtkinter appearance mode")

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
    except Exception:
        log_exception("Failed to apply theme to top-level chrome")

    try:
        style = ttk.Style()
        style.configure("Treeview", background=THEME_SURFACE, foreground=THEME_TEXT,
                        fieldbackground=THEME_SURFACE)
        style.configure("Treeview.Heading", background=THEME_HEADING_BG, foreground=THEME_TEXT)
        style.map("Treeview.Heading", background=[("active", palette["segmented_hover"])])
        style.map("Treeview", background=[("selected", THEME_ACCENT)], foreground=[("selected", "#ffffff")])
        style.configure("Vertical.TScrollbar", background=THEME_SURFACE, troughcolor=THEME_BG,
                        bordercolor=THEME_SURFACE, arrowcolor=THEME_TEXT_MUTED)
        style.configure("Horizontal.TScrollbar", background=THEME_SURFACE, troughcolor=THEME_BG,
                        bordercolor=THEME_SURFACE, arrowcolor=THEME_TEXT_MUTED)

        for tree in _all_known_trees():
            tree.tag_configure(STALE_HIGHLIGHT_TAG, foreground=STALE_HIGHLIGHT_COLOR)
            tree.tag_configure(SEARCH_MATCH_HIGHLIGHT_TAG, background=SEARCH_MATCH_HIGHLIGHT_BG)
            tree.tag_configure("even_row", background=THEME_SURFACE)
            tree.tag_configure("odd_row", background=palette["odd_row"])
    except Exception:
        log_exception("Failed to apply theme to treeviews")

    if persist:
        save_theme_preference(mode)


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
    # panel down on focus loss.
    try:
        if app.focus_get() is header_search_entry:
            return
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


def periodic_refresh():
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
    if admin_tabs_active["value"] and tabview.get().startswith(""):
        refresh_admin_dashboard()
    # Same visibility gating -- computing this involves a log scan +
    # several psutil calls, so it only runs while someone's looking at it.
    if admin_tabs_active["value"] and tabview.get().startswith(""):
        try:
            refresh_performance_tab()
        except Exception:
            log_exception("Performance tab refresh failed")
    if admin_tabs_active["value"] and tabview.get().startswith(""):
        try:
            refresh_backup_history_ui()
        except Exception:
            log_exception("Backup history refresh failed")
    if admin_tabs_active["value"] and tabview.get().startswith(""):
        try:
            refresh_ldap_roster_ui()
        except Exception:
            log_exception("LDAP roster refresh failed")
    # Cheap background guard for immutable log rotation/archival -- this
    # itself only actually stats the log files once every
    # _ROTATION_CHECK_INTERVAL_SEC seconds, so it's safe to call on every
    # tick without adding real overhead.
    try:
        rotate_logs_if_needed()
    except Exception:
        log_exception("Background log rotation check failed")
    try:
        refresh_health_tree()
    except Exception:
        log_exception("PACS health tree refresh failed")
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
    app.after(REFRESH_INTERVAL_MS, periodic_refresh)

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


def do_admin_login():
    ldap_cfg = load_ldap_config()
    show_ldap_tab = ldap_cfg.get("enabled", False) and LDAP3_AVAILABLE

    win = ctk.CTkToplevel(app)
    _keep_toplevel_small(win)
    win.title("Admin Login")
    win.geometry("380x300" if show_ldap_tab else "340x180")
    win.transient(app)
    win.grab_set()

    def _do_pin_login(pin_entry, err_lbl):
        pin = pin_entry.get().strip()
        if verify_admin_pin(pin):
            write_audit_log("ADMIN-LOGIN-SUCCESS", "Admin PIN accepted")
            current_identity.update(username=None, display_name=None, source="local", ldap_role=None)
            admin_session["unlocked"] = True
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
            set_view("admin")
            show_toast_threadsafe("Admin Unlocked", f"{message} You are now viewing the app in Admin mode.")
        else:
            show_toast_threadsafe(
                "Domain Login Successful",
                f"{message} Your role ({info['role']}) uses standard User mode in this app.",
            )

    if not show_ldap_tab:
        ctk.CTkLabel(win, text="Enter Admin PIN", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(16, 8))
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

    ctk.CTkLabel(tab_pin, text="Enter Admin PIN", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(14, 8))
    pin_entry = ctk.CTkEntry(tab_pin, width=200, show="•")
    pin_entry.pack(pady=4)
    pin_err_lbl = ctk.CTkLabel(tab_pin, text="", text_color=THEME_DANGER)
    pin_err_lbl.pack(pady=4)
    pin_entry.bind("<Return>", lambda e: _do_pin_login(pin_entry, pin_err_lbl))
    ctk.CTkButton(tab_pin, text="Unlock", command=lambda: _do_pin_login(pin_entry, pin_err_lbl)).pack(pady=10)

    ctk.CTkLabel(tab_domain, text="Domain Username", font=ctk.CTkFont(size=12)).pack(pady=(14, 2))
    domain_user_entry = ctk.CTkEntry(tab_domain, width=220)
    domain_user_entry.pack(pady=2)
    ctk.CTkLabel(tab_domain, text="Password", font=ctk.CTkFont(size=12)).pack(pady=(8, 2))
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
    win.grab_set()

    ctk.CTkLabel(win, text="Change Admin PIN", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(16, 10))
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
    win.grab_set()

    ctk.CTkLabel(win, text="Welcome — set an Admin PIN",
                 font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(20, 4))
    ctk.CTkLabel(win, text="This PIN protects Admin-only settings.\n"
                           f"Numeric, at least {ADMIN_PIN_MIN_DIGITS} digits.",
                 font=ctk.CTkFont(size=11), text_color=THEME_TEXT_MUTED, justify="center").pack(pady=(0, 14))

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
    if current.startswith(""):
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

# =========================================================
# STARTUP
# =========================================================

def startup():
    """Run all initialisation that needs the GUI to exist first."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    generate_key()

    # Show the main window FIRST. A Toplevel dialog parented to a still-
    # withdrawn root can fail to display or grab focus on some platforms,
    # which left the app stuck invisibly waiting on a PIN dialog no one
    # could see (looked like "running but nothing opens"). Deiconifying
    # before showing the PIN dialog fixes that.
    app.deiconify()
    app.lift()
    app.attributes("-topmost", True)
    app.after(200, lambda: app.attributes("-topmost", False))
    app.update()

    # One-time Admin PIN setup (now shown on top of a visible main window).
    ensure_admin_pin_configured()

    try:
        load_csv()
        _migrate_legacy_push_config()

        _load_receiver_config()
        refresh_destinations_ui()
        refresh_routing_ui()
        load_sop_into_editor()
        populate_tree(rec_tree)
        populate_tree(push_tree)
        refresh_log_view()

        # App always launches in User mode (admin_session starts locked).
        apply_receiver_mode_visibility()
        update_admin_bar_ui()

        # Check disk on startup
        check_disk_space_and_warn()

        # Continuous PACS health monitoring runs regardless of admin/user
        # mode -- it just populates destination_health_cache in the
        # background; the PACS Health tab (admin-only) and the landing
        # Dashboard both read from that same cache.
        start_pacs_health_monitor_thread()
        start_offline_queue_worker_thread()
        start_backup_scheduler_thread()

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
            app.deiconify()
        except Exception:
            pass
        modern_showerror(
            "Startup error",
            "A fatal error occurred during startup. Check app.log for details.\n\n"
            + traceback.format_exc()[-800:]
        )
    app.after(500, pump_events)
    app.after(REFRESH_INTERVAL_MS, periodic_refresh)
    app.mainloop()
