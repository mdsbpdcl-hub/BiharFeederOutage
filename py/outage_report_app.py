# -*- coding: utf-8 -*-
"""
Bihar 11 kV Feeder Outage Report  -  desktop app
================================================
Reads the daily feeder outage CSV files of each company (any number, any
size), makes an Excel report for each DISCOM (SBPDCL, NBPDCL) and one for
Bihar, and sends each DISCOM's report to the Google Sheet through the
Apps Script web app. Each DISCOM has its own tabs there.

Uses only what comes with Python (no pip install needed), so it can be
turned into a single .exe with the build file.

Run:   python outage_report_app.py
"""

import csv
import datetime
import io
import json
import os
import queue
import re
import sqlite3
try:
    import ssl
except ImportError:                    # the browser page runs this file too, without this module
    ssl = None
import sys
import tempfile
import threading
import time
import traceback
import urllib.error
import urllib.request
import webbrowser
import zipfile
from xml.etree import ElementTree
from xml.sax.saxutils import escape as xml_escape

# =============================================================================
#  SETTINGS YOU MAY CHANGE
# =============================================================================

# Paste the web app link here (it ends with /exec). See the set-up notes.
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxzXRkOmfOwReMdvNAW8Baz2h4ppDr7sTP3mO8KvACq6tS7eeDEhXWPHgLqGMhNaWxoXw/exec"

APP_TITLE = "Bihar 11 kV Feeder Outage Report"
APP_VERSION = "1.7"

OVER_HOURS = 4                 # "more than 4 hrs" list (strictly more than)
MANY_INTERRUPTIONS = 10        # a feeder with this many or more interruptions
TOP_N_SHEETS = 5               # Top 5 sheets
RANK_N_STATE = 20              # Feeder Ranks: top and bottom N for Bihar / DISCOM
RANK_N_DIVISION = 10           # Feeder Ranks: top N inside each division
MIN_FEEDERS_DIVISION = 10      # divisions with fewer feeders are not ranked
MIN_FEEDERS_SUBSTATION = 3     # substations with fewer feeders are not ranked
MAX_SECONDS_IN_DAY = 24 * 3600

# The companies that give the outage CSV files, for each DISCOM.
# The window shows one "Browse" line for each name, in this order.
COMPANIES = {
    "NBPDCL": ["NCC", "Adani", "Hi-Print", "Secure"],
    "SBPDCL": ["IntelliSmart", "Genus"],
}
SECTION_ORDER = ["NBPDCL", "SBPDCL"]            # left and right part of the window

# Feeder types that are NOT counted in the averages, for each DISCOM.
# Every other 11 kV feeder is counted. 33 kV feeders are never counted.
#   SBPDCL counts Urban, Rural and Mixed (Agriculture and Industrial are left out).
#   NBPDCL counts every 11 kV feeder except Agriculture.
TYPES_LEFT_OUT = {
    "SBPDCL": ["Agriculture", "Industrial"],
    "NBPDCL": ["Agriculture"],
}
TYPES_LEFT_OUT_OTHER = ["Agriculture", "Industrial"]      # for a circle whose DISCOM is not known

# Divisions for which no row came in the files.
#   False = every division of the feeder list is kept. One with no rows counts as "no outage"
#           and its feeders stay in Total Feeders (as in the SBPDCL Excel report).
#   True  = such a division (or circle) is left out of the averages and marked "Data Received = No".
LEAVE_OUT_DIVISIONS_WITH_NO_ROWS = False

# A file that comes with NO heading row. The app knows it by the number of columns
# and by what sits in them. The numbers are the column numbers, counted from 1.
NO_HEADING_LAYOUTS = [
    {"name": "NCC", "columns": 16,
     "feeder": 2, "fcode": 3, "voltage": 4, "ftype": 5, "date": 6, "events": 7, "duration": 10,
     "substation": 11, "circle": 13, "division": 14, "subdivision": 15},
]

# The app window. It is drawn to fit the screen: text and buttons grow on a big screen.
WINDOW_FILL = 0.92             # how much of the free screen the window takes (1.0 = all of it, 0.75 = smaller)
WINDOW_MAX_ZOOM = 1.6          # the window is never drawn more than this many times its plain size

MASTER_FILE_NAME = "feeder_master.csv"
TYPE_FILE_NAME = "feeder_type_corrections.csv"

# THE AGRICULTURE LIST: feeders whose type in the company files is wrong (for example an Agriculture feeder typed Rural).
# It is ONE shared list, kept in the Google Sheet (tab "Agriculture List") and changed on the web page
# (LIST_PAGE_URL) with the DISCOM's upload password. Before every report the app reads it from Google;
# it keeps a copy next to the app (AGRI_CACHE_NAME) for a day when Google cannot be reached.
# The list below is used only if neither Google nor the saved copy can be read.
LIST_PAGE_URL = "https://mdsbpdcl-hub.github.io/BiharFeederOutage/update.html#list"
AGRI_CACHE_NAME = "agriculture_list_saved_copy.json"
BUILTIN_TYPE_CORRECTIONS = [          # circle, feeder code, meter no, feeder, correct type
    ("BIHARSARIF", "F4426", "SBIHF5000900", "Krishi Feeder", "Agriculture"),
    ("BIHARSARIF", "F2888", "SBIHF5000681", "DAHPAR", "Agriculture"),
    ("SASARAM", "F6129", "SBISF1000100", "3KV O/G Nokha PSS", "Agriculture"),
    ("SASARAM", "F2399", "SBIHF5000946", "HYDAL", "Agriculture"),
    ("BHAGALPUR", "2082451BW106", "7076241", "11KV OG AMARPUR AG I", "Agriculture"),
    ("JAMUI", "2082391CS101", "9929078", "11KV OG JHUNDO AGRICULTURE", "Agriculture"),
    ("JAMUI", "2082394CO101", "7076141", "11KV OG KEDIA AGRICULTURE", "Agriculture"),
    ("JAMUI", "2082381DZ103", "9929084", "11KV OG SIYANI", "Agriculture"),
]
# The browser page puts the list it read from Google here before it makes a report.
SHARED_TYPE_ROWS = None
# A file feeder_type_corrections.csv next to the app could be used instead (old way). Not used unless this is True.
USE_TYPE_LIST_FILE = False

# Fixed number of 11 kV feeders in each division, as used for the averages.
# SBPDCL: from the SBPDCL Excel report of 10.10.2026.  NBPDCL: from the NBPDCL PDF report of 09.10.2026.
# (DISCOM, Circle, Division, Total Feeders, other names used for the division)
BUILTIN_MASTER = [
    ("SBPDCL", "ARRAH", "ARRAH", 69, ""),
    ("SBPDCL", "ARRAH", "BUXAR", 106, ""),
    ("SBPDCL", "ARRAH", "JAGDISHPUR", 62, ""),
    ("SBPDCL", "AURANGABAD", "ARWAL", 35, ""),
    ("SBPDCL", "AURANGABAD", "AURANGABAD", 119, ""),
    ("SBPDCL", "AURANGABAD", "DAUDNAGAR", 62, ""),
    ("SBPDCL", "BHAGALPUR", "AMARPUR", 38, ""),
    ("SBPDCL", "BHAGALPUR", "BANKA", 50, ""),
    ("SBPDCL", "BHAGALPUR", "BHAGALPUR(E)", 64, "BHAGALPUR EAST"),
    ("SBPDCL", "BHAGALPUR", "BHAGALPUR(U)", 66, "BHAGALPUR URBAN"),
    ("SBPDCL", "BHAGALPUR", "NAUGACHIA", 33, ""),
    ("SBPDCL", "BIHARSARIF", "BIHARSHARIFF RURAL", 89, "BIHARSHARIF RURAL|BIHARSARIF RURAL|BIHARSHARIF-R"),
    ("SBPDCL", "BIHARSARIF", "BIHARSHARIF-U", 35, "BIHARSHARIF URBAN|BIHARSHARIFF URBAN|BIHARSARIF URBAN|BIHARSHARIFF-U"),
    ("SBPDCL", "BIHARSARIF", "EKANGARSARAI", 54, ""),
    ("SBPDCL", "BIHARSARIF", "NAWADA", 60, ""),
    ("SBPDCL", "BIHARSARIF", "RAJAULI", 47, ""),
    ("SBPDCL", "BIHARSARIF", "RAJGIR", 51, ""),
    ("SBPDCL", "GAYA JI", "GAYA JI RURAL", 70, "GAYA RURAL|GAYAR|GAYA(R)"),
    ("SBPDCL", "GAYA JI", "GAYA JI URBAN", 66, "GAYA URBAN|GAYAU|GAYA(U)"),
    ("SBPDCL", "GAYA JI", "JAHANABAD", 116, "JEHANABAD"),
    ("SBPDCL", "GAYA JI", "MANPUR", 78, ""),
    ("SBPDCL", "GAYA JI", "SHERGHATI", 62, ""),
    ("SBPDCL", "JAMUI", "JAMUI", 116, ""),
    ("SBPDCL", "JAMUI", "SHEIKHPURA", 93, ""),
    ("SBPDCL", "MUNGER", "LAKHISARAI", 72, ""),
    ("SBPDCL", "MUNGER", "MUNGER", 82, ""),
    ("SBPDCL", "PATNA", "BARH", 49, ""),
    ("SBPDCL", "PATNA", "BIHTA", 66, ""),
    ("SBPDCL", "PATNA", "FATUHA", 35, ""),
    ("SBPDCL", "PATNA", "MASAURHI", 53, ""),
    ("SBPDCL", "PATNA", "PATNA", 15, ""),
    ("SBPDCL", "PESU EAST", "BANKIPUR", 21, ""),
    ("SBPDCL", "PESU EAST", "GULZARBAGH", 26, ""),
    ("SBPDCL", "PESU EAST", "KANKARBAGH 1", 35, "KANKARBAGH-1|KANKARBAGH I"),
    ("SBPDCL", "PESU EAST", "KANKARBAGH 2", 35, "KANKARBAGH-2|KANKARBAGH II"),
    ("SBPDCL", "PESU EAST", "PATNA CITY", 19, ""),
    ("SBPDCL", "PESU EAST", "RAJENDRA NAGAR", 18, ""),
    ("SBPDCL", "PESU WEST", "ASHIYANA", 29, ""),
    ("SBPDCL", "PESU WEST", "DAKBANGLOW", 13, "DAKBUNGLOW"),
    ("SBPDCL", "PESU WEST", "DANAPUR", 32, ""),
    ("SBPDCL", "PESU WEST", "GARDANIBAGH", 34, ""),
    ("SBPDCL", "PESU WEST", "KHAGAUL", 39, ""),
    ("SBPDCL", "PESU WEST", "NEW CAPITAL", 50, ""),
    ("SBPDCL", "PESU WEST", "PATLIPUTRA", 31, ""),
    ("SBPDCL", "SASARAM", "BHABHUA", 118, "BHABUA"),
    ("SBPDCL", "SASARAM", "DEHARI", 88, "DEHRI"),
    ("SBPDCL", "SASARAM", "SASARAM", 123, ""),
    ("NBPDCL", "BEGUSARAI", "BARAUNI", 61, ""),
    ("NBPDCL", "BEGUSARAI", "BEGUSARAI", 76, ""),
    ("NBPDCL", "BEGUSARAI", "KHAGARIA", 48, ""),
    ("NBPDCL", "CHAPRA", "CHAPRA-EAST", 79, "CHAPRA EAST|CHAPRA(E)"),
    ("NBPDCL", "CHAPRA", "CHAPRA WEST", 71, "CHAPRA-WEST|CHAPRA(W)"),
    ("NBPDCL", "CHAPRA", "GOPALGANJ", 57, ""),
    ("NBPDCL", "CHAPRA", "MAHARAJGANJ", 57, ""),
    ("NBPDCL", "CHAPRA", "MIRGANJ", 65, ""),
    ("NBPDCL", "CHAPRA", "SIWAN", 68, ""),
    ("NBPDCL", "DARBHANGA", "BENIPUR", 56, ""),
    ("NBPDCL", "DARBHANGA", "DARBHANGA RURAL", 75, "DARBHANGA(R)"),
    ("NBPDCL", "DARBHANGA", "DARBHANGA URBAN", 41, "DARBHANGA(U)"),
    ("NBPDCL", "DARBHANGA", "JAYNAGAR", 41, "JAINAGAR"),
    ("NBPDCL", "DARBHANGA", "JHANJHARPUR", 56, ""),
    ("NBPDCL", "DARBHANGA", "MADHUBANI", 72, ""),
    ("NBPDCL", "KISHANGANJ", "ARARIA", 60, ""),
    ("NBPDCL", "KISHANGANJ", "BAHADURGANJ", 62, ""),
    ("NBPDCL", "KISHANGANJ", "FORBESGANJ", 43, ""),
    ("NBPDCL", "KISHANGANJ", "KISHANGANJ", 38, ""),
    ("NBPDCL", "MOTIHARI", "BAGHA", 55, ""),
    ("NBPDCL", "MOTIHARI", "BETTIAH", 99, ""),
    ("NBPDCL", "MOTIHARI", "CHAKIA", 73, ""),
    ("NBPDCL", "MOTIHARI", "MOTIHARI", 73, ""),
    ("NBPDCL", "MOTIHARI", "RAXAUL", 46, ""),
    ("NBPDCL", "MUZAFFARPUR", "MUZAFFARPUR(E)", 59, "MUZAFFARPUR EAST"),
    ("NBPDCL", "MUZAFFARPUR", "MUZAFFARPUR(U-I)", 44, "MUZAFFARPUR URBAN 1|MUZAFFARPUR URBAN I|MUZAFFARPUR(U1)|MUZAFFARPUR U-1"),
    ("NBPDCL", "MUZAFFARPUR", "MUZAFFARPUR(U-II)", 50, "MUZAFFARPUR URBAN 2|MUZAFFARPUR URBAN II|MUZAFFARPUR(U2)|MUZAFFARPUR U-2"),
    ("NBPDCL", "MUZAFFARPUR", "MUZAFFARPUR(W)", 74, "MUZAFFARPUR WEST"),
    ("NBPDCL", "MUZAFFARPUR", "PUPRI", 51, ""),
    ("NBPDCL", "MUZAFFARPUR", "SHEOHAR", 22, ""),
    ("NBPDCL", "MUZAFFARPUR", "SITAMARHI", 54, ""),
    ("NBPDCL", "PURNEA", "BARSOI", 34, ""),
    ("NBPDCL", "PURNEA", "KATIHAR", 67, ""),
    ("NBPDCL", "PURNEA", "PURNEA_WEST", 100, "PURNEA|PURNIA|PURNEA WEST|PURNIA WEST|PURNEA(W)"),
    ("NBPDCL", "PURNEA", "PURNEA_EAST", 60, "PURNEA EAST|PURNIA EAST|PURNEA(E)"),
    ("NBPDCL", "SAHARSA", "MADHEPURA", 50, ""),
    ("NBPDCL", "SAHARSA", "RAGHOPUR", 70, ""),
    ("NBPDCL", "SAHARSA", "SAHARSA", 59, ""),
    ("NBPDCL", "SAHARSA", "SIMRI BAKHTIYARPUR", 33, "SIMRI BAKHTIARPUR"),
    ("NBPDCL", "SAHARSA", "SUPAUL", 51, ""),
    ("NBPDCL", "SAHARSA", "UDAIKISHANGANJ", 34, "UDA KISHANGANJ"),
    ("NBPDCL", "SAMASTIPUR", "DALSINGHSARAI", 42, "DALSINGSARAI"),
    ("NBPDCL", "SAMASTIPUR", "HAJIPUR", 69, ""),
    ("NBPDCL", "SAMASTIPUR", "MAHUA", 58, ""),
    ("NBPDCL", "SAMASTIPUR", "ROSERA", 39, "ROSERA"),
    ("NBPDCL", "SAMASTIPUR", "SAMASTIPUR", 69, ""),
]

# Other names used for a circle in the CSV files.
BUILTIN_CIRCLE_NAMES = {
    "GAYA JI": "GAYA|GAYAJI",
    "BIHARSARIF": "BIHARSHARIF|BIHARSHARIFF|BIHAR SHARIF",
    "ARRAH": "ARA",
    "PESU EAST": "PESU(E)|PESU-EAST",
    "PESU WEST": "PESU(W)|PESU-WEST",
    "CHAPRA": "CHHAPRA",
    "PURNEA": "PURNIA",
}

DISCOM_ORDER = ["SBPDCL", "NBPDCL", "NOT MAPPED"]
STATE = "BIHAR"

OUTAGE_BANDS = [            # (label, upper limit in seconds, inclusive)
    ("Up to 15 min", 15 * 60),
    ("15 to 30 min", 30 * 60),
    ("30 min to 1 hr", 3600),
    ("1 to 2 hrs", 2 * 3600),
    ("2 to 4 hrs", 4 * 3600),
    ("4 to 8 hrs", 8 * 3600),
    ("Over 8 hrs", None),
]
INTERRUPTION_BANDS = [      # (label, upper limit, inclusive)
    ("1", 1),
    ("2 to 3", 3),
    ("4 to 5", 5),
    ("6 to 9", 9),
    ("10 or more", None),
]

# =============================================================================
#  SMALL HELPERS
# =============================================================================


class UserError(Exception):
    """An error with a plain-words message for the user."""


class Cancelled(Exception):
    pass


def name_key(text):
    """ARRAH, Arrah and ' arrah ' all become ARRAH. Brackets and dashes are ignored."""
    return re.sub(r"[^A-Z0-9]", "", str(text or "").upper())


def clean_text(text):
    text = str(text or "")
    quick = text.strip()
    if "  " not in quick and quick.isprintable():        # the usual case: nothing to tidy
        return quick
    text = text.replace("\x00", "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


RE_OFFICE_WORD = re.compile(r"[\s_\-]+(CIRCLE|DIVISION|DIVN\.?|DIV\.?)$", re.I)


def office_name(text):
    """'PURNEA CIRCLE' and 'ARARIA Division' become PURNEA and ARARIA."""
    text = clean_text(text)
    short = RE_OFFICE_WORD.sub("", text).strip()
    return short or text


def clean_substation(text):
    """Tidy a substation name. A 10-digit phone number in the name is removed."""
    text = clean_text(text)
    text = re.sub(r"(?<!\d)[6-9]\d{9}(?!\d)", "", text)
    return re.sub(r"\s+", " ", text).strip(" -_,").upper()


def app_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def fmt_date(ordinal):
    return datetime.date.fromordinal(ordinal).strftime("%d.%m.%Y")


def fmt_hhmm(seconds):
    seconds = int(round(seconds))
    return "%d:%02d" % (seconds // 3600, (seconds % 3600) // 60)


def fmt_size(n):
    for unit in ("bytes", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return ("%d %s" % (n, unit)) if unit == "bytes" else ("%.1f %s" % (n, unit))
        n /= 1024.0


MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}


def parse_date(text):
    """Reads a date written in the common ways. Day comes first (Indian style).
    Returns a date, or None if the text is not a date."""
    s = clean_text(text)
    if not s:
        return None
    s = re.sub(r"[T ]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?\s*([AaPp][Mm])?\s*(Z|[+-]\d{2}:?\d{2})?$", "", s).strip()
    try:
        m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
        if m:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.match(r"^(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{2,4})$", s)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y < 100:
                y += 2000
            if mo > 12 and d <= 12:      # written month first
                d, mo = mo, d
            return datetime.date(y, mo, d)
        m = re.match(r"^(\d{1,2})[-/. ]?([A-Za-z]{3,9})[-/., ]*(\d{2,4})$", s)
        if m and m.group(2)[:3].lower() in MONTHS:
            y = int(m.group(3))
            if y < 100:
                y += 2000
            return datetime.date(y, MONTHS[m.group(2)[:3].lower()], int(m.group(1)))
        m = re.match(r"^([A-Za-z]{3,9})[-/. ]*(\d{1,2})[-/., ]+(\d{4})$", s)
        if m and m.group(1)[:3].lower() in MONTHS:
            return datetime.date(int(m.group(3)), MONTHS[m.group(1)[:3].lower()], int(m.group(2)))
        m = re.match(r"^(20\d{2})(\d{2})(\d{2})$", s)
        if m:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None
    return None


def parse_count(text):
    """Number of interruptions. Returns an int, or None if it is not a whole number."""
    s = clean_text(text).replace(",", "")
    if not s:
        return None
    if re.match(r"^\d+$", s):
        return int(s)
    if re.match(r"^\d+\.0*$", s):
        return int(float(s))
    return None


def parse_duration(text, unit_hint=""):
    """Outage time in seconds. Accepts 05:39:31, 5:39, 91:12:37, '1 day, 2:03:04',
    '2h 15m 10s'. A plain number is used only if the heading names the unit.
    Returns None if it cannot be read."""
    s = clean_text(text).lower()
    if not s:
        return None
    days = 0
    m = re.match(r"^(\d+)\s*days?,?\s*(.*)$", s)
    if m:
        days, s = int(m.group(1)), m.group(2).strip()
        if not s:
            return days * 86400
    m = re.match(r"^(\d+)\s*[:.]\s*(\d{1,2})(?:\s*[:.]\s*(\d{1,2})(?:\.\d+)?)?$", s)
    if not m:
        m = re.match(r"^(\d+)\s+(\d{1,2})(?:\s+(\d{1,2}))?$", s)
    if m and (":" in s or " " in s or unit_hint == "hms"):
        h, mi, se = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        if mi > 60 or se > 60:           # 11:60 is taken as 12:00 (some exports round the minutes up to 60)
            return None
        return days * 86400 + h * 3600 + mi * 60 + se
    m = re.match(r"^(?:(\d+)\s*h(?:rs?|ours?)?)?\s*(?:(\d+)\s*m(?:in(?:ute)?s?)?)?\s*(?:(\d+)\s*s(?:ec(?:ond)?s?)?)?$", s)
    if m and any(m.groups()):
        return days * 86400 + int(m.group(1) or 0) * 3600 + int(m.group(2) or 0) * 60 + int(m.group(3) or 0)
    if re.match(r"^\d+(\.\d+)?$", s):
        v = float(s)
        if unit_hint == "min":
            return int(round(v * 60))
        if unit_hint == "sec":
            return int(round(v))
        if unit_hint == "hour":
            return int(round(v * 3600))
    return None


RE_33KV = re.compile(r"(?<!\d)33\s*K\.?\s*V", re.I)
RE_SIXTY = re.compile(r":60(?!\d)")
RE_VOLT_33 = re.compile(r"^\s*33(\.0+)?\s*(K\.?\s*V\.?)?\s*$", re.I)       # the cell of a Voltage column


def feeder_type_label(raw):
    """Returns (label, group). group is MAIN, AGRI, IND or NAP."""
    k = name_key(raw)
    if not k or k in ("NA", "NIL", "NONE", "NULL"):
        return "Not Given", "MAIN"
    if k.startswith("URB") or k in ("U", "TOWN"):
        return "Urban", "MAIN"
    if k.startswith("RUR") or k == "R":
        return "Rural", "MAIN"
    if k.startswith("MIX") or k == "M":
        return "Mixed", "MAIN"
    if k.startswith("AGR") or k in ("AG", "KRISHI"):
        return "Agriculture", "AGRI"
    if k.startswith("IND"):
        return "Industrial", "IND"
    if k == "NAP":
        return "NAP", "NAP"
    return clean_text(raw).title(), "MAIN"


# =============================================================================
#  FEEDER MASTER  (fixed feeder numbers, DISCOM of each circle, other names)
# =============================================================================


class Master(object):
    def __init__(self):
        self.divisions = []          # dicts: discom, circle, division, feeders
        self.by_pair = {}            # (circle key, division key) -> entry
        self.by_div = {}             # division key -> [entries]
        self.circle = {}             # circle key -> (discom, circle name)
        self.source = "built-in list"

    def add(self, discom, circle, division, feeders, div_names="", circle_names=""):
        discom = name_key(discom) or "NOT MAPPED"
        circle = clean_text(circle).upper()
        division = clean_text(division).upper()
        if not circle or not division:
            return
        entry = {"discom": discom, "circle": circle, "division": division, "feeders": int(feeders)}
        self.divisions.append(entry)
        ckeys = [name_key(circle)] + [name_key(x) for x in str(circle_names or "").split("|")]
        ckeys += [name_key(x) for x in BUILTIN_CIRCLE_NAMES.get(circle, "").split("|")]
        ckeys = [k for k in ckeys if k]
        for ck in ckeys:
            self.circle.setdefault(ck, (discom, circle))
        dkeys = [name_key(division)] + [name_key(x) for x in str(div_names or "").split("|")]
        for dk in [k for k in dkeys if k]:
            for ck in ckeys:
                self.by_pair.setdefault((ck, dk), entry)
            lst = self.by_div.setdefault(dk, [])
            if entry not in lst:
                lst.append(entry)

    def find(self, circle_raw, division_raw):
        """Returns (discom, circle name, division name, master entry or None)."""
        ck, dk = name_key(circle_raw), name_key(division_raw)
        entry = self.by_pair.get((ck, dk))
        if entry is None:
            cands = self.by_div.get(dk, [])
            if len(cands) == 1 and ck not in self.circle:
                entry = cands[0]
        if entry is not None:
            return entry["discom"], entry["circle"], entry["division"], entry
        division = clean_text(division_raw).upper() or "NOT GIVEN"
        if ck in self.circle:
            discom, circle = self.circle[ck]
            return discom, circle, division, None
        return "NOT MAPPED", clean_text(circle_raw).upper() or "NOT GIVEN", division, None


def builtin_master():
    m = Master()
    for discom, circle, division, feeders, names in BUILTIN_MASTER:
        m.add(discom, circle, division, feeders, names)
    return m


MASTER_HEADINGS = ["Discom", "Circle", "Division", "Total Feeders",
                   "Other Names For Division", "Other Names For Circle"]


def write_master_file(path):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(MASTER_HEADINGS)
        for discom, circle, division, feeders, names in BUILTIN_MASTER:
            w.writerow([discom, circle, division, feeders, names, BUILTIN_CIRCLE_NAMES.get(circle, "")])


def load_master(log=None):
    """Uses feeder_master.csv kept next to the app if it is there, else the built-in list."""
    path = os.path.join(app_folder(), MASTER_FILE_NAME)
    if not os.path.isfile(path):
        return builtin_master()
    m = Master()
    try:
        with open(path, "r", newline="", encoding="utf-8-sig", errors="replace") as fh:
            rows = list(csv.reader(fh))
    except OSError as e:
        raise UserError("The feeder list file could not be opened:\n%s\n(%s)\n"
                        "Close it if it is open in Excel, or delete it to use the built-in list." % (path, e))
    if not rows:
        raise UserError("The feeder list file is empty:\n%s\nDelete it to use the built-in list." % path)
    head = [name_key(h) for h in rows[0]]

    def col(*names):
        for n in names:
            if n in head:
                return head.index(n)
        return -1
    ci_d, ci_c, ci_v, ci_n = col("DISCOM"), col("CIRCLE"), col("DIVISION"), col("TOTALFEEDERS", "FEEDERS", "TOTALNOOFFEEDER")
    ci_o, ci_oc = col("OTHERNAMESFORDIVISION", "OTHERNAMES"), col("OTHERNAMESFORCIRCLE")
    if min(ci_d, ci_c, ci_v, ci_n) < 0:
        raise UserError("The feeder list file must have the headings: Discom, Circle, Division, Total Feeders.\n"
                        "File: %s\nCorrect the headings, or delete the file to use the built-in list." % path)
    bad = 0
    for r in rows[1:]:
        r = list(r) + [""] * 8
        if not any(clean_text(x) for x in r):
            continue
        n = parse_count(r[ci_n])
        if n is None:
            bad += 1
            continue
        m.add(r[ci_d], r[ci_c], r[ci_v], n, r[ci_o] if ci_o >= 0 else "", r[ci_oc] if ci_oc >= 0 else "")
    if not m.divisions:
        raise UserError("The feeder list file has no usable rows:\n%s\nDelete it to use the built-in list." % path)
    m.source = "file " + MASTER_FILE_NAME
    if log and bad:
        log("warn", "Feeder list: %d row(s) were skipped because Total Feeders is not a number." % bad)
    return m


# =============================================================================
#  TYPE LIST FILE  (switched off: see USE_TYPE_LIST_FILE. Feeder types come from the company files.)
# =============================================================================

RE_AGRI_NAME = re.compile(r"(?<![A-Za-z])(AG|AGR|AGRI|AGRO|AGRICULTURE|AGRICULTURAL|KRISHI)(?![A-Za-z])", re.I)
TYPE_HEADINGS = ["Circle", "Feeder Code", "Meter No", "Feeder", "Correct Type"]


def correct_type(text):
    """Reads the 'Correct Type' cell. Returns (label, group) or None if it is not a known type."""
    k = name_key(text)
    if k in ("33KV", "33", "NAP"):
        return "33 kV", "NAP"
    label, group = feeder_type_label(text)
    if label in ("Urban", "Rural", "Mixed", "Agriculture", "Industrial"):
        return label, group
    return None


class TypeList(object):
    def __init__(self):
        self.by_code, self.by_msn = {}, {}          # (circle key, code or meter) -> type; circle key "" = any circle
        self.source = "built-in list"
        self.count = 0

    def add(self, code, msn, type_text, circle=""):
        t = correct_type(type_text)
        code, msn, ck = clean_text(code).upper(), clean_text(msn).upper(), name_key(circle)
        if t is None or not (code or msn):
            return False
        if code:
            self.by_code[(ck, code)] = t
        if msn:
            self.by_msn[(ck, msn)] = t
        self.count += 1
        return True

    def find(self, code, msn, circle=""):
        ck = name_key(circle)
        if code:
            code = code.upper()
            t = self.by_code.get((ck, code)) or self.by_code.get(("", code))
            if t:
                return t
        if msn:
            msn = msn.upper()
            return self.by_msn.get((ck, msn)) or self.by_msn.get(("", msn))
        return None


def write_type_file(path):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(TYPE_HEADINGS)
        for row in BUILTIN_TYPE_CORRECTIONS:
            w.writerow(row)


def fetch_agriculture_list(log=None, fetch=True):
    """The shared Agriculture list as rows (circle, code, meter, name, type), and where it came from.
    Order: the list the browser page passed in; Google (and a saved copy is kept); the saved copy; the built-in list."""
    if SHARED_TYPE_ROWS is not None:
        return list(SHARED_TYPE_ROWS), "shared list"
    cache = os.path.join(app_folder(), AGRI_CACHE_NAME)
    if fetch and web_link_ready():
        try:
            req = urllib.request.Request(WEB_APP_URL + "?api=types", headers={"User-Agent": "OutageReportApp/" + APP_VERSION})
            with urllib.request.urlopen(req, timeout=25) as resp:
                ans = json.loads(resp.read().decode("utf-8", errors="replace"))
            if ans.get("ok") and isinstance(ans.get("list"), list):
                rows = [(x.get("circle", ""), x.get("code", ""), x.get("meter", ""), x.get("name", ""), x.get("type") or "Agriculture")
                        for x in ans["list"]]
                try:
                    with open(cache, "w", encoding="utf-8") as fh:
                        json.dump({"saved": datetime.datetime.now().strftime("%d-%b-%Y %H:%M"), "list": ans["list"]}, fh, ensure_ascii=False, indent=1)
                except OSError:
                    pass
                return rows, "shared list"
            if log:
                log("warn", "The shared Agriculture list could not be read from Google (%s). The saved copy is used." % ans.get("message", "no list"))
        except Exception as e:
            if log:
                log("warn", "The shared Agriculture list could not be read from Google (%s). The saved copy is used." % e)
    if os.path.isfile(cache):
        try:
            with open(cache, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return [(x.get("circle", ""), x.get("code", ""), x.get("meter", ""), x.get("name", ""), x.get("type") or "Agriculture")
                    for x in data.get("list", [])], "saved copy of the shared list (%s)" % data.get("saved", "")
        except (OSError, ValueError):
            pass
    return list(BUILTIN_TYPE_CORRECTIONS), "starting list built into the app"


def load_type_list(log=None, master=None, fetch=True):
    """The list of feeders whose type is corrected: the shared Agriculture list (or the old file if USE_TYPE_LIST_FILE is True)."""
    tl = TypeList()

    def circle_name(text):
        """The circle as the feeder list writes it. A circle that is not known means: any circle."""
        k = name_key(text)
        if not k:
            return ""
        if master is None:
            return text
        return master.circle[k][1] if k in master.circle else ""
    path = os.path.join(app_folder(), TYPE_FILE_NAME)
    if not USE_TYPE_LIST_FILE:
        rows, source = fetch_agriculture_list(log, fetch)
        for circle, code, msn, _, typ in rows:
            tl.add(code, msn, typ, circle_name(circle))
        tl.source = source
        return tl
    if not os.path.isfile(path):
        for circle, code, msn, _, typ in BUILTIN_TYPE_CORRECTIONS:
            tl.add(code, msn, typ, circle_name(circle))
        return tl
    try:
        with open(path, "r", newline="", encoding="utf-8-sig", errors="replace") as fh:
            rows = list(csv.reader(fh))
    except OSError as e:
        raise UserError("The type list file could not be opened:\n%s\n(%s)\n"
                        "Close it if it is open in Excel, or delete it to use the built-in list." % (path, e))
    head = [name_key(h) for h in (rows[0] if rows else [])]

    def col(*names):
        for n in names:
            if n in head:
                return head.index(n)
        return -1
    ci_c, ci_m, ci_t = col("FEEDERCODE", "FEEDERID"), col("METERNO", "MSN", "METERNUMBER"), col("CORRECTTYPE", "TYPE", "FEEDERTYPE")
    ci_o = col("CIRCLE")
    if ci_t < 0 or (ci_c < 0 and ci_m < 0):
        raise UserError("The type list file must have the headings: Feeder Code, Meter No, Correct Type.\n"
                        "File: %s\nCorrect the headings, or delete the file to use the built-in list." % path)
    bad = 0
    for r in rows[1:]:
        r = list(r) + [""] * 8
        if not any(clean_text(x) for x in r):
            continue
        if not tl.add(r[ci_c] if ci_c >= 0 else "", r[ci_m] if ci_m >= 0 else "", r[ci_t], circle_name(r[ci_o]) if ci_o >= 0 else ""):
            bad += 1
    tl.source = "file " + TYPE_FILE_NAME
    if log and bad:
        log("warn", "Type list: %d row(s) were skipped. Each row needs a Feeder Code or a Meter No, and Correct Type must be one of "
                    "Agriculture, Industrial, Urban, Rural, Mixed, 33 kV." % bad)
    return tl


# =============================================================================
#  READING THE CSV FILES  (row by row, never the whole file in memory)
# =============================================================================

FIELDS_NEEDED = ["circle", "division", "feeder", "events", "duration"]
FIELD_WORDS = {"circle": "Circle", "division": "Division", "feeder": "Feeder",
               "events": "Event Count", "duration": "Duration"}


def map_headings(row):
    """Finds which column holds what, by reading the headings. Returns {field: index}."""
    found = {}

    def take(field, i):
        if field not in found:
            found[field] = i
            return True
        return False

    for i, cell in enumerate(row):
        n = re.sub(r"[^a-z0-9]", "", str(cell or "").lower())
        if not n:
            continue
        has_ss = ("substation" in n) or ("pss" in n)
        if n in ("date", "reportdate", "outagedate", "eventdate", "day", "datadate", "readingdate", "fordate"):
            take("date", i)
        elif "voltage" in n or n in ("kv", "kvlevel"):
            take("voltage", i)
        elif "subdiv" in n:
            take("subdivision", i)
        elif "division" in n or n == "div":
            take("division", i)
        elif "circle" in n:
            take("circle", i)
        elif has_ss and ("code" in n or n.endswith("id")):
            take("sscode", i)
        elif has_ss:
            take("substation", i)
        elif "feeder" in n and ("code" in n or n.endswith("id")):
            take("fcode", i)
        elif ("feeder" in n and ("type" in n or "category" in n)) or n in ("type", "category"):
            take("ftype", i)
        elif "msn" in n or ("meter" in n and ("no" in n or "serial" in n or "number" in n)):
            take("msn", i)
        elif "duration" in n or "outagetime" in n or ("outage" in n and ("hh" in n or "hrs" in n or "min" in n)):
            if take("duration", i):
                if "hh" in n:
                    found["_unit"] = "hms"
                elif "min" in n:
                    found["_unit"] = "min"
                elif "sec" in n:
                    found["_unit"] = "sec"
                elif "hr" in n or "hour" in n:
                    found["_unit"] = "hour"
        elif "event" in n or "interruption" in n or "tripping" in n or n in ("count", "nos") \
                or n.endswith("count") or "noofoutage" in n:
            take("events", i)
        elif "feeder" in n:
            take("feeder", i)
    return found


def is_heading_row(row):
    found = map_headings(row)
    return all(f in found for f in FIELDS_NEEDED), found


class LineSource(object):
    """Gives the text lines of a file one at a time, and knows how far it has read.
    Copes with UTF-8, UTF-8 with BOM, UTF-16 and old Windows text."""

    def __init__(self, path):
        self.path = path
        self.size = max(1, os.path.getsize(path))
        self.fb = open(path, "rb")
        start = self.fb.read(4)
        self.fb.seek(0)
        self.utf16 = start[:2] in (b"\xff\xfe", b"\xfe\xff")
        self.encoding_used = "UTF-16" if self.utf16 else "UTF-8"
        self.text = io.TextIOWrapper(self.fb, encoding="utf-16", errors="replace", newline="") if self.utf16 else None
        self.first = True

    def position(self):
        try:
            return self.fb.tell()
        except (OSError, ValueError):
            return 0

    def lines(self):
        if self.utf16:
            while True:
                line = self.text.readline()
                if not line:
                    return
                yield line.replace("\x00", "")
        else:
            while True:
                raw = self.fb.readline()
                if not raw:
                    return
                if self.first:
                    self.first = False
                    if raw.startswith(b"\xef\xbb\xbf"):
                        raw = raw[3:]
                try:
                    line = raw.decode("utf-8")
                except UnicodeDecodeError:
                    self.encoding_used = "Windows text"
                    line = raw.decode("cp1252", errors="replace")
                yield line.replace("\x00", "")

    def close(self):
        try:
            if self.text is not None:
                self.text.detach()
        except Exception:
            pass
        try:
            self.fb.close()
        except Exception:
            pass


def _row_reads(row, cols):
    """True if the row gives a circle, a count and a time with this column map."""
    def cell(f):
        i = cols.get(f, -1)
        return row[i] if 0 <= i < len(row) else ""
    return bool(clean_text(cell("circle"))) and parse_count(cell("events")) is not None \
        and parse_duration(cell("duration"), cols.get("_unit", "")) is not None


def _share_read(rows, cols):
    return (sum(1 for r in rows if _row_reads(r, cols)) / float(len(rows))) if rows else 0.0


def _shift_for(head_row, found, data_rows):
    """Some report-style exports put the figures one or two columns away from their headings.
    If so, returns the map {heading column: data column}; else None."""
    if len(data_rows) < 1 or _share_read(data_rows, found) >= 0.6:
        return None
    head_pos = [i for i, c in enumerate(head_row) if str(c or "").strip()]
    need = max(1, int(0.02 * len(data_rows)))
    filled = {}
    for r in data_rows:
        for i, c in enumerate(r):
            if c.strip():
                filled[i] = filled.get(i, 0) + 1
    data_pos = sorted(i for i, n in filled.items() if n >= need)
    if len(data_pos) != len(head_pos) or data_pos == head_pos:
        return None
    shift = dict(zip(head_pos, data_pos))
    moved = dict((f, (shift.get(i, i) if f != "_unit" else i)) for f, i in found.items())
    if _share_read(data_rows, moved) >= 0.6 and _share_read(data_rows, moved) > _share_read(data_rows, found):
        return shift
    return None


def _no_heading_layout(rows):
    """Checks the rows against the known files that have no heading row. Returns (name, column map) or None."""
    for lay in NO_HEADING_LAYOUTS:
        cols = dict((f, int(n) - 1) for f, n in lay.items() if f not in ("name", "columns"))
        good = [r for r in rows if len(r) == lay["columns"]]
        if len(good) < max(1, int(0.9 * len(rows))):
            continue
        fit = 0
        for r in good:
            okay = _row_reads(r, cols) and ":" in r[cols["duration"]]
            if okay and "date" in cols:
                okay = parse_date(r[cols["date"]]) is not None
            if okay and "voltage" in cols:
                okay = re.match(r"^\s*\d+(\.\d+)?\s*K\.?\s*V\.?\s*$", r[cols["voltage"]], re.I) is not None
            fit += 1 if okay else 0
        if fit >= max(1, int(0.9 * len(good))):
            return lay["name"], cols
    return None


def find_layout(path):
    """Looks at the top of the file to find the separator and the heading row.
    Returns a dict: sep, notes (text above the heading), fixed (column map of a file with no
    heading row, or None), fixed_name, shift (see _shift_for, or None)."""
    with open(path, "rb") as fb:
        start = fb.read(8)
    if start[:4] == b"PK\x03\x04" or start[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise UserError("This is an Excel file, not a CSV file:\n%s\n"
                        "Open it in Excel, choose File > Save As > CSV (Comma delimited), and add the saved CSV file instead." % path)
    src = LineSource(path)
    try:
        sample = []
        for line in src.lines():
            sample.append(line)
            if len(sample) >= 300:
                break
    finally:
        src.close()
    if not any(s.strip() for s in sample):
        raise UserError("This file is empty:\n%s\nRemove it from the list, or get the file again." % path)
    parsed = {}
    for sep in (",", ";", "\t", "|"):
        try:
            parsed[sep] = list(csv.reader(sample, delimiter=sep))
        except csv.Error:
            continue
    for sep in (",", ";", "\t", "|"):
        rows = parsed.get(sep, [])
        for i, row in enumerate(rows):
            ok, found = is_heading_row(row)
            if ok:
                notes = " ".join(" ".join(r) for r in rows[:i])
                head_key = [name_key(x) for x in row]
                data = [r for r in rows[i + 1:] if sum(1 for x in r if x.strip()) > 2 and [name_key(x) for x in r] != head_key]
                return {"sep": sep, "notes": notes, "fixed": None, "fixed_name": "", "shift": _shift_for(row, found, data)}
    for sep in (",", ";", "\t", "|"):                    # no heading row: is it one of the known files?
        rows = [r for r in parsed.get(sep, []) if sum(1 for x in r if x.strip()) > 2]
        got = _no_heading_layout(rows) if rows else None
        if got:
            return {"sep": sep, "notes": "", "fixed": got[1], "fixed_name": got[0], "shift": None}
    # say which headings are missing, using the best row found
    best = set()
    for sep in (",", ";", "\t", "|"):
        for row in parsed.get(sep, [])[:60]:
            got = set(f for f in map_headings(row) if f in FIELDS_NEEDED)
            if len(got) > len(best):
                best = got
    missing = [FIELD_WORDS[f] for f in FIELDS_NEEDED if f not in best]
    raise UserError(
        "The heading row was not found in this file:\n%s\n"
        "Headings not found: %s.\n"
        "The file must have a row of headings with Circle, Division, Feeder, Event Count and Duration. "
        "Check that this is a feeder outage CSV file. If it is an Excel file, save it as CSV first.\n"
        "(A file with no heading row is read only if its columns are in the order the app knows: %s.)"
        % (path, ", ".join(missing) if missing else "none clearly",
           ", ".join("%s file with %d columns" % (l["name"], l["columns"]) for l in NO_HEADING_LAYOUTS) or "none"))


def dates_inside(path, limit=400):
    """Looks at the first rows of a file and returns the dates written in its Date column, as {date: rows}."""
    out = {}
    try:
        lay = find_layout(path)
        src = LineSource(path)
    except (UserError, OSError):
        return out
    try:
        cols = dict(lay["fixed"]) if lay["fixed"] else None
        n = 0
        for row in csv.reader(src.lines(), delimiter=lay["sep"]):
            n += 1
            if n > limit:
                break
            if cols is None:
                ok, found = is_heading_row(row)
                if ok:
                    cols = dict((f, (lay["shift"].get(i, i) if lay["shift"] and f != "_unit" else i)) for f, i in found.items())
                continue
            i = cols.get("date", -1)
            if 0 <= i < len(row) and _row_reads(row, cols):
                d = parse_date(row[i])
                if d:
                    out[d] = out.get(d, 0) + 1
    except (csv.Error, OSError):
        pass
    finally:
        src.close()
    return out


RE_STAMP = re.compile(r"(?<!\d)(20\d{2})(\d{2})(\d{2})[_\-T ]?(\d{6})(?!\d)")
RE_DMY = re.compile(r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})[.\-_](20\d{2}|\d{2})(?!\d)")
RE_YMD = re.compile(r"(?<!\d)(20\d{2})[.\-_]?(\d{2})[.\-_]?(\d{2})(?!\d)")
RE_DMONY = re.compile(r"(?<!\d)(\d{1,2})[.\-_ ]?([A-Za-z]{3,9})[.\-_ ]?(20\d{2})(?:[_\-T ](\d{1,2})[\-.:h](\d{2}))?")
RE_DDMMYY = re.compile(r"(?<!\d)(\d{2})(\d{2})(\d{2})(?!\d)")


def guess_date_from_names(paths):
    """A file name with a download stamp (20261011_071623, or 11-Oct-2026_14-30) means the data
    is of the day before. A plain date in the name (10.10.2026, 101026) is taken as the data date."""
    def one(name):
        def stamp(m):
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) - datetime.timedelta(days=1)

        def dmony(m):
            if m.group(2)[:3].lower() not in MONTHS:
                raise ValueError
            d = datetime.date(int(m.group(3)), MONTHS[m.group(2)[:3].lower()], int(m.group(1)))
            return d - datetime.timedelta(days=1) if m.group(4) else d

        def dmy(m):
            y = int(m.group(3))
            return datetime.date(y + 2000 if y < 100 else y, int(m.group(2)), int(m.group(1)))

        def ymd(m):
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

        def ddmmyy(m):                    # 101026 : only if it is a day of the last two months
            d = datetime.date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1)))
            if not (0 <= (datetime.date.today() - d).days <= 62):
                raise ValueError
            return d
        for pattern, make in ((RE_STAMP, stamp), (RE_DMONY, dmony), (RE_DMY, dmy), (RE_YMD, ymd), (RE_DDMMYY, ddmmyy)):
            for m in pattern.finditer(name):
                try:
                    return make(m)
                except ValueError:
                    continue
        return None
    best = None
    for p in paths:
        d = one(os.path.basename(p))
        if d and (best is None or d > best):
            best = d
    return best


def guess_date(paths):
    """The report date to offer: the date written inside the files if they carry one, else the date in the file names."""
    inside = {}
    for p in paths:
        for d, n in dates_inside(p).items():
            inside[d] = inside.get(d, 0) + n
    if inside:
        return max(inside, key=lambda d: (inside[d], d))
    return guess_date_from_names(paths)


class Store(object):
    """Keeps one line per feeder per day in a temporary file on disk (not in memory)."""

    def __init__(self):
        fd, self.path = tempfile.mkstemp(prefix="outage_", suffix=".tmp")
        os.close(fd)
        self.db = sqlite3.connect(self.path)
        self.db.execute("PRAGMA journal_mode=OFF")
        self.db.execute("PRAGMA synchronous=OFF")
        self.db.execute(
            "CREATE TABLE f (k TEXT PRIMARY KEY, d INTEGER, discom TEXT, circle TEXT, division TEXT,"
            " mapped INTEGER, sskey TEXT, sscode TEXT, ssname TEXT, fname TEXT, ftype TEXT, msn TEXT,"
            " ev INTEGER, secs INTEGER, cat TEXT, over24 INTEGER, otype TEXT, fixed INTEGER, src INTEGER, fcode TEXT) WITHOUT ROWID")

    def close(self):
        try:
            self.db.close()
        except Exception:
            pass
        try:
            os.remove(self.path)
        except OSError:
            pass


def as_inputs(items):
    """Each input is (DISCOM part of the window, company, file path). A plain path has no DISCOM part."""
    out = []
    for it in items:
        if isinstance(it, (tuple, list)):
            out.append((str(it[0] or ""), str(it[1] or "Not stated"), it[2]))
        else:
            out.append(("", "Not stated", it))
    return out


def read_files(inputs, report_date, master, types, store, log, progress, cancel):
    """Reads every CSV file row by row into the store. Returns the counts."""
    inputs = as_inputs(inputs)
    paths = [x[2] for x in inputs]
    c = {"files": 0, "rows": 0, "blank": 0, "notes": 0, "broken": 0, "zero": 0,
         "dup_same": 0, "dup_diff": 0, "kept": 0, "type_unknown": 0, "own_date": 0, "no_date": 0, "sixty": 0}
    dated = {}                                  # day -> the files that carry this date in their rows
    undated = []                                # files with rows that carry no date
    slots, slot_of = [], {}                     # one slot for each (DISCOM part, company)
    for section, company, _ in inputs:
        if (section, company) not in slot_of:
            slot_of[(section, company)] = len(slots)
            slots.append({"section": section, "company": company, "files": 0, "rows": 0, "broken": 0})
    for p in paths:
        if not os.path.isfile(p):
            raise UserError("This file is no longer there:\n%s\nTake it off with the X next to it or put the file back, then press the button again." % p)
    total_bytes = sum(max(1, os.path.getsize(p)) for p in paths)
    done_bytes = 0
    db = store.db
    default_day = report_date.toordinal()
    unknown_circles, unknown_divs = {}, {}
    found_before, type_before = {}, {}          # remembered answers, so each name is worked out once
    volt_before = {}

    for file_no, (section, company, path) in enumerate(inputs, 1):
        short = os.path.basename(path)
        slot_no = slot_of[(section, company)]
        if section:
            short = "%s, %s: %s" % (section, company, short)
        seen_discom = {}
        if cancel.is_set():
            raise Cancelled()
        try:
            lay = find_layout(path)
            sep, notes = lay["sep"], lay["notes"]
        except PermissionError:
            raise UserError("This file could not be opened:\n%s\nClose it if it is open in another program, then try again." % path)
        except OSError as e:
            raise UserError("This file could not be read:\n%s\n(%s)\nCheck that the drive is connected and the file is not damaged." % (path, e))
        for m in re.finditer(r"\d{1,2}[-/. ](?:\d{1,2}|[A-Za-z]{3,9})[-/., ]+\d{2,4}|20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}", notes or ""):
            nd = parse_date(m.group(0))
            if nd and nd != report_date:
                log("warn", "File %d (%s): its top lines mention the date %s. The report date in use is %s. "
                            "Change the report date if that is wrong." % (file_no, short, nd.strftime("%d.%m.%Y"), report_date.strftime("%d.%m.%Y")))
                break

        src = LineSource(path)
        f_rows = f_kept = f_broken = f_dup = 0
        f_days, f_no_date = set(), 0
        cols = None
        head_key = None
        shown_broken = 0
        if lay["fixed"]:
            cols = dict(lay["fixed"])
            log("info", "File %d (%s) has no heading row. It is read as the %s file: the app knows the order of its columns."
                % (file_no, short, lay["fixed_name"]))
        elif lay["shift"]:
            log("info", "File %d (%s): the figures do not sit under their headings (a report-style export). "
                        "The app lined them up by their order." % (file_no, short))
        try:
            reader = csv.reader(src.lines(), delimiter=sep)
            line_no = 0
            while True:
                try:
                    row = next(reader)
                except StopIteration:
                    break
                except csv.Error:
                    c["rows"] += 1
                    c["broken"] += 1
                    f_broken += 1
                    continue
                line_no = reader.line_num
                c["rows"] += 1
                f_rows += 1
                if f_rows % 2000 == 0:
                    if cancel.is_set():
                        raise Cancelled()
                    progress((done_bytes + src.position()) / float(total_bytes),
                             "Reading file %d of %d: %s" % (file_no, len(paths), short))
                if not row or not any(cell.strip() for cell in row):
                    c["blank"] += 1
                    continue
                if cols is None:
                    ok, found = is_heading_row(row)
                    if ok:
                        cols = found
                        if lay["shift"]:
                            cols = dict((f, (lay["shift"].get(i, i) if f != "_unit" else i)) for f, i in found.items())
                        head_key = [name_key(x) for x in row]
                    c["notes"] += 1
                    continue
                if head_key and name_key(row[0]) == head_key[0] and [name_key(x) for x in row] == head_key:   # heading repeated in the middle
                    c["notes"] += 1
                    continue

                def cell(field):
                    i = cols.get(field, -1)
                    return row[i] if 0 <= i < len(row) else ""

                circle_raw, division_raw = clean_text(cell("circle")), clean_text(cell("division"))
                first = ""
                for x in row:
                    if x.strip():
                        first = x.strip().lower()
                        break
                filled = sum(1 for x in row if x.strip())
                if first.startswith(("total", "grand total", "sub total", "subtotal")) or \
                        circle_raw.lower().endswith(" total") or filled <= 2:
                    c["notes"] += 1
                    continue

                ev = parse_count(cell("events"))
                dur_raw = cell("duration")
                secs = parse_duration(dur_raw, cols.get("_unit", ""))
                if secs is not None and ":60" in dur_raw and RE_SIXTY.search(dur_raw):
                    c["sixty"] += 1
                fname = clean_text(cell("feeder"))
                fcode = clean_text(cell("fcode"))
                msn = clean_text(cell("msn"))
                why = None
                if not circle_raw or not division_raw:
                    why = "Circle or Division is empty"
                elif ev is None:
                    why = "Event Count '%s' is not a whole number" % clean_text(cell("events"))[:20]
                elif secs is None:
                    why = "Duration '%s' could not be read as time" % clean_text(cell("duration"))[:20]
                elif not (fname or fcode or msn):
                    why = "Feeder name, code and meter number are all empty"
                day = default_day
                if why is None and "date" in cols and clean_text(cell("date")):
                    dt = parse_date(cell("date"))
                    if dt is None:
                        why = "Date '%s' could not be read" % clean_text(cell("date"))[:20]
                    else:
                        day = dt.toordinal()
                        c["own_date"] += 1
                        f_days.add(day)
                elif why is None:
                    f_no_date += 1
                if why:
                    c["broken"] += 1
                    f_broken += 1
                    if shown_broken < 5:
                        shown_broken += 1
                        log("warn", "File %d, line %d left out: %s." % (file_no, line_no, why))
                    continue
                if ev == 0 and secs == 0:
                    c["zero"] += 1
                    continue

                pair = (circle_raw, division_raw)
                if pair not in found_before:
                    found_before[pair] = master.find(office_name(circle_raw), office_name(division_raw))
                discom, circle, division, entry = found_before[pair]
                if discom == "NOT MAPPED":
                    unknown_circles[circle] = unknown_circles.get(circle, 0) + 1
                elif entry is None:
                    unknown_divs[(circle, division)] = unknown_divs.get((circle, division), 0) + 1
                type_raw = cell("ftype")
                if type_raw not in type_before:
                    type_before[type_raw] = feeder_type_label(type_raw)
                label, group = type_before[type_raw]
                type_in_file, fixed = label, 0
                fix = types.find(fcode, msn, circle)
                if fix is not None and fix != (label, group):
                    label, group, fixed = fix[0], fix[1], 1          # the type list wins over the file
                volt_raw = cell("voltage")
                if volt_raw not in volt_before:
                    volt_before[volt_raw] = RE_VOLT_33.match(volt_raw) is not None
                is_33 = volt_before[volt_raw]
                left_out = TYPES_LEFT_OUT.get(discom, TYPES_LEFT_OUT_OTHER)
                if group == "NAP" or (not fixed and (is_33 or RE_33KV.search(fname))):
                    cat = "V33"
                elif group == "AGRI" and "Agriculture" in left_out:
                    cat = "AGRI"
                elif group == "IND" and "Industrial" in left_out:
                    cat = "IND"
                else:
                    cat = "MAIN"
                    if label not in ("Urban", "Rural", "Mixed", "Industrial", "Agriculture"):
                        c["type_unknown"] += 1
                sscode = clean_text(cell("sscode")).upper()
                ssname = clean_substation(cell("substation")) or sscode or "NOT GIVEN"
                sskey = sscode or name_key(ssname)
                if msn:
                    key = "%d|%s|%s|%s" % (day, name_key(circle), fcode.upper(), msn.upper())
                elif fcode:          # no meter number: a code alone is not safe, some files use one code for two feeders
                    key = "%d|%s|%s||%s|%s|%s|%s|%s" % (day, name_key(circle), fcode.upper(), name_key(division), sskey,
                                                        name_key(fname), type_in_file, "33" if cat == "V33" else "11")
                else:
                    key = "%d|%s|%s|%s|%s" % (day, name_key(circle), name_key(division), sskey, name_key(fname))
                over24 = 1 if secs > MAX_SECONDS_IN_DAY else 0
                rec = (key, day, discom, circle, division, 1 if entry is not None else 0, sskey, sscode, ssname,
                       fname or fcode or msn, label, msn, ev, secs, cat, over24, type_in_file, fixed, slot_no, fcode.upper())
                seen_discom[discom] = seen_discom.get(discom, 0) + 1
                cur = db.execute("INSERT OR IGNORE INTO f VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rec)
                if cur.rowcount == 1:
                    c["kept"] += 1
                    f_kept += 1
                else:
                    f_dup += 1
                    old = db.execute("SELECT ev, secs FROM f WHERE k=?", (key,)).fetchone()
                    if old == (ev, secs):
                        c["dup_same"] += 1
                    else:
                        c["dup_diff"] += 1
                        if (secs, ev) < (old[1], old[0]):           # same feeder twice: the lower outage is kept
                            db.execute("DELETE FROM f WHERE k=?", (key,))
                            db.execute("INSERT INTO f VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rec)
        except PermissionError:
            raise UserError("This file could not be opened:\n%s\nClose it if it is open in another program, then try again." % path)
        except OSError as e:
            raise UserError("This file could not be read to the end:\n%s\n(%s)\nCheck the drive and try again." % (path, e))
        finally:
            src.close()
        db.commit()
        done_bytes += max(1, os.path.getsize(path))
        c["files"] += 1
        c["no_date"] += f_no_date
        for dd in f_days:
            dated.setdefault(dd, []).append(os.path.basename(path))
        if f_no_date:
            undated.append(os.path.basename(path))
        slots[slot_no]["files"] += 1
        slots[slot_no]["rows"] += f_rows
        slots[slot_no]["broken"] += f_broken
        for other in sorted(seen_discom):
            if section and other in ("SBPDCL", "NBPDCL") and other != section:
                log("warn", "File %d (%s) was added under %s, but %s of its rows belong to %s circles. They are counted under %s. "
                            "Check that the file is in the right part of the window."
                    % (file_no, short, section, "{:,}".format(seen_discom[other]), other, other))
        msg = "File %d (%s): %s rows read, %s feeder rows kept" % (file_no, short, "{:,}".format(f_rows), "{:,}".format(f_kept))
        if f_dup:
            msg += ", %s repeat rows" % "{:,}".format(f_dup)
        if f_broken:
            msg += ", %s broken rows left out" % "{:,}".format(f_broken)
        log("info", msg + ". Text type: %s." % src.encoding_used)
        if f_kept == 0 and f_dup == 0:
            log("warn", "File %d (%s) gave no usable feeder rows. Check that it is the right file." % (file_no, short))
        progress(done_bytes / float(total_bytes), "Read file %d of %d" % (file_no, len(paths)))

    if undated and dated and default_day not in dated:
        best = max(dated, key=lambda dd: len(dated[dd]))
        raise UserError(
            "The report date is set to %s, but the rows of %s are of %s.\n"
            "The other file(s) (%s) carry no date in their rows, so the app would put them on %s and the report would "
            "show two days mixed up. No report was made.\n"
            "Set the report date to %s and press the button again. "
            "If the files really are of different days, make the report of each day in its own run."
            % (fmt_date(default_day), ", ".join(dated[best][:3]), fmt_date(best), ", ".join(undated[:3]),
               fmt_date(default_day), fmt_date(best)))
    c["unknown_circles"] = unknown_circles
    c["unknown_divs"] = unknown_divs
    c["slots"] = slots
    return c


# =============================================================================
#  MAKING THE TABLES
# =============================================================================


class Table(object):
    """One sheet: a heading row and data rows.
    cols: list of (heading, kind, formula or None).  kind: text int num1 num2 dur date.
    A formula uses [Heading] to point at another column of the same row."""

    def __init__(self, name, cols):
        self.name = name
        self.cols = [(c + (None,))[:3] for c in cols]
        self.rows = []

    def heads(self):
        return [c[0] for c in self.cols]

    def col(self, heading):
        return self.heads().index(heading)

    def add(self, **kw):
        raise NotImplementedError


def safe_div(a, b):
    return (float(a) / b) if b else 0.0


# The measure block used by Summary, Circle and Division (and Day Wise).
F_DEN = "([Total Feeders]*[Days])"
MEASURE_COLS = [
    ("Total Feeders", "int"),
    ("Feeders With Outage", "int"),
    ("Feeders With Outage %", "num1", "IF(%s=0,0,100*[Feeders With Outage]/%s)" % (F_DEN, F_DEN)),
    ("Interruptions", "int"),
    ("Outage Duration (HH:MM)", "dur"),
    ("Outage Min", "num2"),
    ("Avg Interruption per Feeder", "num2", "IF(%s=0,0,[Interruptions]/%s)" % (F_DEN, F_DEN)),
    ("Avg Outage per Feeder (HH:MM)", "dur", "IF(%s=0,0,[Outage Duration (HH:MM)]/%s)" % (F_DEN, F_DEN)),
    ("Avg Outage Min per Feeder", "num2", "IF(%s=0,0,[Outage Duration (HH:MM)]*1440/%s)" % (F_DEN, F_DEN)),
    ("Min per Interruption", "num2", "IF([Interruptions]=0,0,[Outage Duration (HH:MM)]*1440/[Interruptions])"),
    ("Feeders Over %d Hrs" % OVER_HOURS, "int"),
    ("Feeders %d Plus Interruptions" % MANY_INTERRUPTIONS, "int"),
    ("Longest Outage (HH:MM)", "dur"),
    ("Most Interruptions", "int"),
]


def measure_values(total, days, m):
    """m = (n, ev, secs, over, many, maxsecs, maxev). Returns the values of MEASURE_COLS."""
    n, ev, secs, over, many, maxsecs, maxev = m
    den = total * days
    return [total, n, 100.0 * safe_div(n, den), ev, secs, round(secs / 60.0, 2),
            safe_div(ev, den), safe_div(secs, den), safe_div(secs / 60.0, den), safe_div(secs / 60.0, ev),
            over, many, maxsecs, maxev]


AGG = ("COUNT(*), COALESCE(SUM(ev),0), COALESCE(SUM(secs),0), COALESCE(SUM(secs>%d),0),"
       " COALESCE(SUM(ev>=%d),0), COALESCE(MAX(secs),0), COALESCE(MAX(ev),0)" % (OVER_HOURS * 3600, MANY_INTERRUPTIONS))
ZERO = (0, 0, 0, 0, 0, 0, 0)


def add_m(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2], a[3] + b[3], a[4] + b[4], max(a[5], b[5]), max(a[6], b[6]))


def discom_rank(d):
    return DISCOM_ORDER.index(d) if d in DISCOM_ORDER else len(DISCOM_ORDER)


def rank_row_order(scopes):
    """The order of the rows in the sheet Feeder Ranks."""
    place = dict((s, i) for i, s in enumerate(scopes))

    def key(r):
        div = r[0] == "DIVISION"
        return (place.get(r[0], 99), discom_rank(r[4]) if div else 0, (r[6], r[7]) if div else ("", ""), r[1], r[2] != "Top", r[3])
    return key


def build_tables(store, master, counts, log, progress, cancel):
    db = store.db
    q = lambda sql, args=(): db.execute(sql, args).fetchall()

    def step(frac, text):
        if cancel.is_set():
            raise Cancelled()
        progress(frac, text)

    step(0.0, "Working out the figures")
    n_all = q("SELECT COUNT(*) FROM f")[0][0]
    if n_all == 0:
        raise UserError("No feeder rows were found in the files. No report was made.\n"
                        "Check that the files are feeder outage CSV files and are not empty.")
    d_min, d_max, days = q("SELECT MIN(d), MAX(d), COUNT(DISTINCT d) FROM f")[0]
    ok = "over24=0"
    main = "over24=0 AND cat='MAIN'"

    # ---- divisions ---------------------------------------------------------
    div_m = {}
    for r in q("SELECT discom, circle, division, MAX(mapped), %s FROM f WHERE %s GROUP BY discom, circle, division" % (AGG, main)):
        div_m[(r[0], r[1], r[2])] = (r[3], tuple(r[4:]))
    circles_with_rows = set((r[0], r[1]) for r in q("SELECT DISTINCT discom, circle FROM f"))
    divs_with_rows = set((r[0], r[1], r[2]) for r in q("SELECT DISTINCT discom, circle, division FROM f"))
    discoms_with_rows = set(d for d, _ in circles_with_rows)

    div_rows = []        # dict per division
    seen = set()
    for e in master.divisions:
        if e["discom"] not in discoms_with_rows:
            continue                                     # whole DISCOM not in these files
        key = (e["discom"], e["circle"], e["division"])
        if key in seen:
            continue
        seen.add(key)
        has_rows = key in divs_with_rows                 # False = no row at all came for the division
        cir_rows = (e["discom"], e["circle"]) in circles_with_rows
        keep = not LEAVE_OUT_DIVISIONS_WITH_NO_ROWS
        div_rows.append({"discom": e["discom"], "circle": e["circle"], "division": e["division"],
                         "total": e["feeders"], "src": "Feeder list", "received": has_rows or keep,
                         "m": div_m.get(key, (1, ZERO))[1],
                         "circle_received": cir_rows or keep, "has_rows": has_rows, "cir_rows": cir_rows})
    for key in sorted(divs_with_rows):
        if key not in seen:                              # not in the feeder list
            seen.add(key)
            m = div_m.get(key, (0, ZERO))[1]
            div_rows.append({"discom": key[0], "circle": key[1], "division": key[2], "total": m[0],
                             "src": "File (not in feeder list)", "received": True, "m": m, "circle_received": True,
                             "has_rows": True, "cir_rows": True})
    div_rows.sort(key=lambda r: (discom_rank(r["discom"]), r["circle"], r["division"]))

    t_div = Table("Division", [("Discom", "text"), ("Circle", "text"), ("Division", "text"), ("Days", "int")]
                  + MEASURE_COLS + [("Feeder Count From", "text"), ("Data Received", "text"), ("Big Enough To Rank", "text")])
    too_many = []
    for r in div_rows:
        big = r["received"] and r["total"] >= MIN_FEEDERS_DIVISION
        t_div.rows.append([r["discom"], r["circle"], r["division"], days] + measure_values(r["total"], days, r["m"])
                          + [r["src"], "Yes" if r["received"] else "No", "Yes" if big else "No"])
        if r["src"] == "Feeder list" and r["m"][0] > r["total"] * days:
            too_many.append(r)

    # ---- circles -----------------------------------------------------------
    step(0.1, "Working out the figures")
    t_cir = Table("Circle", [("Discom", "text"), ("Circle", "text"), ("Days", "int"), ("Divisions", "int")]
                  + MEASURE_COLS + [("Data Received", "text")])
    cir = {}
    for r in div_rows:
        k = (r["discom"], r["circle"])
        c = cir.setdefault(k, {"n": 0, "total": 0, "m": ZERO, "received": r["circle_received"], "listed": 0})
        c["listed"] += r["total"]
        if not r["received"]:
            continue                                     # a division with no data is not in the circle figures
        c["n"] += 1
        c["total"] += r["total"]
        c["m"] = add_m(c["m"], r["m"])
    for k in sorted(cir, key=lambda k: (discom_rank(k[0]), k[1])):
        c = cir[k]
        t_cir.rows.append([k[0], k[1], days, c["n"]] + measure_values(c["total"] if c["received"] else c["listed"], days, c["m"])
                          + ["Yes" if c["received"] else "No"])

    # ---- substations -------------------------------------------------------
    step(0.2, "Working out the figures")
    t_ss = Table("Substation", [
        ("Discom", "text"), ("Circle", "text"), ("Division", "text"), ("Substation Code", "text"), ("Substation", "text"),
        ("Feeders With Outage", "int"), ("Interruptions", "int"), ("Outage Duration (HH:MM)", "dur"), ("Outage Min", "num2"),
        ("Avg Interruption per Feeder With Outage", "num2", "IF([Feeders With Outage]=0,0,[Interruptions]/[Feeders With Outage])"),
        ("Avg Outage per Feeder With Outage (HH:MM)", "dur", "IF([Feeders With Outage]=0,0,[Outage Duration (HH:MM)]/[Feeders With Outage])"),
        ("Avg Outage Min per Feeder With Outage", "num2", "IF([Feeders With Outage]=0,0,[Outage Duration (HH:MM)]*1440/[Feeders With Outage])"),
        ("Feeders Over %d Hrs" % OVER_HOURS, "int"), ("Longest Outage (HH:MM)", "dur"), ("Most Interruptions", "int"),
        ("Big Enough To Rank", "text")])
    ss_count = {}
    for r in q("SELECT discom, circle, division, sskey, MIN(sscode), MIN(ssname), %s FROM f WHERE %s"
               " GROUP BY discom, circle, division, sskey" % (AGG, main)):
        n, ev, secs, over, many, mx, mxev = r[6:]
        t_ss.rows.append([r[0], r[1], r[2], r[4], r[5], n, ev, secs, round(secs / 60.0, 2),
                          safe_div(ev, n), safe_div(secs, n), safe_div(secs / 60.0, n), over, mx, mxev,
                          "Yes" if n >= MIN_FEEDERS_SUBSTATION else "No"])
        ss_count[r[0]] = ss_count.get(r[0], 0) + 1
    t_ss.rows.sort(key=lambda r: (discom_rank(r[0]), r[1], r[2], r[4], r[3]))

    # ---- summary -----------------------------------------------------------
    other = {}
    for r in q("SELECT discom, cat, over24, COUNT(*) FROM f GROUP BY discom, cat, over24"):
        o = other.setdefault(r[0], {"V33": 0, "AGRI": 0, "IND": 0, "O24": 0})
        if r[2]:
            o["O24"] += r[3]
        elif r[1] in o:
            o[r[1]] += r[3]
    direct = dict((r[0], tuple(r[1:])) for r in q("SELECT discom, %s FROM f WHERE %s GROUP BY discom" % (AGG, main)))
    t_sum = Table("Summary", [("Discom", "text"), ("Data From", "date"), ("Data Up To", "date"), ("Days", "int"),
                              ("Circles", "int"), ("Divisions", "int"), ("Substations With Outage", "int")]
                  + MEASURE_COLS + [("33 kV Feeders With Outage", "int"), ("Agriculture Feeders With Outage", "int"),
                                    ("Industrial Feeders Left Out", "int"), ("Over 24 Hrs Left Out", "int")])
    discoms = sorted(set(r["discom"] for r in div_rows), key=discom_rank)
    state = {"cir": 0, "div": 0, "ss": 0, "total": 0, "m": ZERO, "o": {"V33": 0, "AGRI": 0, "IND": 0, "O24": 0}}
    sum_rows = []
    for d in discoms:
        rows_d = [r for r in div_rows if r["discom"] == d and r["received"]]
        m = ZERO
        for r in rows_d:
            m = add_m(m, r["m"])
        if m != direct.get(d, ZERO):
            raise UserError("Self-check failed: the division figures of %s do not add up to its total. "
                            "No report was made. Please send the log to the person who looks after this app." % d)
        o = other.get(d, {"V33": 0, "AGRI": 0, "IND": 0, "O24": 0})
        total = sum(r["total"] for r in rows_d)
        ncir = len(set(r["circle"] for r in rows_d))
        sum_rows.append([d, d_min, d_max, days, ncir, len(rows_d), ss_count.get(d, 0)] + measure_values(total, days, m)
                        + [o["V33"], o["AGRI"], o["IND"], o["O24"]])
        state["cir"] += ncir
        state["div"] += len(rows_d)
        state["ss"] += ss_count.get(d, 0)
        state["total"] += total
        state["m"] = add_m(state["m"], m)
        for kk in state["o"]:
            state["o"][kk] += o[kk]
    t_sum.rows.append([STATE, d_min, d_max, days, state["cir"], state["div"], state["ss"]]
                      + measure_values(state["total"], days, state["m"])
                      + [state["o"]["V33"], state["o"]["AGRI"], state["o"]["IND"], state["o"]["O24"]])
    t_sum.rows.extend(sum_rows)

    # ---- feeder type -------------------------------------------------------
    step(0.3, "Working out the figures")
    t_type = Table("Feeder Type", [
        ("Discom", "text"), ("Circle", "text"), ("Division", "text"), ("Feeder Type", "text"), ("Counted In Averages", "text"),
        ("Feeders With Outage", "int"), ("Interruptions", "int"), ("Outage Duration (HH:MM)", "dur"), ("Outage Min", "num2"),
        ("Feeders Over %d Hrs" % OVER_HOURS, "int")])
    for r in q("SELECT discom, circle, division, ftype, cat, COUNT(*), SUM(ev), SUM(secs), SUM(secs>%d) FROM f"
               " WHERE %s AND cat<>'V33' GROUP BY discom, circle, division, ftype, cat" % (OVER_HOURS * 3600, ok)):
        t_type.rows.append([r[0], r[1], r[2], r[3], "Yes" if r[4] == "MAIN" else "No", r[5], r[6], r[7],
                            round(r[7] / 60.0, 2), r[8]])
    t_type.rows.sort(key=lambda r: (discom_rank(r[0]), r[1], r[2], r[3]))

    # ---- bands -------------------------------------------------------------
    def band_case(col, bands):
        parts = []
        for i, (_, top) in enumerate(bands):
            parts.append(("WHEN %s<=%d THEN %d" % (col, top, i + 1)) if top is not None else ("ELSE %d" % (i + 1)))
        return "CASE " + " ".join(parts) + " END"

    def band_table(name, heading, col, bands):
        t = Table(name, [("Discom", "text"), ("Circle", "text"), ("Band No", "int"), (heading, "text"),
                         ("Feeders", "int"), ("Interruptions", "int"), ("Outage Min", "num2")])
        got = {}
        for r in q("SELECT discom, circle, %s AS b, COUNT(*), SUM(ev), SUM(secs) FROM f WHERE %s"
                   " GROUP BY discom, circle, b" % (band_case(col, bands), main)):
            got[(r[0], r[1], r[2])] = r[3:]
        for (d, c2) in sorted(set((k[0], k[1]) for k in got), key=lambda k: (discom_rank(k[0]), k[1])):
            for i, (label, _) in enumerate(bands):
                n, ev, secs = got.get((d, c2, i + 1), (0, 0, 0))
                t.rows.append([d, c2, i + 1, label, n, ev, round(secs / 60.0, 2)])
        return t
    t_ob = band_table("Outage Bands", "Outage Band", "secs", OUTAGE_BANDS)
    t_ib = band_table("Interruption Bands", "Interruption Band", "ev", INTERRUPTION_BANDS)

    # ---- feeder lists ------------------------------------------------------
    step(0.45, "Making the feeder lists")
    FCOLS = "discom, d, circle, division, ssname, fname, ftype, msn, ev, secs"
    LIST_COLS = [("Date", "date"), ("Circle", "text"), ("Division", "text"), ("Substation", "text"), ("Feeder", "text"),
                 ("Feeder Type", "text"), ("Meter No", "text"), ("Interruptions", "int"),
                 ("Outage Duration (HH:MM)", "dur"), ("Outage Min", "num2")]

    def list_vals(r):        # r = FCOLS row, without discom
        return [r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], round(r[9] / 60.0, 2)]

    ORDERS = {
        ("Outage", "Top"): "secs DESC, ev DESC, circle, division, fname, k",
        ("Outage", "Bottom"): "secs ASC, ev ASC, circle, division, fname, k",
        ("Interruptions", "Top"): "ev DESC, secs DESC, circle, division, fname, k",
        ("Interruptions", "Bottom"): "ev ASC, secs ASC, circle, division, fname, k",
        ("Min per Interruption", "Top"): "(secs*1.0/ev) DESC, secs DESC, circle, division, fname, k",
        ("Min per Interruption", "Bottom"): "(secs*1.0/ev) ASC, secs ASC, circle, division, fname, k",
    }
    real_discoms = [d for d in discoms if d in direct]

    def top_sheet(name, order):
        t = Table(name, [("Discom", "text"), ("Rank", "int")] + LIST_COLS)
        for d in real_discoms:
            for i, r in enumerate(q("SELECT %s FROM f WHERE %s AND discom=? ORDER BY %s LIMIT %d"
                                    % (FCOLS, main, order, TOP_N_SHEETS), (d,)), 1):
                t.rows.append([d, i] + list_vals(r))
        return t
    t_top_o = top_sheet("Top %d Outage" % TOP_N_SHEETS, ORDERS[("Outage", "Top")])
    t_top_i = top_sheet("Top %d Interruption" % TOP_N_SHEETS, ORDERS[("Interruptions", "Top")])

    t_over = Table("Over %d Hrs" % OVER_HOURS, [("Discom", "text")] + LIST_COLS)
    for r in q("SELECT %s FROM f WHERE %s AND secs>%d ORDER BY circle, division, secs DESC, fname, k" % (FCOLS, main, OVER_HOURS * 3600)):
        t_over.rows.append([r[0]] + list_vals(r))
    t_over.rows.sort(key=lambda r: discom_rank(r[0]))       # stable: keeps circle, division, time order

    step(0.6, "Making the feeder lists")
    t_rank = Table("Feeder Ranks", [("List For", "text"), ("Measure", "text"), ("Side", "text"), ("Rank", "int"), ("Discom", "text")]
                   + LIST_COLS + [("Min per Interruption", "num2")])

    def rank_vals(r):
        return list_vals(r) + [round(safe_div(r[9] / 60.0, r[8]), 2)]
    for scope in [STATE] + real_discoms:
        where, args = (main + " AND ev>0", ()) if scope == STATE else (main + " AND ev>0 AND discom=?", (scope,))
        for (measure, side), order in sorted(ORDERS.items()):
            for i, r in enumerate(q("SELECT %s FROM f WHERE %s ORDER BY %s LIMIT %d" % (FCOLS, where, order, RANK_N_STATE), args), 1):
                t_rank.rows.append([scope, measure, side, i, r[0]] + rank_vals(r))
    for measure in ("Outage", "Interruptions"):
        last, n = None, 0
        cur = db.execute("SELECT %s FROM f WHERE %s ORDER BY discom, circle, division, %s" % (FCOLS, main, ORDERS[(measure, "Top")]))
        for r in cur:
            k = (r[0], r[2], r[3])
            if k != last:
                last, n = k, 0
            n += 1
            if n <= RANK_N_DIVISION:
                t_rank.rows.append(["DIVISION", measure, "Top", n, r[0]] + rank_vals(r))
    t_rank.rows.sort(key=rank_row_order([STATE] + real_discoms + ["DIVISION"]))

    # ---- 33 kV, over 24 hrs, day wise ---------------------------------------
    step(0.75, "Making the other sheets")
    t_33 = Table("33 kV", [("Discom", "text"), ("Circle", "text"), ("Division", "text"), ("Feeders With Outage", "int"),
                           ("Interruptions", "int"), ("Outage Duration (HH:MM)", "dur"), ("Outage Min", "num2"),
                           ("Feeders Over %d Hrs" % OVER_HOURS, "int"), ("Longest Outage (HH:MM)", "dur")])
    for r in q("SELECT discom, circle, division, COUNT(*), SUM(ev), SUM(secs), SUM(secs>%d), MAX(secs) FROM f"
               " WHERE %s AND cat='V33' GROUP BY discom, circle, division" % (OVER_HOURS * 3600, ok)):
        t_33.rows.append([r[0], r[1], r[2], r[3], r[4], r[5], round(r[5] / 60.0, 2), r[6], r[7]])
    t_33.rows.sort(key=lambda r: (discom_rank(r[0]), r[1], r[2]))

    t_ag = Table("Agriculture", [
        ("Discom", "text"), ("Circle", "text"), ("Division", "text"), ("Feeders With Outage", "int"), ("Interruptions", "int"),
        ("Outage Duration (HH:MM)", "dur"), ("Outage Min", "num2"),
        ("Avg Interruption per Feeder With Outage", "num2", "IF([Feeders With Outage]=0,0,[Interruptions]/[Feeders With Outage])"),
        ("Avg Outage per Feeder With Outage (HH:MM)", "dur", "IF([Feeders With Outage]=0,0,[Outage Duration (HH:MM)]/[Feeders With Outage])"),
        ("Avg Outage Min per Feeder With Outage", "num2", "IF([Feeders With Outage]=0,0,[Outage Duration (HH:MM)]*1440/[Feeders With Outage])"),
        ("Feeders Over %d Hrs" % OVER_HOURS, "int"), ("Longest Outage (HH:MM)", "dur"), ("Moved By Type List", "int")])
    for r in q("SELECT discom, circle, division, COUNT(*), SUM(ev), SUM(secs), SUM(secs>%d), MAX(secs), SUM(fixed) FROM f"
               " WHERE %s AND cat='AGRI' GROUP BY discom, circle, division" % (OVER_HOURS * 3600, ok)):
        t_ag.rows.append([r[0], r[1], r[2], r[3], r[4], r[5], round(r[5] / 60.0, 2), safe_div(r[4], r[3]), safe_div(r[5], r[3]),
                          safe_div(r[5] / 60.0, r[3]), r[6], r[7], r[8]])
    t_ag.rows.sort(key=lambda r: (discom_rank(r[0]), r[1], r[2]))

    t_agf = Table("Agriculture Feeders", [
        ("Discom", "text"), ("Date", "date"), ("Circle", "text"), ("Division", "text"), ("Substation", "text"), ("Feeder", "text"),
        ("Type In File", "text"), ("Moved By Type List", "text"), ("Meter No", "text"), ("Interruptions", "int"),
        ("Outage Duration (HH:MM)", "dur"), ("Outage Min", "num2")])
    for r in q("SELECT %s, otype, fixed FROM f WHERE %s AND cat='AGRI' ORDER BY circle, division, secs DESC, fname, k" % (FCOLS, ok)):
        t_agf.rows.append([r[0], r[1], r[2], r[3], r[4], r[5], r[10], "Yes" if r[11] else "No", r[7], r[8], r[9], round(r[9] / 60.0, 2)])
    t_agf.rows.sort(key=lambda r: discom_rank(r[0]))

    t_24 = Table("Over 24 Hrs Left Out", [("Discom", "text"), ("Date", "date"), ("Circle", "text"), ("Division", "text"),
                                          ("Substation", "text"), ("Feeder", "text"), ("Feeder Type", "text"), ("Voltage", "text"),
                                          ("Meter No", "text"), ("Interruptions", "int"), ("Outage Duration (HH:MM)", "dur"),
                                          ("Outage Hrs", "num1")])
    for r in q("SELECT %s, cat FROM f WHERE over24=1 ORDER BY secs DESC, k" % FCOLS):
        t_24.rows.append([r[0], r[1], r[2], r[3], r[4], r[5], r[6], "33 kV" if r[10] == "V33" else "11 kV", r[7], r[8], r[9],
                          round(r[9] / 3600.0, 1)])
    t_24.rows.sort(key=lambda r: discom_rank(r[0]))

    t_day = Table("Day Wise", [("Date", "date"), ("Discom", "text"), ("Days", "int")] + MEASURE_COLS)
    total_of = dict((row[0], row[t_sum.col("Total Feeders")]) for row in t_sum.rows)
    per_day = {}
    for r in q("SELECT d, discom, %s FROM f WHERE %s GROUP BY d, discom" % (AGG, main)):
        per_day[(r[0], r[1])] = tuple(r[2:])
    for day in sorted(set(k[0] for k in per_day)):
        m_state = ZERO
        day_rows = []
        for d in discoms:
            m = per_day.get((day, d), ZERO)
            m_state = add_m(m_state, m)
            day_rows.append([day, d, 1] + measure_values(total_of.get(d, 0), 1, m))
        t_day.rows.append([day, STATE, 1] + measure_values(total_of.get(STATE, 0), 1, m_state))
        t_day.rows.extend(day_rows)

    # ---- company wise ------------------------------------------------------
    t_co = Table("Company", [("Discom", "text"), ("Company", "text"), ("Files", "int"), ("Rows Read", "int"), ("Broken Rows", "int"),
                             ("Feeder Rows Kept", "int"), ("Used For Averages", "int"), ("Agriculture", "int"), ("Industrial", "int"),
                             ("33 kV", "int"), ("Over 24 Hrs Left Out", "int")])
    by_slot = {}
    for r in q("SELECT src, cat, over24, COUNT(*) FROM f GROUP BY src, cat, over24"):
        b = by_slot.setdefault(r[0], {"MAIN": 0, "AGRI": 0, "IND": 0, "V33": 0, "O24": 0})
        b["O24" if r[2] else r[1]] += r[3]
    for i, sl in enumerate(counts.get("slots", [])):
        b = by_slot.get(i, {"MAIN": 0, "AGRI": 0, "IND": 0, "V33": 0, "O24": 0})
        t_co.rows.append([sl["section"] or "NOT STATED", sl["company"], sl["files"], sl["rows"], sl["broken"], sum(b.values()),
                          b["MAIN"], b["AGRI"], b["IND"], b["V33"], b["O24"]])
    t_co.rows.sort(key=lambda r: discom_rank(r[0]))         # stable: keeps the order of the window

    # ---- data check --------------------------------------------------------
    cat_n = dict((r[0], r[1]) for r in q("SELECT cat, COUNT(*) FROM f WHERE over24=0 GROUP BY cat"))
    n_over24 = q("SELECT COUNT(*) FROM f WHERE over24=1")[0][0]
    no_row_divs = [r for r in div_rows if not r["has_rows"] and r["cir_rows"]]
    no_row_circles = sorted(set((r["discom"], r["circle"]) for r in div_rows if not r["cir_rows"]))
    no_row_text = "left out of the averages" if LEAVE_OUT_DIVISIONS_WITH_NO_ROWS else "kept, counted as no outage"
    n_fixed = q("SELECT COUNT(*) FROM f WHERE fixed=1")[0][0]
    fixed_to = q("SELECT ftype, COUNT(*) FROM f WHERE fixed=1 GROUP BY ftype ORDER BY ftype")
    agri_like = []
    for r in db.execute("SELECT fname, division, msn, ftype, discom, fcode, circle FROM f WHERE %s ORDER BY circle, division, fname" % main):
        if RE_AGRI_NAME.search(r[0]):
            agri_like.append(r)
    code_twice = q("SELECT fcode, circle, COUNT(*), GROUP_CONCAT(fname || ' [' || otype || CASE WHEN cat='V33' THEN ', 33 kV]' ELSE ']' END, ' and ')"
                   " FROM f WHERE fcode<>'' AND msn='' GROUP BY d, circle, fcode HAVING COUNT(*)>1 ORDER BY circle, fcode")
    t_chk = Table("Data Check", [("Item", "text"), ("Count", "int")])
    t_chk.rows = [
        ["CSV files read", counts["files"]],
        ["Rows read in all files", counts["rows"]],
        ["Blank rows skipped", counts["blank"]],
        ["Heading, note and total rows skipped", counts["notes"]],
        ["Broken rows left out", counts["broken"]],
        ["Rows with no outage (0 interruptions and 0 time) skipped", counts["zero"]],
        ["Outage times written with 60 minutes, such as 11:60 (read as the next hour)", counts.get("sixty", 0)],
        ["Repeat rows with the same figures (counted once)", counts["dup_same"]],
        ["Repeat rows with different figures (lower outage kept)", counts["dup_diff"]],
        ["Feeder rows kept", counts["kept"]],
        ["11 kV feeders used for averages", cat_n.get("MAIN", 0)],
        ["11 kV Agriculture feeders (left out of averages, shown separately)", cat_n.get("AGRI", 0)],
        ["11 kV Industrial feeders left out of averages (SBPDCL rule; NBPDCL counts them)", cat_n.get("IND", 0)],
        ["33 kV feeders (shown separately)", cat_n.get("V33", 0)],
        ["Feeders over 24 hours in a day (left out of all figures)", n_over24],
        ["Feeders with type not given or of another kind, such as Others or Dedicated (counted as 11 kV)", counts["type_unknown"]],
        ["Feeders whose type was changed by the type list", n_fixed],
        ["Feeders not typed Agriculture whose name looks like Agriculture (still in averages)", len(agri_like)],
        ["Feeder codes that come on more than one row with no meter number (every row kept)", len(code_twice)],
        ["Circles not in the feeder list", len(counts["unknown_circles"])],
        ["Divisions not in the feeder list", len(counts["unknown_divs"])],
        ["Divisions in the feeder list with no rows (%s)" % no_row_text, len(no_row_divs)],
        ["Circles in the feeder list with no rows (%s)" % no_row_text, len(no_row_circles)],
        ["Divisions where feeders with outage are more than Total Feeders", len(too_many)],
    ]

    # ---- self-check: every sheet must add up to the same totals -------------
    step(0.9, "Checking that all sheets agree")
    s = t_sum.rows[0]
    want = (s[t_sum.col("Feeders With Outage")], s[t_sum.col("Interruptions")], s[t_sum.col("Outage Duration (HH:MM)")])

    def total_of_table(t, where=None, n="Feeders With Outage", secs="Outage Duration (HH:MM)", scale=1):
        rows = [r for r in t.rows if (where is None or where(r))]
        return (sum(r[t.col(n)] for r in rows), sum(r[t.col("Interruptions")] for r in rows),
                int(round(sum(r[t.col(secs)] for r in rows) * scale)))
    checks = [
        ("Summary DISCOM rows", total_of_table(t_sum, lambda r: r[0] != STATE)),
        ("Circle", total_of_table(t_cir)),
        ("Division", total_of_table(t_div)),
        ("Substation", total_of_table(t_ss)),
        ("Feeder Type", total_of_table(t_type, lambda r: r[4] == "Yes")),
        ("Outage Bands", total_of_table(t_ob, n="Feeders", secs="Outage Min", scale=60)),
        ("Interruption Bands", total_of_table(t_ib, n="Feeders", secs="Outage Min", scale=60)),
        ("Day Wise", total_of_table(t_day, lambda r: r[1] == STATE)),
    ]
    for name, got in checks:
        same = got[0] == want[0] and got[1] == want[1] and abs(got[2] - want[2]) <= max(2, len(t_ob.rows))
        if not same:
            raise UserError("Self-check failed: sheet '%s' adds up to %s but the Summary says %s. No report was made. "
                            "Please send the log to the person who looks after this app." % (name, got, want))
    over_n = s[t_sum.col("Feeders Over %d Hrs" % OVER_HOURS)]
    if len(t_over.rows) != over_n:
        raise UserError("Self-check failed: the Over %d Hrs list has %d feeders but the Summary says %d. No report was made."
                        % (OVER_HOURS, len(t_over.rows), over_n))
    if counts["rows"] != (counts["blank"] + counts["notes"] + counts["broken"] + counts["zero"]
                          + counts["dup_same"] + counts["dup_diff"] + counts["kept"]):
        raise UserError("Self-check failed: the rows read do not match the rows counted. No report was made.")
    if counts["kept"] != sum(cat_n.values()) + n_over24:
        raise UserError("Self-check failed: the feeder rows kept do not match the feeder groups. No report was made.")
    if sum(r[5] for r in t_co.rows) != counts["kept"]:
        raise UserError("Self-check failed: the Company sheet does not add up to the feeder rows kept. No report was made.")
    if not (sum(r[3] for r in t_ag.rows) == len(t_agf.rows) == cat_n.get("AGRI", 0)
            == sum(r[5] for r in t_type.rows if r[3] == "Agriculture")):
        raise UserError("Self-check failed: the Agriculture sheets do not agree with each other. No report was made.")
    log("ok", "Self-check passed: Summary, Circle, Division, Substation, Feeder Type, both band sheets, Day Wise "
              "and the Over %d Hrs list all give the same totals. The Agriculture sheets agree with each other." % OVER_HOURS)

    # ---- messages for the user ---------------------------------------------
    for (d, c2) in no_row_circles:
        log("warn", "No rows came for circle %s (%s). Its divisions are %s. "
                    "CHECK WHETHER A COMPANY'S FILE IS MISSING: if it is, the %s averages are too low." % (c2, d, no_row_text, d))
    if no_row_divs:
        log("warn", "No rows came for %d division(s): %s. They are %s. Check whether a file is missing."
            % (len(no_row_divs), ", ".join(r["division"] for r in no_row_divs[:12]) + (" ..." if len(no_row_divs) > 12 else ""), no_row_text))
    if n_fixed:
        log("info", "Type list: %d feeder(s) were moved to another type (%s)."
            % (n_fixed, ", ".join("%d to %s" % (n, t) for t, n in fixed_to)))
    for dname in sorted(set(r[4] for r in agri_like), key=discom_rank):
        these = [r for r in agri_like if r[4] == dname]
        log("warn", "%s: %d feeder(s) are not typed Agriculture but the name looks like Agriculture. They ARE counted in the averages. "
                    "If they are Agriculture feeders, add them to the Agriculture list (web page, or More > Agriculture list), or ask the company to correct the type. The full list: %s"
            % (dname, len(these), "; ".join("%s (%s, typed %s, %s)" % (r[0], r[1], r[3] or "blank",
                                                                        ("meter " + r[2]) if r[2] else ("code " + (r[5] or "not given")))
                                            for r in these)))
    if code_twice:
        log("warn", "%d feeder code(s) come on more than one row, with no meter number to tell the rows apart. "
                    "Every such row is kept as it came. Ask the company which is right: %s%s"
            % (len(code_twice), "; ".join("%s in %s (%s)" % (r[0], r[1], r[3]) for r in code_twice[:8]),
               " ..." if len(code_twice) > 8 else ""))
    if counts["dup_diff"]:
        log("warn", "%d feeder row(s) came more than once with different figures. The lower outage was kept for each."
            % counts["dup_diff"])
    for circle, n in sorted(counts["unknown_circles"].items()):
        log("warn", "Circle '%s' (%d rows) is not in the feeder list, so its DISCOM is not known. It is shown as NOT MAPPED. "
                    "Add it to %s (More > Feeder list)." % (circle, n, MASTER_FILE_NAME))
    for (circle, division), n in sorted(counts["unknown_divs"].items()):
        log("warn", "Division '%s' of circle %s (%d rows) is not in the feeder list. Its averages use only the feeders in the file. "
                    "Add it to %s (More > Feeder list)." % (division, circle, n, MASTER_FILE_NAME))
    for r in too_many:
        log("warn", "Division %s: %d feeders had an outage but the feeder list says it has only %d. Please update the feeder list."
            % (r["division"], r["m"][0], r["total"]))
    if n_over24:
        log("warn", "%d feeder(s) show more than 24 hours of outage in one day. They are left out of all figures and "
                    "listed in the sheet 'Over 24 Hrs Left Out'." % n_over24)
    if counts["own_date"] and days > 1:
        log("info", "The files hold %d different dates, from %s to %s." % (days, fmt_date(d_min), fmt_date(d_max)))

    tables = [t_sum, t_cir, t_div, t_ss, t_type, t_ob, t_ib, t_top_o, t_top_i, t_over, t_rank, t_ag, t_agf, t_33, t_24, t_day, t_co, t_chk]
    info = {"d_min": d_min, "d_max": d_max, "days": days, "discoms": real_discoms, "summary": t_sum,
            "present": sorted(discoms_with_rows, key=discom_rank),
            "agri_like": [{"name": r[0], "division": r[1], "meter": r[2], "type": r[3], "discom": r[4], "code": r[5], "circle": r[6]} for r in agri_like]}
    return tables, info


def filter_tables(tables, discom):
    """The same sheets, holding only one DISCOM."""
    out = []
    for t in tables:
        if t.name == "Data Check":
            continue
        heads = t.heads()
        nt = Table(t.name, t.cols)
        if "Discom" in heads:
            i = heads.index("Discom")
            nt.rows = [r for r in t.rows if r[i] == discom]
        else:
            nt.rows = list(t.rows)
        if t.name == "Feeder Ranks":
            nt.rows = [r for r in nt.rows if r[0] in (discom, "DIVISION")]
        out.append(nt)
    return out


# =============================================================================
#  ONE DISCOM NOW, THE OTHER FROM A REPORT MADE EARLIER
# =============================================================================


def rows_from_file(path, like):
    """Reads a DISCOM report made earlier and gives its rows back in the app's own form.
    Returns {sheet name: rows}. Raises UserError if the file does not fit this version of the app."""
    sheets = dict((sh["name"], sh) for sh in read_xlsx(path, results=True))
    out = {}
    for t in like:
        if t.name == "Data Check":
            continue
        sh = sheets.get(t.name)
        if sh is None or sh["head"] != t.heads():
            raise UserError("its sheet '%s' is missing or has other headings; it was made by another version of the app" % t.name)
        rows = []
        for raw in sh["rows"]:
            row = []
            for (_, kind, formula), v in zip(t.cols, raw):
                if kind == "text":
                    row.append(str(v))
                elif v == "" or v is None or isinstance(v, str):
                    row.append(0)
                elif kind == "int":
                    row.append(int(round(float(v))))
                elif kind == "dur":
                    row.append(float(v) * 86400.0 if formula else int(round(float(v) * 86400.0)))
                elif kind == "date":
                    row.append(int(round(float(v))) + EXCEL_DAY0)
                else:
                    row.append(float(v))
            rows.append(row)
        out[t.name] = rows
    return out


def add_discom(tables, other_rows):
    """Joins the DISCOM just read with a DISCOM taken from a report made earlier,
    and works out the BIHAR lines again from the two."""
    over_head = "Feeders Over %d Hrs" % OVER_HOURS
    many_head = "Feeders %d Plus Interruptions" % MANY_INTERRUPTIONS
    out = []
    bihar_total = 0
    totals = {}
    for t in tables:
        nt = Table(t.name, t.cols)
        heads = t.heads()
        if t.name == "Data Check":
            nt.rows = [list(r) for r in t.rows] + [["DISCOMs taken from a report made earlier (not read in this run)", 1]]
            out.append(nt)
            continue
        rows = [list(r) for r in t.rows] + [list(r) for r in other_rows.get(t.name, [])]
        col = t.col

        def m_of(r):
            return (r[col("Feeders With Outage")], r[col("Interruptions")], r[col("Outage Duration (HH:MM)")], r[col(over_head)],
                    r[col(many_head)], r[col("Longest Outage (HH:MM)")], r[col("Most Interruptions")])

        if t.name == "Summary":
            rows = sorted([r for r in rows if r[0] != STATE], key=lambda r: discom_rank(r[0]))
            m = ZERO
            for r in rows:
                m = add_m(m, m_of(r))
                totals[r[0]] = r[col("Total Feeders")]
            tot = lambda h: sum(r[col(h)] for r in rows)
            days = max(r[col("Days")] for r in rows)
            bihar_total = tot("Total Feeders")
            nt.rows = [[STATE, min(r[col("Data From")] for r in rows), max(r[col("Data Up To")] for r in rows), days,
                        tot("Circles"), tot("Divisions"), tot("Substations With Outage")] + measure_values(bihar_total, days, m)
                       + [tot("33 kV Feeders With Outage"), tot("Agriculture Feeders With Outage"),
                          tot("Industrial Feeders Left Out"), tot("Over 24 Hrs Left Out")]] + rows
        elif t.name == "Day Wise":
            rows = [r for r in rows if r[1] != STATE]
            got = dict(((r[0], r[1]), r) for r in rows)
            discoms = sorted(set(r[1] for r in rows) | set(totals), key=discom_rank)
            for day in sorted(set(r[0] for r in rows)):
                m, day_rows = ZERO, []
                for d in discoms:
                    r = got.get((day, d)) or ([day, d, 1] + measure_values(totals.get(d, 0), 1, ZERO))
                    m = add_m(m, m_of(r))
                    day_rows.append(r)
                nt.rows.append([day, STATE, 1] + measure_values(bihar_total, 1, m))
                nt.rows.extend(day_rows)
        elif t.name == "Feeder Ranks":
            rows = [r for r in rows if r[0] != STATE]
            lists = [r for r in rows if r[0] != "DIVISION"]
            bihar = []
            for measure in ("Outage", "Interruptions", "Min per Interruption"):
                for side in ("Top", "Bottom"):
                    g = -1 if side == "Top" else 1
                    cand = [r for r in lists if r[1] == measure and r[2] == side]
                    tail = lambda r: (r[6], r[7], r[9], r[5], r[11])        # circle, division, feeder, date, meter
                    if measure == "Outage":
                        cand.sort(key=lambda r: (g * r[13], g * r[12]) + tail(r))
                    elif measure == "Interruptions":
                        cand.sort(key=lambda r: (g * r[12], g * r[13]) + tail(r))
                    else:
                        cand.sort(key=lambda r: (g * safe_div(r[13], r[12]), g * r[13]) + tail(r))
                    for i, r in enumerate(cand[:RANK_N_STATE], 1):
                        bihar.append([STATE, measure, side, i] + r[4:])
            scopes = [STATE] + sorted(set(r[0] for r in lists), key=discom_rank) + ["DIVISION"]
            nt.rows = sorted(bihar + rows, key=rank_row_order(scopes))
        else:
            i = heads.index("Discom")
            nt.rows = sorted(rows, key=lambda r: discom_rank(r[i]))     # stable: each DISCOM keeps its own order
        out.append(nt)

    by = dict((t.name, t) for t in out)
    s = by["Summary"].rows[0]
    want = (s[by["Summary"].col("Feeders With Outage")], s[by["Summary"].col("Interruptions")], s[by["Summary"].col("Outage Duration (HH:MM)")])
    for name in ("Circle", "Division", "Substation"):
        t = by[name]
        got = (sum(r[t.col("Feeders With Outage")] for r in t.rows), sum(r[t.col("Interruptions")] for r in t.rows),
               sum(r[t.col("Outage Duration (HH:MM)")] for r in t.rows))
        if got != want:
            raise UserError("Self-check failed after joining the two DISCOMs: sheet '%s' adds up to %s but the Summary says %s. "
                            "No report was made. Make both DISCOM reports again in one run." % (name, got, want))
    return out


# =============================================================================
#  WRITING THE EXCEL FILE  (plain .xlsx, made with the zip tools of Python)
# =============================================================================

EXCEL_DAY0 = datetime.date(1899, 12, 30).toordinal()
STYLE_OF = {"head": 1, "text": 2, "int": 3, "num1": 4, "num2": 5, "dur": 6, "date": 7}
FORMAT_OF = {"text": "@", "int": "#,##0", "num1": "0.0", "num2": "0.00", "dur": "[h]:mm", "date": "dd-mmm-yyyy"}

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="4"><numFmt numFmtId="164" formatCode="0.0"/><numFmt numFmtId="165" formatCode="[h]:mm"/>'
    '<numFmt numFmtId="166" formatCode="dd\\-mmm\\-yyyy"/><numFmt numFmtId="167" formatCode="#,##0"/></numFmts>'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>'
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills>'
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>'
    '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="8">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
    '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    '<xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>')


def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def xml_text(v):
    v = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\uFFFE\uFFFF]", " ", str(v))
    return xml_escape(v)


def num_text(v):
    if isinstance(v, int):
        return str(v)
    return repr(round(float(v), 12))


def formula_for_row(template, heads, row_no):
    def rep(m):
        return "%s%d" % (col_letter(heads.index(m.group(1))), row_no)
    return re.sub(r"\[([^\]]+)\]", rep, template)


def cell_number(kind, v):
    """The number Excel stores for a value of this kind."""
    if kind == "dur":
        return float(v) / 86400.0
    if kind == "date":
        return int(v) - EXCEL_DAY0
    return v


def write_xlsx(path, tables, cancel=None):
    names = [t.name for t in tables]
    for n in names:
        if len(n) > 31 or re.search(r"[\[\]:*?/\\]", n):
            raise UserError("Sheet name not allowed in Excel: %s" % n)
    tmp = path + ".part"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                       '<Default Extension="xml" ContentType="application/xml"/>'
                       '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                       '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                       + "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
                                 for i in range(len(tables)))
                       + '</Types>')
            z.writestr("_rels/.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                       '</Relationships>')
            z.writestr("xl/workbook.xml",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                       + "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (xml_text(n), i + 1, i + 1) for i, n in enumerate(names))
                       + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
            z.writestr("xl/_rels/workbook.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       + "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
                                 for i in range(len(tables)))
                       + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(tables) + 1)
                       + '</Relationships>')
            z.writestr("xl/styles.xml", STYLES_XML)
            for si, t in enumerate(tables):
                heads = t.heads()
                ncol = len(heads)
                widths = [min(45, max(9, len(h) * 0.62 + 2)) for h in heads]
                for row in t.rows[:400]:
                    for ci, v in enumerate(row):
                        if t.cols[ci][1] == "text":
                            widths[ci] = min(45, max(widths[ci], len(str(v)) + 2))
                last = "%s%d" % (col_letter(ncol - 1), len(t.rows) + 1)
                with z.open("xl/worksheets/sheet%d.xml" % (si + 1), "w") as fh:
                    w = lambda s: fh.write(s.encode("utf-8"))
                    w('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                      '<dimension ref="A1:%s"/>'
                      '<sheetViews><sheetView workbookViewId="0"%s><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
                      '<sheetFormatPr defaultRowHeight="15"/><cols>' % (last, ' tabSelected="1"' if si == 0 else ""))
                    for ci, wd in enumerate(widths):
                        w('<col min="%d" max="%d" width="%.1f" customWidth="1"/>' % (ci + 1, ci + 1, wd))
                    w('</cols><sheetData><row r="1" ht="45" customHeight="1">')
                    for ci, h in enumerate(heads):
                        w('<c r="%s1" t="inlineStr" s="1"><is><t xml:space="preserve">%s</t></is></c>' % (col_letter(ci), xml_text(h)))
                    w('</row>')
                    letters = [col_letter(ci) for ci in range(ncol)]
                    for ri, row in enumerate(t.rows):
                        if cancel is not None and ri % 5000 == 0 and cancel.is_set():
                            raise Cancelled()
                        rn = ri + 2
                        parts = ['<row r="%d">' % rn]
                        for ci, v in enumerate(row):
                            _, kind, formula = t.cols[ci]
                            ref = "%s%d" % (letters[ci], rn)
                            if kind == "text":
                                parts.append('<c r="%s" t="inlineStr" s="2"><is><t xml:space="preserve">%s</t></is></c>' % (ref, xml_text(v)))
                            elif formula:
                                parts.append('<c r="%s" s="%d"><f>%s</f><v>%s</v></c>'
                                             % (ref, STYLE_OF[kind], xml_text(formula_for_row(formula, heads, rn)), num_text(cell_number(kind, v))))
                            else:
                                parts.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, STYLE_OF[kind], num_text(cell_number(kind, v))))
                        parts.append('</row>')
                        w("".join(parts))
                    w('</sheetData><autoFilter ref="A1:%s"/></worksheet>' % last)
        if os.path.exists(path):
            os.remove(path)
        os.rename(tmp, path)
    except PermissionError:
        _remove_quietly(tmp)
        raise UserError("The Excel file could not be saved:\n%s\nIt is open in Excel, or the folder does not allow saving. "
                        "Close the file in Excel (or choose another folder with 'Change') and press the button again." % path)
    except OSError as e:
        _remove_quietly(tmp)
        raise UserError("The Excel file could not be saved:\n%s\n(%s)\nCheck that the output folder exists and the disk is not full." % (path, e))
    except BaseException:
        _remove_quietly(tmp)
        raise


def _remove_quietly(path):
    try:
        os.remove(path)
    except OSError:
        pass


# =============================================================================
#  READING AN EXCEL REPORT BACK  (to send exactly what is in the file)
# =============================================================================

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
BUILTIN_FORMATS = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 14: "dd-mm-yyyy",
                   15: "d-mmm-yy", 20: "h:mm", 21: "h:mm:ss", 22: "dd-mm-yyyy h:mm", 46: "[h]:mm:ss", 49: "@"}


def _ref_to_col(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx(path, results=False):
    """Reads the whole report file. Returns a list of sheets:
    {"name", "head": [...], "format": [...], "rows": [[...]]}. A formula cell is text starting with '='.
    With results=True a formula cell gives the number stored for it instead."""
    try:
        z = zipfile.ZipFile(path)
    except FileNotFoundError:
        raise UserError("The report file was not found:\n%s" % path)
    except PermissionError:
        raise UserError("The report file could not be opened:\n%s\nClose it in Excel and try again." % path)
    except (zipfile.BadZipFile, OSError):
        raise UserError("This is not a proper Excel (.xlsx) file, or it is damaged:\n%s\nMake the report again." % path)
    with z:
        try:
            names = z.namelist()
            shared = []
            if "xl/sharedStrings.xml" in names:
                for si in ElementTree.fromstring(z.read("xl/sharedStrings.xml")).findall(NS + "si"):
                    shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
            formats = []
            if "xl/styles.xml" in names:
                st = ElementTree.fromstring(z.read("xl/styles.xml"))
                custom = {}
                for nf in st.iter(NS + "numFmt"):
                    custom[int(nf.get("numFmtId"))] = (nf.get("formatCode") or "").replace("\\", "")
                xfs = st.find(NS + "cellXfs")
                for xf in (xfs.findall(NS + "xf") if xfs is not None else []):
                    fid = int(xf.get("numFmtId") or 0)
                    formats.append(custom.get(fid, BUILTIN_FORMATS.get(fid, "General")))
            wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
            rels = ElementTree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            target = dict((r.get("Id"), r.get("Target")) for r in rels)
            sheets = []
            for sh in wb.find(NS + "sheets").findall(NS + "sheet"):
                tgt = target[sh.get(RNS + "id")].lstrip("/")
                part = tgt if tgt.startswith("xl/") else "xl/" + tgt
                head, fmts, rows = None, None, []
                with z.open(part) as fh:
                    for _, el in ElementTree.iterparse(fh, events=("end",)):
                        if el.tag != NS + "row":
                            continue
                        vals, styles = [], []
                        for c in el.findall(NS + "c"):
                            ci = _ref_to_col(c.get("r") or "")
                            if ci < 0:
                                ci = len(vals)
                            while len(vals) <= ci:
                                vals.append("")
                                styles.append(None)
                            t = c.get("t")
                            f = c.find(NS + "f")
                            v = c.find(NS + "v")
                            if f is not None and (f.text or "").strip() and not results:
                                val = "=" + f.text.strip()
                            elif t == "inlineStr":
                                val = "".join(x.text or "" for x in c.iter(NS + "t"))
                            elif t == "s":
                                val = shared[int(v.text)] if v is not None and v.text else ""
                            elif t in ("str", "e"):
                                val = v.text if v is not None and v.text else ""
                            elif t == "b":
                                val = "TRUE" if (v is not None and v.text == "1") else "FALSE"
                            elif v is None or v.text is None or v.text == "":
                                val = ""
                            else:
                                num = float(v.text)
                                val = int(num) if num == int(num) and abs(num) < 1e15 and "." not in v.text and "E" not in v.text.upper() else num
                            is_text = t in ("inlineStr", "s", "str")
                            vals[ci] = val
                            sidx = int(c.get("s") or 0)
                            styles[ci] = "@" if (is_text and f is None) else (formats[sidx] if sidx < len(formats) else "General")
                        el.clear()
                        if head is None:
                            head = [str(x) for x in vals]
                            while head and head[-1] == "":
                                head.pop()
                            continue
                        vals = (vals + [""] * len(head))[:len(head)]
                        styles = (styles + [None] * len(head))[:len(head)]
                        if not any(x != "" for x in vals):
                            continue
                        if fmts is None:
                            fmts = list(styles)
                        else:
                            for i, s in enumerate(styles):
                                if fmts[i] is None and s is not None:
                                    fmts[i] = s
                        rows.append(vals)
                if head is None:
                    raise UserError("Sheet '%s' in the report file has no heading row. Make the report again." % sh.get("name"))
                fmts = [(f or "General") for f in (fmts or [None] * len(head))]
                sheets.append({"name": sh.get("name"), "head": head, "format": fmts, "rows": rows})
        except UserError:
            raise
        except Exception as e:
            raise UserError("The report file could not be read:\n%s\n(%s)\nMake the report again and do not edit it before sending." % (path, e))
    if not sheets:
        raise UserError("The report file has no sheets:\n%s" % path)
    return sheets


# =============================================================================
#  SENDING TO THE GOOGLE SHEET
# =============================================================================


def web_link_ready():
    return WEB_APP_URL.startswith("https://script.google.com/") and WEB_APP_URL.rstrip("/").endswith("/exec")


def discom_of_report(sheets, path):
    """A DISCOM report file holds one DISCOM only. Returns its name, or raises UserError."""
    names = set()
    for sh in sheets:
        if sh["name"] == "Summary" and "Discom" in sh["head"]:
            i = sh["head"].index("Discom")
            names = set(str(r[i]).strip().upper() for r in sh["rows"])
    if len(names) == 1 and list(names)[0] in ("SBPDCL", "NBPDCL"):
        return list(names)[0]
    if STATE in names or len(names) > 1:
        raise UserError("This file holds more than one DISCOM, so it was not sent:\n%s\n"
                        "Each DISCOM is sent on its own, so that one team cannot change the figures of the other. "
                        "Choose the file whose name ends with SBPDCL.xlsx or NBPDCL.xlsx." % path)
    raise UserError("This file is not a DISCOM report made by this app, so it was not sent:\n%s\n"
                    "Choose the file whose name ends with SBPDCL.xlsx or NBPDCL.xlsx." % path)


def send_parts(parts, password, log, progress=None):
    """Sends each DISCOM report on its own. parts: list of (DISCOM, file). Returns the dashboard link.
    If one DISCOM is refused the other is still sent; the error is raised at the end."""
    link, errors = None, []
    for discom, path in parts:
        try:
            link = send_report(path, password, log, progress)
        except UserError as e:
            errors.append("%s was NOT sent. %s" % (discom, e))
            if len(parts) > 1:
                log("error", errors[-1])
    if errors:
        err = UserError("\n\n".join(errors) + ("\n\n(The other DISCOM was sent and is on the dashboard.)" if len(errors) < len(parts) else ""))
        err.link = link                 # the dashboard link, if at least one DISCOM was sent
        raise err
    return link


def upload_body(xlsx_path, password):
    """The text sent to the web app for ONE DISCOM report: (DISCOM, text, rows, sheets). Used by the app and by the browser page."""
    sheets = read_xlsx(xlsx_path)
    discom = discom_of_report(sheets, xlsx_path)
    n_rows = sum(len(s["rows"]) for s in sheets)
    body = json.dumps({"action": "upload", "discom": discom, "password": password, "app": APP_VERSION, "sheets": sheets},
                      ensure_ascii=False, separators=(",", ":"))
    return discom, body, n_rows, len(sheets)


def send_report(xlsx_path, password, log, progress=None):
    """Reads the whole Excel file of ONE DISCOM, then sends it to the web app. Returns the dashboard link."""
    if not web_link_ready():
        raise UserError("The web app link is not yet written in the app, so nothing was sent.\n"
                        "Open outage_report_app.py, put the link in the line WEB_APP_URL near the top, and build the .exe again. "
                        "See the set-up notes, Part B.")
    if not password:
        raise UserError("No upload password was given, so nothing was sent. Press the button again and type the password when it is asked.")
    discom, body, n_rows, n_sheets = upload_body(xlsx_path, password)
    body = body.encode("utf-8")
    log("info", "Sending the %s report: %d sheets (%s rows, %s) ..." % (discom, n_sheets, "{:,}".format(n_rows), fmt_size(len(body))))
    if progress:
        progress(None, "Sending %s to the Google Sheet" % discom)
    req = urllib.request.Request(WEB_APP_URL, data=body, headers={"Content-Type": "text/plain;charset=utf-8",
                                                                  "User-Agent": "OutageReportApp/" + APP_VERSION})
    try:
        with urllib.request.urlopen(req, timeout=420) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise UserError("Google refused the upload (code %d). The Google Sheet was not changed.\n"
                            "In Apps Script, open Deploy > Manage deployments and check: Execute as 'Me', Who has access 'Anyone'." % e.code)
        if e.code == 404:
            raise UserError("The web app link was not found (code 404). The Google Sheet was not changed.\n"
                            "Check the link in WEB_APP_URL. It must be the link of the latest deployment and end with /exec.")
        raise UserError("Google gave an error (code %d). The Google Sheet was not changed. Wait a minute and press the button again, or use More > Send a report made earlier." % e.code)
    except urllib.error.URLError as e:
        reason = getattr(e, "reason", e)
        if ssl is not None and isinstance(reason, ssl.SSLError):
            raise UserError("A safe connection to Google could not be made (%s). Nothing was sent.\n"
                            "Check the date and time of this computer, and ask IT whether the office network blocks script.google.com." % reason)
        raise UserError("Could not reach Google (%s). Nothing was sent.\n"
                        "Check the internet connection, then press the button again, or use More > Send a report made earlier and pick the DISCOM file just made." % reason)
    except (TimeoutError, OSError) as e:
        raise UserError("Google did not answer in time (%s). The upload may or may not have finished.\n"
                        "Open the dashboard and check the date. If it is old, use More > Send a report made earlier and pick the DISCOM file." % e)
    try:
        ans = json.loads(raw)
    except ValueError:
        raise UserError("The web app answered with a web page instead of a result, so the upload was not accepted.\n"
                        "In Apps Script, open Deploy > Manage deployments and check: Execute as 'Me', Who has access 'Anyone'. "
                        "Also check that WEB_APP_URL is the link ending with /exec.")
    if not ans.get("ok"):
        code = ans.get("code", "")
        msg = ans.get("message", "The web app did not accept the upload.")
        if code == "BAD_PASSWORD":
            raise UserError("Wrong upload password for %s. The Google Sheet was not changed.\n%s\n"
                            "Press the button again: the app will ask for the %s password again. (To send only this DISCOM: More > Send a report made earlier, file ending with %s.xlsx.)"
                            % (discom, msg, discom, discom))
        if code == "LOCKED":
            raise UserError("Uploading for %s is locked after too many wrong passwords. The Google Sheet was not changed.\n%s" % (discom, msg))
        raise UserError("The %s upload was not accepted. %s" % (discom, msg))
    log("ok", "Google Sheet updated. %s" % ans.get("message", "done"))
    return ans.get("dashboard") or WEB_APP_URL


# =============================================================================
#  THE WHOLE JOB
# =============================================================================


def output_names(folder, d_max, discoms):
    stamp = fmt_date(d_max)
    main = os.path.join(folder, "%s_11KV_Outage_Report_Bihar.xlsx" % stamp)
    parts = dict((d, os.path.join(folder, "%s_11KV_Outage_Report_%s.xlsx" % (stamp, d))) for d in discoms)
    return main, parts


def make_report(inputs, out_folder, report_date, keep_other, log, progress, cancel):
    """Reads the files and writes the Excel reports. Returns (main file, info).
    inputs: list of (DISCOM part, company, file path), or plain file paths.
    keep_other: if only one DISCOM is read, take the other from the report already made for the same date."""
    inputs = as_inputs(inputs)
    if not inputs:
        raise UserError("No CSV file was added. Press 'Browse' next to a company name and choose its CSV file.")
    if not out_folder or not os.path.isdir(out_folder):
        raise UserError("The folder for the Excel reports does not exist:\n%s\nPress 'Change' and choose a folder." % (out_folder or "(none chosen)"))
    t0 = time.time()
    master = load_master(log)
    log("info", "Feeder numbers taken from the %s (%d divisions)." % (master.source, len(master.divisions)))
    types = load_type_list(log, master)
    log("info", "Agriculture list: %d feeder(s), from the %s. Every other feeder's type is taken as the company files give it."
        % (types.count, types.source))
    log("info", "Report date for rows that carry no date: %s." % report_date.strftime("%d.%m.%Y"))
    log("info", "Counting rule: 33 kV feeders are never counted. %s."
        % ". ".join("%s leaves out %s" % (d, " and ".join(TYPES_LEFT_OUT[d]) or "no type") for d in sorted(TYPES_LEFT_OUT)))
    store = Store()
    try:
        counts = read_files(inputs, report_date, master, types, store, log, lambda f, t: progress(0.75 * f, t), cancel)
        tables, info = build_tables(store, master, counts, log, lambda f, t: progress(0.75 + 0.15 * f, t), cancel)
    finally:
        store.close()
    both = ("SBPDCL", "NBPDCL")
    main_path, part_paths = output_names(out_folder, info["d_max"], both)
    progress(0.92, "Saving the Excel files")
    for d in info["present"]:
        if d in part_paths:                      # the DISCOMs read now: each gets its own file
            write_xlsx(part_paths[d], filter_tables(tables, d), cancel)
            log("ok", "%s report saved: %s" % (d, part_paths[d]))
    read_now = [d for d in both if d in info["present"]]
    day_text = fmt_date(info["d_max"])
    kept = []
    for d in both:
        if d in info["present"]:
            continue
        only = " and ".join(read_now + kept) or "what was read"
        path = part_paths[d]
        if not keep_other:
            log("warn", "No %s file was read. The Bihar report will show %s only." % (d, only))
        elif not os.path.isfile(path):
            log("warn", "No %s file was read, and no %s report of %s is in the output folder. "
                        "The Bihar report will show %s only." % (d, d, day_text, only))
        else:
            try:
                other = rows_from_file(path, tables)
                srow, st = other["Summary"][0], info["summary"]
                if (srow[st.col("Data From")], srow[st.col("Data Up To")], srow[st.col("Days")]) != (info["d_min"], info["d_max"], info["days"]):
                    raise UserError("it is not for the same date")
                tables = add_discom(tables, other)
            except UserError as e:
                if str(e).startswith("Self-check"):
                    raise
                log("warn", "The %s report made earlier could not be used (%s): %s\n"
                            "The Bihar report will show %s only. To have both, read both DISCOMs in one run."
                    % (d, str(e).split("\n")[0], path, only))
                continue
            kept.append(d)
            made = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d.%m.%Y at %H:%M")
            log("ok", "%s figures of %s were taken from the report made earlier (on %s): %s" % (d, day_text, made, path))
    write_xlsx(main_path, tables, cancel)
    log("ok", "Bihar report saved: %s" % main_path)
    t_sum = [t for t in tables if t.name == "Summary"][0]
    info["summary"] = t_sum
    info["kept"] = kept
    info["parts"] = [(d, part_paths[d]) for d in both if d in info["present"]]     # what was read now, and so may be sent
    for row in t_sum.rows:
        log("info", "%-7s total feeders %s, with outage %s, interruptions %s, average %s interruptions and %s (HH:MM) outage per feeder."
            % (row[0], "{:,}".format(row[t_sum.col("Total Feeders")]), "{:,}".format(row[t_sum.col("Feeders With Outage")]),
               "{:,}".format(row[t_sum.col("Interruptions")]), "%.2f" % row[t_sum.col("Avg Interruption per Feeder")],
               fmt_hhmm(row[t_sum.col("Avg Outage per Feeder (HH:MM)")])))
    if info["days"] == 1:
        log("info", "Data date: %s. Time taken: %.1f seconds." % (fmt_date(info["d_max"]), time.time() - t0))
    else:
        log("info", "Data from %s up to %s. Time taken: %.1f seconds." % (fmt_date(info["d_min"]), fmt_date(info["d_max"]), time.time() - t0))
    progress(1.0, "Excel report ready")
    return main_path, info


# =============================================================================
#  THE WINDOW
# =============================================================================

COL = {"bg": "#EEF3F9", "card": "#FFFFFF", "navy": "#0B2A4A", "blue": "#1565C0", "teal": "#00897B",
       "orange": "#EF6C00", "red": "#C62828", "green": "#1E9E5A", "ink": "#1B2733", "muted": "#5B6B7B",
       "line": "#D5DEE8", "logbg": "#0F1B2A", "logink": "#E6EDF5", "gold": "#FFB300",
       "sb": "#2A78D6", "nb": "#EB6834", "soft": "#F1F5FA", "amber": "#B26A00", "amberbg": "#FFF4D6",
       "greenbg": "#E3F6EA", "bluebg": "#E4EEFB", "redbg": "#FDE7E7", "greybg": "#E9EEF4"}
PART_COLOUR = {"NBPDCL": COL["nb"], "SBPDCL": COL["sb"], STATE: COL["teal"]}

BTN_WITH = "Create Report\nwith Dashboard Update"
BTN_WITHOUT = "Create Report\nwithout Dashboard Update"


def _shade(colour, by):
    """A lighter (by > 0) or darker (by < 0) shade of a #RRGGBB colour."""
    r, g, b = int(colour[1:3], 16), int(colour[3:5], 16), int(colour[5:7], 16)
    f = (lambda v: int(v + (255 - v) * by)) if by >= 0 else (lambda v: int(v * (1 + by)))
    return "#%02X%02X%02X" % (f(r), f(g), f(b))


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk

    try:                                   # sharp text on Windows
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title(APP_TITLE)
    root.configure(bg=COL["bg"])
    FONT = "Segoe UI"

    # ---- the window is sized to the screen: everything (text, buttons, gaps) grows or shrinks together ----
    BASE_W, BASE_H = 1040, 688                      # the size the window is drawn for, on a plain 1366 x 768 screen
    dpi = max(1.0, root.winfo_fpixels("1i") / 96.0)   # 1.0 at 100 %, 1.25 at 125 %, 1.5 at 150 % Windows display scale
    area = [0, 0, root.winfo_screenwidth(), root.winfo_screenheight() - int(50 * dpi)]      # left, top, width, height that may be used
    try:                                            # on Windows: the screen without the taskbar
        from ctypes import wintypes
        rect = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):
            area = [rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top]
    except Exception:
        pass
    room_w, room_h = area[2] - int(16 * dpi), area[3] - int(40 * dpi)        # less the window border and title bar
    fit = min(room_w / (BASE_W * dpi), room_h / (BASE_H * dpi))             # how much bigger the drawing could be and still fit
    zoom = fit * WINDOW_FILL if fit * WINDOW_FILL >= 1.0 else min(fit, 1.0)
    zoom = max(0.7, min(WINDOW_MAX_ZOOM, zoom))
    try:
        root.tk.call("tk", "scaling", float(root.tk.call("tk", "scaling")) * zoom)       # all text sizes follow
    except Exception:
        zoom = 1.0
    scale = dpi * zoom

    def S(n):
        """A length of the drawing, in screen dots."""
        return int(round(n * scale))
    win_w, win_h = min(S(BASE_W), room_w), min(S(BASE_H), room_h)
    root.geometry("%dx%d+%d+%d" % (win_w, win_h, area[0] + max(0, (area[2] - win_w) // 2), area[1] + max(0, (room_h - win_h) // 2)))
    root.minsize(min(win_w, S(900)), min(win_h, S(600)))

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", font=(FONT, 10), background=COL["bg"], foreground=COL["ink"])
    style.configure("TEntry", fieldbackground="#FFFFFF", padding=S(4))
    style.configure("Out.Horizontal.TProgressbar", troughcolor=COL["line"], background=COL["teal"],
                    bordercolor=COL["line"], lightcolor=COL["teal"], darkcolor=COL["teal"], thickness=S(14))

    def button(parent, text, colour, command, size=10, padx=S(12), pady=S(5), width=None):
        b = tk.Button(parent, text=text, command=command, bg=colour, fg="#FFFFFF", activebackground=_shade(colour, -0.18),
                      activeforeground="#FFFFFF", relief="flat", bd=0, cursor="hand2", disabledforeground="#F4F7FA",
                      font=(FONT, size, "bold"), padx=padx, pady=pady, highlightthickness=0)
        if width:
            b.configure(width=width)
        b.colour = colour
        b.bind("<Enter>", lambda e: str(b.cget("state")) != "disabled" and b.configure(bg=_shade(b.colour, -0.10)))
        b.bind("<Leave>", lambda e: str(b.cget("state")) != "disabled" and b.configure(bg=b.colour))
        return b

    msgs = queue.Queue()
    cancel_event = threading.Event()
    state = {"busy": False, "password": "", "last_report": None, "dashboard": None, "checks": 0}
    slots = {}                 # (DISCOM, company) -> list of file paths
    slot_label, part_info = {}, {}
    date_var, out_var = tk.StringVar(), tk.StringVar()
    date_typed = {"by_user": False}

    # ---------------------------------------------------------------- header
    head = tk.Canvas(root, height=S(64), highlightthickness=0, bd=0, bg=COL["navy"])
    head.pack(fill="x")
    more = tk.Menubutton(head, text="More  \u25BE", bg="#1B4C86", fg="#FFFFFF", activebackground=COL["blue"], activeforeground="#FFFFFF",
                         relief="flat", bd=0, font=(FONT, 10, "bold"), padx=S(14), pady=S(6), cursor="hand2", highlightthickness=0)
    menu = tk.Menu(more, tearoff=0, font=(FONT, 10))
    more.configure(menu=menu)

    def paint_head(event=None):
        w = max(head.winfo_width(), S(400))
        head.delete("paint")
        stops = [(0.0, (11, 42, 74)), (0.55, (21, 101, 192)), (1.0, (0, 137, 123))]
        steps = 90
        for i in range(steps):
            t = i / float(steps - 1)
            for (t0, c0), (t1, c1) in zip(stops, stops[1:]):
                if t0 <= t <= t1:
                    k = (t - t0) / (t1 - t0)
                    col = "#%02X%02X%02X" % tuple(int(c0[j] + (c1[j] - c0[j]) * k) for j in range(3))
                    break
            head.create_rectangle(int(w * i / steps), 0, int(w * (i + 1) / steps) + 1, S(60), fill=col, outline="", tags="paint")
        head.create_rectangle(0, S(60), w, S(64), fill=COL["gold"], outline="", tags="paint")
        head.create_text(S(22), S(21), text="\u26A1  " + APP_TITLE, anchor="w", fill="#FFFFFF", font=(FONT, 17, "bold"), tags="paint")
        head.create_text(S(24), S(45), text="Choose the files   \u2192   press one button   \u2192   Excel report and dashboard",
                         anchor="w", fill="#D6E6F7", font=(FONT, 10), tags="paint")
        head.coords("more", w - S(18), S(30))
        head.tag_raise("more")
    head.create_window(S(600), S(30), window=more, anchor="e", tags="more")
    head.bind("<Configure>", paint_head)

    body = tk.Frame(root, bg=COL["bg"])
    body.pack(fill="both", expand=True, padx=S(16), pady=(S(8), S(8)))

    def card(parent, number, title, colour, hint=""):
        outer = tk.Frame(parent, bg=COL["line"])
        inner = tk.Frame(outer, bg=COL["card"])
        inner.pack(fill="both", expand=True, padx=S(1), pady=S(1))
        if title:
            top = tk.Frame(inner, bg=COL["card"])
            top.pack(fill="x", padx=S(14), pady=(S(6), S(0)))
            tk.Label(top, text=" %s " % number, bg=colour, fg="#FFFFFF", font=(FONT, 11, "bold"), padx=S(4), pady=S(1)).pack(side="left")
            tk.Label(top, text="  " + title, bg=COL["card"], fg=COL["navy"], font=(FONT, 12, "bold")).pack(side="left")
            if hint:
                tk.Label(top, text=hint, bg=COL["card"], fg=COL["muted"], font=(FONT, 9)).pack(side="right")
        box = tk.Frame(inner, bg=COL["card"])
        box.pack(fill="both", expand=True, padx=S(14), pady=(S(6 if title else 10), S(8)))
        return outer, box

    # ---------------------------------------------------------------- 1. the files
    c1, f1 = card(body, "1", "Choose the outage files", COL["blue"], "Add each company's CSV file. Fill one DISCOM, or both.")
    c1.pack(fill="x")
    panels = tk.Frame(f1, bg=COL["card"])
    panels.pack(fill="x")

    def chosen_files():
        out = []
        for section in SECTION_ORDER:
            for company in COMPANIES.get(section, []):
                for p in slots[(section, company)]:
                    out.append((section, company, p))
        return out

    def short_path(p, n=58):
        return p if len(p) <= n else p[:18] + " ... " + p[-(n - 23):]

    def refresh_files():
        for section in SECTION_ORDER:
            names = COMPANIES.get(section, [])
            have = 0
            for company in names:
                got = slots[(section, company)]
                size = 0
                for p in got:
                    try:
                        size += os.path.getsize(p)
                    except OSError:
                        pass
                have += 1 if got else 0
                lab = slot_label[(section, company)]
                if not got:
                    lab.configure(text="no file chosen", fg="#8A97A6", bg=COL["soft"], font=(FONT, 9, "italic"))
                elif len(got) == 1:
                    lab.configure(text="\u2714  %s   (%s)" % (os.path.basename(got[0]), fmt_size(size)), fg="#0F5132", bg=COL["greenbg"], font=(FONT, 9, "bold"))
                else:
                    lab.configure(text="\u2714  %d files (%s):  %s" % (len(got), fmt_size(size), ",  ".join(os.path.basename(p) for p in got)),
                                  fg="#0F5132", bg=COL["greenbg"], font=(FONT, 9, "bold"))
            part_info[section].set("%d of %d companies" % (have, len(names)) if have else "not in this report")
        paths = [x[2] for x in chosen_files()]
        if not date_typed["by_user"]:
            g = guess_date(paths) or (datetime.date.today() - datetime.timedelta(days=1))
            date_var.set(g.strftime("%d.%m.%Y"))
        if paths and not out_var.get().strip():
            out_var.set(os.path.dirname(paths[0]))
        out_show.configure(text=short_path(out_var.get().strip()) or "(the folder of the first file you choose)")

    def browse(section, company):
        if state["busy"]:
            return
        got = filedialog.askopenfilenames(title="%s, %s: choose the outage CSV file(s)" % (section, company),
                                          filetypes=[("CSV files", "*.csv *.txt"), ("All files", "*.*")])
        already = dict((x[2], x) for x in chosen_files())
        for p in got:
            p = os.path.normpath(p)
            if p in already:
                if already[p][:2] != (section, company):
                    messagebox.showinfo(APP_TITLE, "%s is already chosen under %s, %s.\nIt was not added again."
                                        % (os.path.basename(p), already[p][0], already[p][1]))
                continue
            slots[(section, company)].append(p)
            already[p] = (section, company, p)
        refresh_files()

    def clear_slot(section, company):
        if not state["busy"]:
            del slots[(section, company)][:]
            refresh_files()

    def clear_all():
        if not state["busy"]:
            for k in slots:
                del slots[k][:]
            refresh_files()

    for ci, section in enumerate(SECTION_ORDER):
        panels.columnconfigure(ci, weight=1, uniform="part")
        colour = PART_COLOUR.get(section, COL["navy"])
        box = tk.Frame(panels, bg=colour)
        box.grid(row=0, column=ci, sticky="nsew", padx=(S(0), S(6)) if ci == 0 else (S(6), S(0)))
        bar = tk.Frame(box, bg=colour)
        bar.pack(fill="x")
        tk.Label(bar, text="  " + section, bg=colour, fg="#FFFFFF", font=(FONT, 12, "bold")).pack(side="left", pady=S(3))
        part_info[section] = tk.StringVar()
        tk.Label(bar, textvariable=part_info[section], bg=colour, fg="#FFFFFF", font=(FONT, 9)).pack(side="right", padx=S(10))
        inner = tk.Frame(box, bg=COL["card"])
        inner.pack(fill="both", expand=True, padx=S(2), pady=(S(0), S(2)))
        inner.columnconfigure(1, weight=1)
        for ri, company in enumerate(COMPANIES.get(section, [])):
            slots[(section, company)] = []
            tk.Label(inner, text=company, bg=COL["card"], fg=COL["ink"], font=(FONT, 10, "bold"), anchor="w",
                     width=11).grid(row=ri, column=0, sticky="w", padx=(S(10), S(4)), pady=S(2))
            lab = tk.Label(inner, text="no file chosen", bg=COL["soft"], fg="#8A97A6", font=(FONT, 9, "italic"), anchor="w", width=10, padx=S(8), pady=S(3))
            lab.grid(row=ri, column=1, sticky="ew", pady=S(2))
            slot_label[(section, company)] = lab
            button(inner, "Browse", colour, lambda s2=section, c2=company: browse(s2, c2), size=9, padx=S(10), pady=S(2)).grid(row=ri, column=2, padx=(S(6), S(3)), pady=S(2))
            x = button(inner, "\u2715", "#9AA7B5", lambda s2=section, c2=company: clear_slot(s2, c2), size=9, padx=S(7), pady=S(2))
            x.grid(row=ri, column=3, padx=(S(0), S(8)), pady=S(2))

    drow = tk.Frame(f1, bg=COL["card"])
    drow.pack(fill="x", pady=(S(7), S(0)))
    tk.Label(drow, text="Report date", bg=COL["card"], fg=COL["ink"], font=(FONT, 10, "bold")).pack(side="left")
    date_entry = ttk.Entry(drow, textvariable=date_var, width=12, font=(FONT, 10))
    date_entry.pack(side="left", padx=(S(8), S(0)))
    date_entry.bind("<Key>", lambda e: date_typed.update(by_user=True))
    tk.Label(drow, text="  the day the outage data belongs to", bg=COL["card"], fg=COL["muted"], font=(FONT, 9)).pack(side="left")

    def browse_out():
        d = filedialog.askdirectory(title="Choose the folder for the Excel reports")
        if d:
            out_var.set(os.path.normpath(d))
            refresh_files()
    button(drow, "Change", COL["muted"], browse_out, size=9, padx=S(9), pady=S(2)).pack(side="right")
    out_show = tk.Label(drow, text="", bg=COL["card"], fg=COL["ink"], font=(FONT, 9))
    out_show.pack(side="right", padx=(S(6), S(8)))
    tk.Label(drow, text="Excel reports are saved in:", bg=COL["card"], fg=COL["muted"], font=(FONT, 9)).pack(side="right")

    # ---------------------------------------------------------------- 2. the two buttons
    c2, f2 = card(body, "2", "Create the report", COL["green"])
    c2.pack(fill="x", pady=(S(8), S(0)))
    brow = tk.Frame(f2, bg=COL["card"])
    brow.pack(fill="x")
    brow.columnconfigure(0, weight=1, uniform="b")
    brow.columnconfigure(1, weight=1, uniform="b")
    with_btn = button(brow, BTN_WITH, COL["green"], lambda: create(True), size=13, pady=S(7))
    with_btn.grid(row=0, column=0, sticky="ew", padx=(S(0), S(6)))
    without_btn = button(brow, BTN_WITHOUT, COL["blue"], lambda: create(False), size=13, pady=S(7))
    without_btn.grid(row=0, column=1, sticky="ew", padx=(S(6), S(0)))
    tk.Label(brow, text="Makes the Excel report, sends it to the dashboard, and opens the dashboard.", bg=COL["card"], fg=COL["muted"],
             font=(FONT, 9)).grid(row=1, column=0, pady=(S(4), S(0)))
    tk.Label(brow, text="Makes the Excel report only. The dashboard stays as it is.", bg=COL["card"], fg=COL["muted"],
             font=(FONT, 9)).grid(row=1, column=1, pady=(S(4), S(0)))
    prow = tk.Frame(f2, bg=COL["card"])
    prow.pack(fill="x", pady=(S(7), S(0)))
    cancel_btn = button(prow, "Stop", COL["red"], lambda: cancel(), size=9, padx=S(12), pady=S(2))
    prog = ttk.Progressbar(prow, style="Out.Horizontal.TProgressbar", maximum=1000)
    prog.pack(side="left", fill="x", expand=True)

    # ---------------------------------------------------------------- 3. the result
    c3, f3 = card(body, "", "", COL["teal"])
    c3.pack(fill="both", expand=True, pady=(S(8), S(0)))
    banner = tk.Label(f3, text="", anchor="w", font=(FONT, 11, "bold"), padx=S(12), pady=S(5), justify="left")
    banner.pack(fill="x")
    tiles = tk.Frame(f3, bg=COL["card"])
    tiles.pack(fill="x", pady=(S(8), S(0)))
    tile = {}
    for ti, name in enumerate([STATE, "SBPDCL", "NBPDCL"]):
        tiles.columnconfigure(ti, weight=1, uniform="t")
        colour = PART_COLOUR[name]
        t = tk.Frame(tiles, bg=colour)
        t.grid(row=0, column=ti, sticky="nsew", padx=(S(0 if ti == 0 else 5), S(0 if ti == 2 else 5)))
        top = tk.Frame(t, bg=colour)
        top.pack(fill="x", padx=S(10), pady=(S(6), S(0)))
        head_l = tk.Label(top, text="Bihar" if name == STATE else name, bg=colour, fg="#FFFFFF", font=(FONT, 11, "bold"), anchor="w")
        head_l.pack(side="left")
        l3 = tk.Label(top, text=" ", bg=colour, fg="#FFFFFF", font=(FONT, 9), anchor="e")
        l3.pack(side="right")
        line = tk.Frame(t, bg=colour)
        line.pack(fill="x", padx=S(10), pady=(S(0), S(6)))
        big = tk.Label(line, text="\u2013", bg=colour, fg="#FFFFFF", font=(FONT, 22, "bold"), anchor="w")
        big.pack(side="left")
        side = tk.Frame(line, bg=colour)
        side.pack(side="left", fill="x", expand=True, padx=(S(12), S(0)))
        l1 = tk.Label(side, text="average outage per feeder", bg=colour, fg="#FFFFFF", font=(FONT, 9), anchor="w")
        l1.pack(fill="x")
        l2 = tk.Label(side, text=" ", bg=colour, fg="#FFFFFF", font=(FONT, 9, "bold"), anchor="w")
        l2.pack(fill="x")
        tile[name] = ((t, top, head_l, line, big, side, l1, l2, l3), colour)
    foot = tk.Frame(f3, bg=COL["card"])
    foot.pack(fill="x", pady=(S(8), S(0)))
    check_lab = tk.Label(foot, text="", bg=COL["card"], fg=COL["amber"], font=(FONT, 10, "bold"), anchor="w", cursor="hand2")
    check_lab.pack(side="left")

    def set_banner(kind, text):
        bg, fg = {"idle": (COL["bluebg"], "#0B3C7A"), "work": (COL["bluebg"], "#0B3C7A"), "ok": (COL["greenbg"], "#0F5132"),
                  "made": (COL["bluebg"], "#0B3C7A"), "warn": (COL["amberbg"], "#7A4B00"), "bad": (COL["redbg"], "#8E1B1B"),
                  "stop": (COL["greybg"], COL["ink"])}[kind]
        banner.configure(text=text, bg=bg, fg=fg)

    def set_tiles(rows):
        """rows: {name: (average outage text, interruptions per feeder, feeders with outage, total feeders)} or None to empty."""
        for name, (parts, colour) in tile.items():
            big, l1, l2, l3 = parts[4], parts[6], parts[7], parts[8]
            r = (rows or {}).get(name)
            bg = colour if (r or rows is None) else "#B8C2CE"
            for w in parts:
                w.configure(bg=bg)
            if r:
                big.configure(text=r[0])
                l1.configure(text="average outage per feeder")
                l2.configure(text="%s interruptions per feeder" % r[1])
                l3.configure(text="%s of %s feeders had an outage" % (r[2], r[3]))
            else:
                big.configure(text="\u2013")
                l1.configure(text="average outage per feeder" if rows is None else "not in this report")
                l2.configure(text=" ")
                l3.configure(text=" ")

    # ---------------------------------------------------------------- details window (the log)
    det = tk.Toplevel(root)
    det.title("Details of the work")
    det.geometry("%dx%d" % (min(S(900), room_w), min(S(460), room_h)))
    det.configure(bg=COL["bg"])
    det.withdraw()
    det.protocol("WM_DELETE_WINDOW", det.withdraw)
    dbar = tk.Frame(det, bg=COL["bg"])
    dbar.pack(fill="x", side="bottom", padx=S(10), pady=S(8))
    dbody = tk.Frame(det, bg=COL["bg"])
    dbody.pack(fill="both", expand=True, padx=S(10), pady=(S(10), S(0)))
    log_box = tk.Text(dbody, wrap="word", bg=COL["logbg"], fg=COL["logink"], insertbackground="#FFFFFF",
                      font=("Consolas", 10), bd=0, padx=S(10), pady=S(8), state="disabled")
    lsb = ttk.Scrollbar(dbody, orient="vertical", command=log_box.yview)
    log_box.configure(yscrollcommand=lsb.set)
    lsb.pack(side="right", fill="y")
    log_box.pack(side="left", fill="both", expand=True)
    log_box.tag_configure("info", foreground=COL["logink"])
    log_box.tag_configure("ok", foreground="#7CE0A3")
    log_box.tag_configure("warn", foreground="#FFD166")
    log_box.tag_configure("error", foreground="#FF8A80")
    log_box.tag_configure("time", foreground="#7F93A8")

    def show_details():
        det.deiconify()
        det.lift()
    check_lab.bind("<Button-1>", lambda e: show_details())

    def log_now(kind, text):
        log_box.configure(state="normal")
        log_box.insert("end", time.strftime("%H:%M:%S  "), "time")
        mark = {"ok": "OK     ", "warn": "CHECK  ", "error": "ERROR  "}.get(kind, "       ")
        log_box.insert("end", mark + text + "\n", kind)
        log_box.see("end")
        log_box.configure(state="disabled")
        if kind == "warn":
            state["checks"] += 1

    def log(kind, text):                       # safe from the worker thread
        msgs.put(("log", kind, text))

    def progress(frac, text):
        msgs.put(("progress", frac, text))

    def log_text():
        return log_box.get("1.0", "end-1c")

    def copy_log():
        root.clipboard_clear()
        root.clipboard_append(log_text())
        messagebox.showinfo(APP_TITLE, "The details are copied. Paste them with Ctrl+V.", parent=det)

    def save_log():
        p = filedialog.asksaveasfilename(title="Save the details", defaultextension=".txt", parent=det,
                                         initialfile="outage_report_log_%s.txt" % time.strftime("%Y%m%d_%H%M%S"),
                                         filetypes=[("Text file", "*.txt")])
        if not p:
            return
        try:
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(log_text())
        except OSError as e:
            messagebox.showerror(APP_TITLE, "The details could not be saved here:\n%s\n(%s)\nChoose another folder." % (p, e), parent=det)

    button(dbar, "Copy", COL["navy"], copy_log, width=10).pack(side="left")
    button(dbar, "Save as a file", COL["navy"], save_log).pack(side="left", padx=S(8))
    button(dbar, "Close", COL["muted"], det.withdraw, width=10).pack(side="right")

    # ---------------------------------------------------------------- small actions
    def open_folder():
        d = out_var.get().strip()
        if not os.path.isdir(d):
            messagebox.showinfo(APP_TITLE, "No folder yet. Choose a file first, or press 'Change' to choose the folder.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(d)
            else:
                webbrowser.open("file://" + d)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "The folder could not be opened:\n%s\n(%s)" % (d, e))

    def open_dashboard():
        link = state["dashboard"] or (WEB_APP_URL if web_link_ready() else "")
        if not link:
            messagebox.showinfo(APP_TITLE, "The dashboard link is not yet written in the app. See the set-up notes, Part B.")
            return
        try:
            webbrowser.open(link)
        except Exception:
            messagebox.showinfo(APP_TITLE, "The browser could not be opened. Open this link yourself:\n%s" % link)

    def feeder_list():
        p = os.path.join(app_folder(), MASTER_FILE_NAME)
        try:
            if not os.path.isfile(p):
                write_master_file(p)
                log_now("ok", "Feeder list file made: %s" % p)
            messagebox.showinfo(APP_TITLE, "The feeder list is kept next to the app:\n\n%s\n   the total number of feeders in each division\n\n"
                                           "Open it in Excel, change it, and save it as CSV with the same name. "
                                           "Delete the file to go back to the list inside the app.\n\nFolder: %s" % (MASTER_FILE_NAME, app_folder()))
            if sys.platform.startswith("win"):
                os.startfile(app_folder())
        except Exception as e:
            messagebox.showerror(APP_TITLE, "The feeder list file could not be made here:\n%s\n(%s)\n"
                                            "Keep the app in a folder where you can save files, such as Documents." % (app_folder(), e))

    def ask_password():
        """The upload password is asked the first time it is needed after the app is opened. It is never saved."""
        if state["password"]:
            return state["password"]
        pw = simpledialog.askstring(APP_TITLE, "Upload password for the dashboard.\n\n"
                                               "It is asked once each time the app is opened and is never saved.", show="\u2022", parent=root) or ""
        state["password"] = pw
        return pw

    def set_busy(on):
        state["busy"] = on
        for b in (with_btn, without_btn):
            b.configure(state="disabled" if on else "normal", bg=_shade(b.colour, 0.55) if on else b.colour, cursor="arrow" if on else "hand2")
        if on:
            cancel_btn.pack(side="right", padx=(S(10), S(0)))
        else:
            cancel_btn.pack_forget()

    def worker(job):
        try:
            job()
        except Cancelled:
            log("warn", "Stopped by you. No report was made and the dashboard was not changed.")
            msgs.put(("done", "stop", "Stopped by you. Nothing was made or sent."))
        except UserError as e:
            log("error", str(e))
            msgs.put(("fail", str(e)))
        except MemoryError:
            m = "The computer ran out of memory. Close other programs and try again, or read fewer files at a time."
            log("error", m)
            msgs.put(("fail", m))
        except Exception as e:
            log("error", "Something unexpected went wrong: %s. Please save the details and send them to the person who looks after this app.\n%s"
                % (e, traceback.format_exc()))
            msgs.put(("fail", "Something unexpected went wrong: %s\n\nPress More > Details of the work, save the details, "
                              "and send them to the person who looks after this app." % e))

    def start(job, text):
        cancel_event.clear()
        state["checks"] = 0
        check_lab.configure(text="")
        set_busy(True)
        set_banner("work", text)
        prog.configure(mode="determinate", value=0)
        threading.Thread(target=worker, args=(job,), daemon=True).start()

    def figures(info):
        t = info["summary"]
        out = {}
        for row in t.rows:
            out[row[0]] = (fmt_hhmm(row[t.col("Avg Outage per Feeder (HH:MM)")]), "%.2f" % row[t.col("Avg Interruption per Feeder")],
                           "{:,}".format(row[t.col("Feeders With Outage")]), "{:,}".format(row[t.col("Total Feeders")]))
        if len(out) == 2 and STATE in out:           # one DISCOM only: the Bihar line would only repeat it
            del out[STATE]
        return out

    def create(send):
        if state["busy"]:
            return
        text = date_var.get().strip()
        d = parse_date(text)
        problem = None
        inputs, folder = chosen_files(), out_var.get().strip()
        if not inputs:
            problem = "No file is chosen yet.\n\nPress 'Browse' next to a company name and choose its CSV file."
        elif d is None:
            problem = "The report date '%s' is not a date.\n\nType it as DD.MM.YYYY, for example 10.10.2026." % text
        elif d > datetime.date.today():
            problem = "The report date %s is in the future.\n\nCorrect it and press the button again." % d.strftime("%d.%m.%Y")
        elif send and not web_link_ready():
            problem = ("The dashboard link is not yet written in the app, so the dashboard cannot be updated.\n\n"
                       "Press 'Create Report without Dashboard Update', or see the set-up notes, Part B.")
        if problem:
            messagebox.showerror(APP_TITLE, problem)
            return
        missing = []
        for section in SECTION_ORDER:
            names = COMPANIES.get(section, [])
            if any(slots[(section, c)] for c in names):
                empty = [c for c in names if not slots[(section, c)]]
                if empty:
                    missing.append("%s:  %s" % (section, ", ".join(empty)))
        if missing and not messagebox.askyesno(
                APP_TITLE, "No file is chosen for:\n\n     %s\n\nTheir divisions will be counted as having NO outage, so the averages "
                           "will look better than they are.\n\nMake the report anyway?" % "\n     ".join(missing)):
            return
        pw = ""
        if send:
            pw = ask_password()
            if not pw:
                messagebox.showinfo(APP_TITLE, "No password was given, so nothing was done.\n\n"
                                               "To make the Excel report only, press 'Create Report without Dashboard Update'.")
                return
        log_now("info", "---- Started: %d file(s), %s ----" % (len(inputs), "with dashboard update" if send else "without dashboard update"))
        for m in missing:
            log_now("warn", "No file was chosen for %s. Their divisions are counted as no outage." % m)

        def job():
            main_path, info = make_report(inputs, folder, d, True, log, progress, cancel_event)
            state["last_report"] = info["parts"][0][1] if info["parts"] else main_path
            day = fmt_date(info["d_max"])
            msgs.put(("tiles", figures(info)))
            if not send:
                msgs.put(("done", "made", "\u2714  Excel report of %s is ready.   The dashboard was not updated." % day))
                return
            if cancel_event.is_set():
                raise Cancelled()
            if not info["parts"]:
                raise UserError("The Excel report was made, but it holds no SBPDCL or NBPDCL rows, so the dashboard was not updated.")
            names = [x[0] for x in info["parts"]]
            if info["kept"]:
                log("info", "%s was not read in this run, so it is not sent. Its figures on the dashboard stay as they are." % " and ".join(info["kept"]))
            try:
                link = send_parts(info["parts"], pw, log, progress)
            except UserError as e:
                if "password" in str(e).lower():
                    msgs.put(("forget",))
                if getattr(e, "link", None):
                    msgs.put(("open", e.link))
                raise UserError("The Excel report of %s was made and saved, but the dashboard was NOT fully updated.\n\n%s" % (day, e))
            log("ok", "Opening the dashboard in the browser.")
            msgs.put(("open", link))
            msgs.put(("done", "ok", "\u2714  Excel report of %s is ready.   Dashboard updated for %s. It opens in the browser." % (day, " and ".join(names))))
        set_tiles(None)
        start(job, "Working ...")

    def send_existing():
        if state["busy"]:
            return
        if not web_link_ready():
            messagebox.showinfo(APP_TITLE, "The dashboard link is not yet written in the app. See the set-up notes, Part B.")
            return
        start_dir = out_var.get().strip() if os.path.isdir(out_var.get().strip()) else None
        p = filedialog.askopenfilename(title="Choose the DISCOM report to send (file ending with SBPDCL.xlsx or NBPDCL.xlsx)", initialdir=start_dir,
                                       initialfile=os.path.basename(state["last_report"]) if state["last_report"] else None,
                                       filetypes=[("Excel report", "*.xlsx")])
        if not p:
            return
        pw = ask_password()
        if not pw:
            return
        log_now("info", "---- Sending the report file: %s ----" % p)

        def job():
            try:
                link = send_report(p, pw, log, progress)
            except UserError as e:
                if "password" in str(e).lower():
                    msgs.put(("forget",))
                raise
            log("ok", "Opening the dashboard in the browser.")
            msgs.put(("open", link))
            msgs.put(("done", "ok", "\u2714  Dashboard updated from %s. It opens in the browser." % os.path.basename(p)))
        start(job, "Sending ...")

    def cancel():
        if state["busy"]:
            cancel_event.set()
            set_banner("work", "Stopping ...")

    menu.add_command(label="Send a report made earlier ...", command=send_existing)
    menu.add_command(label="Open the folder of the reports", command=open_folder)
    menu.add_command(label="Open the dashboard", command=open_dashboard)
    menu.add_separator()
    menu.add_command(label="Details of the work (log)", command=show_details)
    menu.add_command(label="Feeder list (total feeders) ...", command=feeder_list)
    menu.add_command(label="Agriculture list (opens the web page) ...", command=lambda: webbrowser.open(LIST_PAGE_URL))
    menu.add_command(label="Clear all files", command=clear_all)
    button(foot, "Open the dashboard", COL["teal"], open_dashboard, size=9, padx=S(10), pady=S(3)).pack(side="right")
    button(foot, "Open the folder", COL["blue"], open_folder, size=9, padx=S(10), pady=S(3)).pack(side="right", padx=S(8))

    def show_checks():
        n = state["checks"]
        check_lab.configure(text=("\u26A0  %d point%s to check  \u2013  click here to read" % (n, "" if n == 1 else "s")) if n else "")

    def pump():
        try:
            while True:
                m = msgs.get_nowait()
                if m[0] == "log":
                    log_now(m[1], m[2])
                elif m[0] == "progress":
                    if m[1] is None:
                        if str(prog.cget("mode")) != "indeterminate":
                            prog.configure(mode="indeterminate")
                            prog.start(12)
                    else:
                        if str(prog.cget("mode")) != "determinate":
                            prog.stop()
                            prog.configure(mode="determinate")
                        prog.configure(value=int(1000 * max(0.0, min(1.0, m[1]))))
                    if state["busy"]:
                        set_banner("work", m[2] + " ...")
                elif m[0] == "tiles":
                    set_tiles(m[1])
                elif m[0] == "forget":
                    state["password"] = ""                  # a wrong password is asked again next time
                elif m[0] == "open":
                    state["dashboard"] = m[1]
                    try:
                        webbrowser.open(m[1])
                    except Exception:
                        log_now("warn", "The browser could not be opened. Open this link yourself: %s" % m[1])
                elif m[0] in ("done", "fail"):
                    prog.stop()
                    prog.configure(mode="determinate")
                    set_busy(False)
                    show_checks()
                    if m[0] == "fail":
                        prog.configure(value=0)
                        first = m[1].split("\n")[0]
                        set_banner("bad", "\u2716  " + (first if len(first) < 120 else first[:117] + "..."))
                        messagebox.showerror(APP_TITLE, m[1])
                    else:
                        prog.configure(value=1000 if m[1] != "stop" else 0)
                        set_banner(m[1], m[2])
        except queue.Empty:
            pass
        root.after(100, pump)

    def on_close():
        if state["busy"]:
            if not messagebox.askyesno(APP_TITLE, "Work is still going on. Stop it and close?"):
                return
            cancel_event.set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    refresh_files()
    set_tiles(None)
    set_banner("idle", "Choose the files above, then press one of the two buttons.")
    log_now("info", "%s, version %s." % (APP_TITLE, APP_VERSION))
    try:
        m = load_master()
        log_now("info", "Feeder numbers come from the %s (%d divisions)." % (m.source, len(m.divisions)))
        tl = load_type_list(fetch=False)
        log_now("info", "Agriculture list: %d feeder(s) in the %s. The latest shared list is read from Google before every report."
                % (tl.count, tl.source))
    except UserError as e:
        log_now("error", str(e))
        set_banner("bad", "\u2716  " + str(e).split("\n")[0])
    state["checks"] = 0
    root.update_idletasks()
    need_w, need_h = root.winfo_reqwidth(), root.winfo_reqheight()
    if need_w > win_w or need_h > win_h:
        win_w, win_h = min(max(win_w, need_w), room_w), min(max(win_h, need_h), room_h)
        root.geometry("%dx%d+%d+%d" % (win_w, win_h, area[0] + max(0, (area[2] - win_w) // 2), area[1] + max(0, (room_h - win_h) // 2)))
    root.after(100, pump)
    root.mainloop()


# =============================================================================
#  COMMAND LINE (for testing or for running without the window)
#     python outage_report_app.py --cli --out FOLDER --date 10.10.2026 file1.csv file2.csv
# =============================================================================


def run_cli(argv):
    import argparse
    ap = argparse.ArgumentParser(description=APP_TITLE)
    ap.add_argument("--cli", action="store_true")
    ap.add_argument("--out", required=True)
    ap.add_argument("--date", default="")
    ap.add_argument("--nb", action="append", default=[], metavar="COMPANY=FILE", help="an NBPDCL file, for example --nb Adani=a.csv")
    ap.add_argument("--sb", action="append", default=[], metavar="COMPANY=FILE", help="an SBPDCL file, for example --sb Genus=b.csv")
    ap.add_argument("--no-keep", action="store_true", help="do not take the other DISCOM from a report made earlier")
    ap.add_argument("--split", action="store_true", help="no longer needed: the DISCOM files are always made")
    ap.add_argument("--send", action="store_true")
    ap.add_argument("--password", default="")
    ap.add_argument("files", nargs="*")
    a = ap.parse_args(argv)

    def log(kind, text):
        print("%-6s %s" % ({"ok": "OK", "warn": "CHECK", "error": "ERROR"}.get(kind, ""), text))
    inputs = []
    for section, items in (("NBPDCL", a.nb), ("SBPDCL", a.sb)):
        for item in items:
            company, _, path = item.partition("=")
            inputs.append((section, company, path) if path else (section, "Not stated", company))
    inputs += list(a.files)
    paths = [x[2] for x in as_inputs(inputs)]
    d = parse_date(a.date) if a.date else (guess_date(paths) or datetime.date.today() - datetime.timedelta(days=1))
    if d is None:
        print("ERROR  The date '%s' is not a date. Use DD.MM.YYYY." % a.date)
        return 2
    try:
        main_path, info = make_report(inputs, a.out, d, not a.no_keep, log, lambda f, t: None, threading.Event())
        if a.send:
            print("OK     Dashboard:", send_parts(info["parts"], a.password, log))
    except UserError as e:
        print("ERROR ", e)
        return 1
    return 0


if __name__ == "__main__":
    if "--cli" in sys.argv[1:]:
        sys.exit(run_cli(sys.argv[1:]))
    run_gui()
