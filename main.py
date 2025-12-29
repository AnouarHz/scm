import os
import re
import time
import json
import random
import string
import builtins
import traceback
import win32com.client as win32
from pathlib import Path

import requests
import pandas as pd
from playwright.sync_api import sync_playwright

# ==========================================
# 1) PROJEKT-ROOT & DATEIEN
# ==========================================
BASE_DIR = Path(__file__).resolve().parent
ROOT = BASE_DIR / "ROOT"

FOLDER_IN = ROOT / "Antworten_Erhalt"
FOLDER_FINAL = ROOT / "Einzelberichte_Lieferanten"
FOLDER_NUTZ = ROOT / "Nutzwertanalyse"

FILE_SUPPLIERS = ROOT / "1. SCM-Anwendungen(MA)_Lieferantenuebersicht.xlsx"
FILE_ERP = ROOT / "4. SCM-Anwendungen(MA)_ERP-System.xlsx"
FILE_SCALE = ROOT / "3. SCM-Anwendungen(MA)_Gesamtbewertung.xlsx"
FILE_NUTZ_TEMPLATE = ROOT / "5. SCM-Nutzwertanalyse.xlsx"

for d in [FOLDER_IN, FOLDER_FINAL, FOLDER_NUTZ]:
    d.mkdir(parents=True, exist_ok=True)

# ==========================================
# 2) SERVER (Form Backend)
# ==========================================
FORM_SERVER = os.getenv("SCM_FORM_SERVER", "http://localhost:8000")

# ==========================================
# 3) OUTLOOK FINAL MAIL
# ==========================================
SEND_TO_FINAL = "anouar97@gmx.de"

# ==========================================
# 4) RUN / STATE
# ==========================================
ROUND_ID = "".join(random.choices(string.digits, k=8))
STATE_FILE = ROOT / f"round_state_{ROUND_ID}.json"
FILE_NUTZ_OUT = FOLDER_NUTZ / f"Nutzwertanalyse_R{ROUND_ID}.xlsx"

# Felder, bei denen nach Punkteberechnung der Original-ERP-Text im Bericht stehen soll:
KEEP_ORIGINAL_TEXT_FIELDS = {"co2-emissionen", "zahlungsbedingungen"}  # lower-case keys


# ==========================================
# STATE HELPERS
# ==========================================
def load_state():
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except:
            pass
    return {
        "round_id": ROUND_ID,
        "sent": {},         # supplier_id -> {name,email,sent_at}
        "responses": {},    # supplier_id -> {submitted_at, filename}
        "nutzwert_done": False,
        "final_mail_sent": False
    }

def save_state(state):
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

def status_line(state):
    total = len(state.get("sent", {}))
    got = len(state.get("responses", {}))
    return f"[STATUS] {got} von {total} Antworten"

def all_done(state) -> bool:
    total = len(state.get("sent", {}))
    got = len(state.get("responses", {}))
    return total > 0 and got >= total


# ==========================================
# SCALE LOADER
# ==========================================
def get_comprehensive_scale():
    df = pd.read_excel(FILE_SCALE, sheet_name="Skala", header=None)
    scale_map = {}
    for i in range(4, len(df)):
        crit = builtins.str(df.iloc[i, 0]).strip()
        if not crit or crit.lower() in ("nan", "none"):
            continue
        scale_map[crit] = {
            0: builtins.str(df.iloc[i, 4]),
            20: builtins.str(df.iloc[i, 5]),
            40: builtins.str(df.iloc[i, 6]),
            60: builtins.str(df.iloc[i, 7]),
            80: builtins.str(df.iloc[i, 8]),
            100: builtins.str(df.iloc[i, 9]),
        }
    return scale_map

def find_matching_criterion(crit_name, scale_data):
    c_clean = builtins.str(crit_name).strip().lower()
    if crit_name in scale_data:
        return crit_name
    for k in scale_data:
        kk = builtins.str(k).strip().lower()
        if kk == c_clean or c_clean in kk or kk in c_clean:
            return k
    return None


# ==========================================
# ERP VALUE NORMALIZATION / INTERVAL PARSING
# ==========================================
def _norm_text(s) -> str:
    s = "" if s is None else builtins.str(s)
    s = s.replace("\u00A0", " ")
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def extract_first_number(val):
    """Extract first numeric token from text (handles 11 000, 11.000, 98,5, kg CO2e, % etc.)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)

    s = builtins.str(val)

    # normalize whitespace and common units
    s = s.replace("\u00A0", " ")
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = s.lower()

    # remove units / labels (keep numbers)
    s = s.replace("kg co2e", "").replace("kgco2e", "").replace("co2e", "").replace("kg", "").replace("%", "")
    s = s.strip()

    # IMPORTANT: handle German thousand separators with dots (e.g. "11.000" -> "11000")
    # We only remove dots when they are used as thousand separators: 1-3 digits + (.xxx)+
    s_compact = s.replace(" ", "")
    m_th = re.search(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", s_compact)
    if m_th:
        token = m_th.group()
        token = token.replace(".", "")      # thousands dots out
        token = token.replace(",", ".")     # decimal comma to dot
        try:
            return float(token)
        except:
            pass

    # general fallback: remove spaces for "11 000" and parse first number
    s2 = s.replace(" ", "")
    s2 = s2.replace(",", ".")
    m = re.search(r"[-+]?\d*\.\d+|\d+", s2)
    if not m:
        return None
    try:
        return float(m.group())
    except:
        return None

def normalize_percent_if_needed(x):
    """
    Prozent-Normalisierung:
      Wenn 0 <= x <= 1 -> x*100
    """
    if isinstance(x, (int, float)):
        if 0 <= x <= 1:
            return round(x * 100, 2)
        return round(float(x), 2)
    return x

def parse_scale_condition(scale_text: str, erp_numeric: float) -> bool:
    """
    Mathematische Prüfung für Skalen:
      - Intervalle "11 000 – 13 999" oder "95-97,99%"
      - Operatoren: ≥, >=, <, <=
    """
    if erp_numeric is None:
        return False

    t = _norm_text(scale_text)
    if not t or t in ("nan", "none"):
        return False

    # cleanup numeric parsing
    raw = t
    raw = raw.replace("%", "")
    raw = raw.replace("kg co2e", "").replace("kgco2e", "").replace("co2e", "").replace("kg", "")
    raw = raw.replace(" ", "")
    raw = raw.replace(",", ".")  # 97,99 -> 97.99

    try:
        # <=
        if "<=" in raw:
            n = extract_first_number(raw.split("<=")[1])
            return n is not None and erp_numeric <= n

        # >=
        if ">=" in raw:
            n = extract_first_number(raw.split(">=")[1])
            return n is not None and erp_numeric >= n

        # Unicode ≥
        if "≥" in raw:
            n = extract_first_number(raw.split("≥")[1])
            return n is not None and erp_numeric >= n

        # Unicode ≤
        if "≤" in raw:
            n = extract_first_number(raw.split("≤")[1])
            return n is not None and erp_numeric <= n

        # <
        # (Achtung: nicht mit <= kollidieren, oben schon behandelt)
        if "<" in raw:
            n = extract_first_number(raw.split("<")[1])
            return n is not None and erp_numeric < n

        # >
        if ">" in raw:
            n = extract_first_number(raw.split(">")[1])
            return n is not None and erp_numeric > n

        # Intervall: – oder -
        if "–" in t or "-" in t:
            parts = re.split(r"[–-]", t)
            if len(parts) >= 2:
                low = extract_first_number(parts[0])
                high = extract_first_number(parts[1])
                if low is not None and high is not None:
                    return low <= erp_numeric <= high

    except:
        return False

    return False


def is_negative_presence_value(val) -> bool:
    """
    Fix ISO/Presence:
    - 'nicht vorhanden', 'nein', 'no', 0, 'false' => negativ
    """
    s = _norm_text(val)
    if s in ("", "nan", "none"):
        return True
    negatives = ["nicht vorhanden", "nichtvorhanden", "nein", "no", "false", "0", "keine", "kein"]
    return any(n in s for n in negatives)


def map_erp_to_points(kriterium: str, erp_value, scale_data: dict) -> int:
    """
    Robustes ERP-Mapping nach deinen Fix-Regeln:
    - Prozent normalisieren (0..1 -> *100)
    - Intervalle/Operatoren mathematisch prüfen
    - bidirektionales Text-Matching
    - Priorität 100->0
    - ISO Fix: 'nicht vorhanden' => 0 (nicht 100)
    - CoC Fix: BME soll als 100 erkannt werden, wenn Skala "KB oder BME"
    """
    actual = find_matching_criterion(kriterium, scale_data)
    if not actual:
        return 0

    crit_l = _norm_text(actual)

    # ISO Fix: falls ERP Wert negativ, direkt 0
    if crit_l.startswith("iso ") or "iso " in crit_l:
        if is_negative_presence_value(erp_value):
            return 0

    # numeric path
    num = extract_first_number(erp_value)
    if num is not None:
        num = normalize_percent_if_needed(num)

        for pts in [100, 80, 60, 40, 20, 0]:
            st = scale_data[actual].get(pts, "")
            if parse_scale_condition(st, num):
                return pts

    # text path (Incoterms/EDI/Zahlungsbedingungen/CoC etc.)
    val_str = _norm_text(erp_value)

    for pts in [100, 80, 60, 40, 20, 0]:
        st = _norm_text(scale_data[actual].get(pts, ""))

        if not st or st in ("nan", "none"):
            continue

        # bidirektional
        if val_str and (val_str in st or st in val_str):
            return pts

        # CoC Spezial: wenn ERP "bme" und Skala enthält "bme"
        if "code of conduct" in crit_l or crit_l == "coc" or "coc" in crit_l:
            if "bme" in val_str and "bme" in st and pts == 100:
                return 100
            if "kb" in val_str and "kb" in st and pts == 100:
                return 100

    return 0


# ==========================================
# SERVER API CALLS
# ==========================================
def get_form_link(supplier_id: str, round_id: str) -> str:
    r = requests.get(f"{FORM_SERVER}/issue-link", params={"supplier_id": supplier_id, "round_id": round_id}, timeout=10)
    r.raise_for_status()
    return r.json()["url"]

def list_submissions(round_id: str):
    r = requests.get(f"{FORM_SERVER}/api/submissions", params={"round_id": round_id}, timeout=10)
    r.raise_for_status()
    return r.json()  # [{supplier_id, submitted_at}]

def download_submission_xlsx(round_id: str, supplier_id: str) -> bytes:
    r = requests.get(f"{FORM_SERVER}/api/xlsx", params={"round_id": round_id, "supplier_id": supplier_id}, timeout=30)
    r.raise_for_status()
    return r.content


# ==========================================
# OUTLOOK SEND FINAL MAIL
# ==========================================
def send_final_mail_outlook(page, to_email: str, subject: str, body: str, attachment_path: Path):
    page.goto("https://outlook.office.com/mail/")
    page.wait_for_selector('button[aria-label*="Neue"]', timeout=60000)

    page.click('button[aria-label*="Neue"]')
    page.wait_for_timeout(1000)

    page.fill('div[aria-label="An"]', to_email)
    page.fill('input[placeholder*="Betreff"]', subject)
    page.keyboard.type(body)

    page.locator('button[aria-label*="Datei anfügen"]').first.click()
    with page.expect_file_chooser() as fc:
        page.locator('button[aria-label*="Diesen Computer durchsuchen"]').first.click()
    fc.value.set_files(builtins.str(attachment_path.resolve()))

    page.wait_for_timeout(2000)
    page.click('button[aria-label*="Senden"]')
    page.wait_for_selector('div[aria-label="An"]', state="hidden", timeout=30000)


# ==========================================
# DISPATCH (send link mails)
# ==========================================
def run_rpa_dispatch(page, round_id, state):
    print(f"\n[PHASE 1] Versand Runde {round_id} gestartet...")
    df_supp = pd.read_excel(FILE_SUPPLIERS, sheet_name="Lieferanten", header=2).dropna(subset=["Lieferant_Name"])

    page.goto("https://outlook.office.com/mail/")
    page.wait_for_selector('button[aria-label*="Neue"]', timeout=60000)

    for _, row in df_supp.iterrows():
        try:
            s_id = builtins.str(row["Lieferant_ID"]).strip()
            email = builtins.str(row["Email"]).strip()
            name = builtins.str(row.get("Name", "")).strip()
            lname = builtins.str(row.get("Lieferant_Name", "")).strip()

            form_url = get_form_link(s_id, round_id)

            page.click('button[aria-label*="Neue"]')
            page.fill('div[aria-label="An"]', email)
            page.fill('input[placeholder*="Betreff"]', f"SCM-Bewertung | Runde {round_id} | {s_id}")
            page.keyboard.press("Tab")
            msg = (
                f"Hallo {name},\n\n"
                f"bitte füllen Sie die Lieferantenbewertung über diesen Link aus:\n\n"
                f"{form_url}\n\n"
                f"Vielen Dank!\n\n"
                f"(Runde {round_id})"
            )
            page.keyboard.type(msg)
            page.click('button[aria-label*="Senden"]')
            page.wait_for_selector('div[aria-label="An"]', state="hidden", timeout=15000)

            state["sent"][s_id] = {"name": lname, "email": email, "sent_at": time.time()}
            save_state(state)

            print(f" [OK] Link an {lname} ({s_id})")
            print(status_line(state))
        except:
            page.keyboard.press("Escape")


# ==========================================
# REPORT GENERATION
# ==========================================
def safe_str(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return builtins.str(x).strip()

def process_supplier_from_xlsx(file_path: Path, round_id: str, supplier_id: str, scale_data: dict):
    """
    file_path enthält Kriterium/Bewertung aus Formular.
    Erstellt/aktualisiert Einzelbericht.
    """
    try:
        df_man = pd.read_excel(file_path)
        if df_man is None or df_man.empty:
            print(" [!] Manuelle Antwortdatei leer.")
            return False

        # Supplier name
        df_supp = pd.read_excel(FILE_SUPPLIERS, sheet_name="Lieferanten", header=2).copy()
        df_supp["Lieferant_ID_norm"] = df_supp["Lieferant_ID"].astype(builtins.str).str.strip()
        match = df_supp[df_supp["Lieferant_ID_norm"] == supplier_id]
        if match.empty:
            print(f" [!] Lieferant_ID {supplier_id} nicht in Lieferantenliste gefunden.")
            return False

        supplier_name = match["Lieferant_Name"].values[0]

        # ERP dict
        df_erp = pd.read_excel(FILE_ERP, sheet_name=supplier_name, header=None)
        erp_dict = dict(zip(df_erp[0][1:], df_erp[1][1:]))

        final_rows = []

        # Manuelle Kriterien: Wert ist Skalenbeschreibung
        val_col = [c for c in df_man.columns if "bewertung" in builtins.str(c).lower()][0]

        for _, row in df_man.iterrows():
            crit = safe_str(row.get("Kriterium"))
            if not crit or _norm_text(crit) in ("nan", "none"):
                continue  # <- verhindert die "Lücke"

            pts_raw = row.get(val_col)
            try:
                pts = int(pts_raw)
            except:
                pts = 0

            actual = find_matching_criterion(crit, scale_data)
            desc = scale_data.get(actual, {}).get(pts, builtins.str(pts))
            final_rows.append({"Kriterium": crit, "Wert": desc, "Skalapunkte": pts})

        # ERP Kriterien: Wert ist normalisierter/Originalwert (je nach Feld), Punkte via map_erp_to_points
        for crit, val in erp_dict.items():
            crit_s = safe_str(crit)
            if not crit_s or _norm_text(crit_s) in ("nan", "none"):
                continue

            pts = map_erp_to_points(crit_s, val, scale_data)

            # Anzeige-Wert:
            # - Prozent normalisieren fürs Display
            display_val = val
            num = extract_first_number(val)
            if num is not None:
                num2 = normalize_percent_if_needed(num)
                # Prozent anzeigen, wenn Ursprung im 0..1 oder Text hatte %
                if isinstance(val, (int, float)) and 0 <= float(val) <= 1:
                    display_val = f"{num2:.2f}%"
                elif "%" in safe_str(val):
                    display_val = f"{num2:.2f}%"
                else:
                    display_val = f"{num2:.2f}"

            # Sonderwunsch: CO2-Emissionen und Zahlungsbedingungen sollen im Bericht den Originaltext haben
            if _norm_text(crit_s) in KEEP_ORIGINAL_TEXT_FIELDS:
                display_val = safe_str(val)

            final_rows.append({"Kriterium": crit_s, "Wert": display_val, "Skalapunkte": pts})

        out = FOLDER_FINAL / f"Bericht_{supplier_name}_R{round_id}.xlsx"
        pd.DataFrame(final_rows, columns=["Kriterium", "Wert", "Skalapunkte"]).to_excel(out, index=False)
        # Nutzwertanalyse sofort updaten (inkrementell)
        update_nutzwertanalyse_for_supplier(out, supplier_name)

        print(f" [FINISH] Bericht erstellt/aktualisiert: {out.name}")
        return True

    except Exception as e:
        print(f" [!] Fehler bei Bericht: {e}")
        traceback.print_exc()
        return False


# ==========================================
# POLLING LOOP (Server)
# ==========================================
def poll_server_and_process(round_id: str, state: dict, scale_info: dict):
    print(f"\n[PHASE 2] Server-Polling aktiv (Runde {round_id})...")

    while True:
        try:
            if all_done(state):
                print(f"\n[OK] Alle Antworten verarbeitet. {status_line(state)}")
                break

            subs = list_submissions(round_id)  # [{supplier_id, submitted_at}]
            for item in subs:
                sid = safe_str(item.get("supplier_id"))
                submitted_at = float(item.get("submitted_at", 0))

                # Nur Lieferanten, die wir angeschrieben haben
                if sid not in state.get("sent", {}):
                    continue

                prev = state.get("responses", {}).get(sid)
                if prev and submitted_at <= float(prev.get("submitted_at", 0)):
                    continue  # keine neuere Version

                # Download XLSX
                xlsx_bytes = download_submission_xlsx(round_id, sid)
                filename = f"Antwort_{sid}_R{round_id}_{int(submitted_at)}.xlsx"
                file_path = FOLDER_IN / filename
                file_path.write_bytes(xlsx_bytes)

                ok = process_supplier_from_xlsx(file_path, round_id, sid, scale_info)
                if ok:
                    state["responses"][sid] = {"submitted_at": submitted_at, "filename": filename}
                    save_state(state)
                    print(status_line(state))

            time.sleep(6)
        except Exception as e:
            print(f"[WARN] Polling-Fehler: {e}")
            time.sleep(6)


# ==========================================
# NUTZWERTANALYSE (merged-safe)
# ==========================================
def _read_supplier_report(report_path: Path) -> dict:
    df = pd.read_excel(report_path)
    out = {}
    for _, r in df.iterrows():
        crit = safe_str(r.get("Kriterium"))
        if not crit or _norm_text(crit) in ("nan", "none"):
            continue
        try:
            pts = int(r.get("Skalapunkte", 0))
        except:
            pts = 0
        out[crit] = pts
    return out

def excel_set_value_safe(ws, row: int, col: int, value):
    """
    Schreibt in Merge-Zellen immer in die Top-Left-Zelle der MergeArea.
    """
    cell = ws.Cells(row, col)
    try:
        if cell.MergeCells:
            cell = cell.MergeArea.Cells(1, 1)
    except:
        pass
    cell.Value = value


def excel_find_rows(ws, start_row=3, max_scan_rows=400, template_nutzwert_col=4):
    """
    Findet:
      - criteria_rows: Zeilen, wo in der Template-Nutzwert-Spalte (z.B. D) eine Formel steht
      - sum_row: Zeile, wo in Spalte A "Summe Nutzwerte" steht
    """
    criteria_rows = []
    sum_row = None

    for r in range(start_row, start_row + max_scan_rows):
        a_val = ws.Cells(r, 1).Value  # Spalte A
        if a_val and "summe nutzwerte" in builtins.str(a_val).strip().lower():
            sum_row = r

        tmpl = ws.Cells(r, template_nutzwert_col).Formula  # z.B. D
        if tmpl and isinstance(tmpl, builtins.str) and tmpl.startswith("="):
            criteria_rows.append(r)

    return criteria_rows, sum_row


def build_nutzwertanalyse(round_id: str, state: dict):
    """
    Excel-COM Variante:
      - öffnet das Template in echtem Excel
      - erweitert Spalten dynamisch je Lieferant
      - schreibt Bewertungen rein
      - setzt Nutzwert-Formeln (=Gewichtung*Bewertung)
      - setzt Summe Nutzwerte je Lieferant
      - speichert in ROOT/Nutzwertanalyse/
    """
    if state.get("nutzwert_done"):
        return None

    if not FILE_NUTZ_TEMPLATE.exists():
        print(f"[!] Nutzwert-Template fehlt: {FILE_NUTZ_TEMPLATE}")
        return None

    report_files = sorted(FOLDER_FINAL.glob(f"Bericht_*_R{round_id}.xlsx"))
    if not report_files:
        print("[!] Keine Einzelberichte gefunden – Nutzwertanalyse wird nicht erstellt.")
        return None

    suppliers = []
    for p in report_files:
        m = re.match(r"Bericht_(.*)_R(\d+)\.xlsx", p.name, flags=re.IGNORECASE)
        suppliers.append((m.group(1) if m else p.stem, p))

    # === Excel starten
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Performance / Stabilität
    try:
        excel.ScreenUpdating = False
    except:
        pass

    try:
        # xlCalculationAutomatic = -4105
        excel.Calculation = -4105
    except:
        pass

    wb = None
    try:
        wb = excel.Workbooks.Open(builtins.str(FILE_NUTZ_TEMPLATE.resolve()))
        ws = wb.Worksheets(1)

        # Layout-Parameter wie bisher
        HEADER_ROW_1 = 1
        HEADER_ROW_2 = 2
        START_COL = 3  # C
        KRIT_COL = 1   # A
        W_COL = 2      # B (Gewichtung)

        # Kriterienzeilen und Summenzeile finden (anhand Template-Formeln)
        # Template hat typischerweise erste Nutzwertspalte bei D = 4
        criteria_rows, sum_row = excel_find_rows(ws, start_row=3, max_scan_rows=600, template_nutzwert_col=4)

        if not criteria_rows:
            print("[!] Konnte Kriterienzeilen nicht erkennen (keine Formeln im Template gefunden).")
            return None

        # Hilfsfunktion: Punkte aus Report zu Template-Kriterium matchen
        def read_report_map(report_path: Path) -> dict:
            return _read_supplier_report(report_path)

        def find_points(rep_map: dict, template_crit: str) -> int:
            if template_crit in rep_map:
                return int(rep_map[template_crit])
            t = builtins.str(template_crit).strip().lower()
            for k, v in rep_map.items():
                kk = builtins.str(k).strip().lower()
                if kk == t or t in kk or kk in t:
                    return int(v)
            return 0

        # Spalten je Lieferant schreiben
        for idx, (s_name, report_path) in enumerate(suppliers):
            col_bew = START_COL + idx * 2
            col_nutz = START_COL + idx * 2 + 1

            # Header setzen
            excel_set_value_safe(ws, HEADER_ROW_1, col_bew, s_name)
            excel_set_value_safe(ws, HEADER_ROW_2, col_bew, "Bewertung")
            excel_set_value_safe(ws, HEADER_ROW_2, col_nutz, "Nutzwert")

            rep_map = read_report_map(report_path)

            # pro Kriterium die Bewertung eintragen und Nutzwertformel setzen
            for r in criteria_rows:
                crit_txt = ws.Cells(r, KRIT_COL).Value  # Spalte A
                if not crit_txt:
                    continue

                pts = find_points(rep_map, crit_txt)
                ws.Cells(r, col_bew).Value = int(pts)

                # Formel: =Gewichtung*Bewertung  (B*r * Bewertungcol*r)
                w_cell = ws.Cells(r, W_COL).Address  # z.B. "$B$7"
                b_cell = ws.Cells(r, col_bew).Address  # z.B. "$C$7"
                w_cell = builtins.str(w_cell).replace("$", "")
                b_cell = builtins.str(b_cell).replace("$", "")
                ws.Cells(r, col_nutz).Formula = f"={w_cell}*{b_cell}"

            # Summe Nutzwerte
            if sum_row:
                first = ws.Cells(criteria_rows[0], col_nutz).Address
                last = ws.Cells(criteria_rows[-1], col_nutz).Address
                first = builtins.str(first).replace("$", "")
                last = builtins.str(last).replace("$", "")
                ws.Cells(sum_row, col_nutz).Formula = f"=SUM({first}:{last})"

        # Speichern
        out_path = (FOLDER_NUTZ / f"Nutzwertanalyse_R{round_id}.xlsx").resolve()
        wb.SaveAs(builtins.str(out_path))

        state["nutzwert_done"] = True
        save_state(state)
        print(f"\n[FINISH] Nutzwertanalyse erstellt (Excel COM): {out_path}")
        return Path(out_path)

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=True)
        except:
            pass
        try:
            excel.Quit()
        except:
            pass

def excel_open_or_create_nutzwert(excel, out_path: Path, template_path: Path):
    """
    Öffnet bestehende Nutzwertanalyse oder erzeugt sie aus Template.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if out_path.exists():
        wb = excel.Workbooks.Open(builtins.str(out_path.resolve()))
    else:
        wb = excel.Workbooks.Open(builtins.str(template_path.resolve()))
        wb.SaveAs(builtins.str(out_path.resolve()))
    ws = wb.Worksheets(1)
    return wb, ws


def excel_find_supplier_column(ws, supplier_name: str, header_row=1, start_col=3, max_cols=400):
    """
    Sucht, ob Lieferant schon vorhanden ist (Header Row 1, Bewertungsspalte).
    Gibt (col_bew, col_nutz) zurück oder (None, None).
    """
    target = builtins.str(supplier_name).strip().lower()
    c = start_col
    while c <= max_cols:
        v = ws.Cells(header_row, c).Value
        if v and builtins.str(v).strip().lower() == target:
            return c, c + 1
        c += 2
    return None, None


def excel_next_free_supplier_column(ws, header_row=1, start_col=3, max_cols=400):
    """
    Nächste freie 2-Spalten-Gruppe finden.
    """
    c = start_col
    while c <= max_cols:
        v = ws.Cells(header_row, c).Value
        if v is None or builtins.str(v).strip() == "":
            return c, c + 1
        c += 2
    raise RuntimeError("Keine freie Spalte mehr in Nutzwertanalyse (max_cols erreicht).")


def update_nutzwertanalyse_for_supplier(report_path: Path, supplier_name: str):
    """
    Inkrementelles Update:
    - öffnet/erstellt Nutzwertanalyse
    - schreibt/aktualisiert genau diesen Lieferanten
    - speichert
    """
    if not FILE_NUTZ_TEMPLATE.exists():
        print(f"[!] Nutzwert-Template fehlt: {FILE_NUTZ_TEMPLATE}")
        return None

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        try:
            excel.ScreenUpdating = False
        except:
            pass
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except:
            pass

        out_path = (FOLDER_NUTZ / f"Nutzwertanalyse_R{ROUND_ID}.xlsx").resolve()
        wb, ws = excel_open_or_create_nutzwert(excel, Path(out_path), FILE_NUTZ_TEMPLATE)

        try:
            # Layout-Parameter
            HEADER_ROW_1 = 1
            HEADER_ROW_2 = 2
            START_COL = 3  # C
            KRIT_COL = 1   # A
            W_COL = 2      # B (Gewichtung)

            # Kriterienzeilen / Summe finden (Template-Formeln in D)
            criteria_rows, sum_row = excel_find_rows(ws, start_row=3, max_scan_rows=600, template_nutzwert_col=4)
            if not criteria_rows:
                print("[!] Keine Kriterienzeilen im Template erkannt (Formeln fehlen?).")
                wb.Close(SaveChanges=False)
                return None

            # Spalte für Supplier finden oder neu anlegen
            col_bew, col_nutz = excel_find_supplier_column(ws, supplier_name, header_row=HEADER_ROW_1, start_col=START_COL)
            if col_bew is None:
                col_bew, col_nutz = excel_next_free_supplier_column(ws, header_row=HEADER_ROW_1, start_col=START_COL)

                excel_set_value_safe(ws, HEADER_ROW_1, col_bew, supplier_name)
                excel_set_value_safe(ws, HEADER_ROW_2, col_bew, "Bewertung")
                excel_set_value_safe(ws, HEADER_ROW_2, col_nutz, "Nutzwert")

            # Report lesen
            rep_map = _read_supplier_report(report_path)

            def find_points(template_crit: str) -> int:
                if template_crit in rep_map:
                    return int(rep_map[template_crit])
                t = builtins.str(template_crit).strip().lower()
                for k, v in rep_map.items():
                    kk = builtins.str(k).strip().lower()
                    if kk == t or t in kk or kk in t:
                        return int(v)
                return 0

            # Werte schreiben + Formeln setzen
            for r in criteria_rows:
                crit_txt = ws.Cells(r, KRIT_COL).Value
                if not crit_txt:
                    continue

                pts = find_points(crit_txt)
                ws.Cells(r, col_bew).Value = int(pts)

                w_cell = ws.Cells(r, W_COL).Address  # z.B. "$B$7"
                b_cell = ws.Cells(r, col_bew).Address  # z.B. "$C$7"
                w_cell = builtins.str(w_cell).replace("$", "")
                b_cell = builtins.str(b_cell).replace("$", "")
                ws.Cells(r, col_nutz).Formula = f"={w_cell}*{b_cell}"

            # Summe
            if sum_row:
                first = ws.Cells(criteria_rows[0], col_nutz).Address
                last = ws.Cells(criteria_rows[-1], col_nutz).Address
                first = builtins.str(first).replace("$", "")
                last = builtins.str(last).replace("$", "")
                ws.Cells(sum_row, col_nutz).Formula = f"=SUM({first}:{last})"

            # Recalc + Save
            try:
                wb.RefreshAll()
            except:
                pass
            try:
                excel.CalculateFull()
            except:
                pass

            wb.Save()
            print(f"[OK] Nutzwertanalyse inkrementell aktualisiert: {out_path}")
            return Path(out_path)

        finally:
            try:
                wb.Close(SaveChanges=True)
            except:
                pass

    finally:
        try:
            excel.Quit()
        except:
            pass


# ==========================================
# MAIN
# ==========================================
if __name__ == "__main__":
    scale_info = get_comprehensive_scale()
    state = load_state()

    with sync_playwright() as p:
        user_data_dir = ROOT / "Playwright_SCM_Profile"
        browser = p.chromium.launch_persistent_context(builtins.str(user_data_dir), headless=False, slow_mo=600)
        page = browser.new_page()

        # Phase 1: Versand Links
        run_rpa_dispatch(page, ROUND_ID, state)

        # Phase 2: Poll Server, verarbeite Antworten (inkl. Updates bei Duplikaten)
        poll_server_and_process(ROUND_ID, state, scale_info)

        # Nutzwertanalyse
        out_path = build_nutzwertanalyse(ROUND_ID, state)

        # Abschlussmail
        if out_path and not state.get("final_mail_sent", False):
            try:
                subject = f"SCM Nutzwertanalyse Runde {ROUND_ID}"
                body = "Hi,\n\nanbei die Nutzwertanalyse.\n\nViele Grüße\nSCM Bot"
                send_final_mail_outlook(page, SEND_TO_FINAL, subject, body, out_path)
                state["final_mail_sent"] = True
                save_state(state)
                print(f"[FINISH] Abschlussmail mit Nutzwertanalyse gesendet an {SEND_TO_FINAL}")
            except Exception as e:
                print(f"[!] Konnte Abschlussmail nicht senden: {e}")
