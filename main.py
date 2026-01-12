from __future__ import annotations

import os, re, time, json, random, string, builtins, traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Dict, Tuple, List

import requests
import pandas as pd
import win32com.client as win32
from playwright.sync_api import sync_playwright


# =========================
# CONFIG + STATE
# =========================

@dataclass(frozen=True)
class Config:
    base_dir: Path = field(default_factory=lambda: Path(__file__).resolve().parent)
    root: Path = field(init=False)

    folder_in: Path = field(init=False)
    folder_final: Path = field(init=False)
    folder_nutz: Path = field(init=False)
    folder_kons: Path = field(init=False)  # ✅ NEU: eigener Ordner

    file_suppliers: Path = field(init=False)
    file_erp: Path = field(init=False)
    file_scale: Path = field(init=False)

    # Nutzwertanalyse Template + Master
    file_nutz_template: Path = field(init=False)
    file_nutz_master: Path = field(init=False)

    # Konsolidierung Template + Master (NEU)
    file_kons_template: Path = field(init=False)
    file_kons_master: Path = field(init=False)

    form_server: str = field(default_factory=lambda: os.getenv("SCM_FORM_SERVER", "http://localhost:8000"))
    send_to_final: str = "anouar97@gmx.de"

    keep_original_text_fields: frozenset[str] = frozenset({"co2-emissionen", "zahlungsbedingungen"})

    def __post_init__(self):
        object.__setattr__(self, "root", self.base_dir / "ROOT")
        object.__setattr__(self, "folder_in", self.root / "Antworten_Erhalt")
        object.__setattr__(self, "folder_final", self.root / "Einzelberichte_Lieferanten")
        object.__setattr__(self, "folder_nutz", self.root / "Nutzwertanalyse")

        # ✅ Konsolidierung-Ordner
        object.__setattr__(self, "folder_kons", self.root / "Konsolidierung")

        object.__setattr__(self, "file_suppliers", self.root / "1. SCM-Anwendungen(MA)_Lieferantenuebersicht.xlsx")
        object.__setattr__(self, "file_erp", self.root / "4. SCM-Anwendungen(MA)_ERP-System.xlsx")
        object.__setattr__(self, "file_scale", self.root / "3. SCM-Anwendungen(MA)_Gesamtbewertung.xlsx")

        object.__setattr__(self, "file_nutz_template", self.root / "5. SCM-Nutzwertanalyse.xlsx")
        object.__setattr__(self, "file_nutz_master", self.folder_nutz / "Nutzwertanalyse_Zentral.xlsx")

        # ✅ neue Konsolidierung (separate Datei)
        object.__setattr__(self, "file_kons_template", self.root / "6. SCM-Konsolidierung.xlsx")
        # ✅ Master liegt im Ordner "Konsolidierung"
        object.__setattr__(self, "file_kons_master", self.folder_kons / "Konsolidierung_Zentral.xlsx")

    def ensure_dirs(self):
        for d in (self.folder_in, self.folder_final, self.folder_nutz, self.folder_kons):
            d.mkdir(parents=True, exist_ok=True)


@dataclass
class State:
    path: Path
    round_id: str
    sent: Dict[str, dict] = field(default_factory=dict)      # sid -> meta
    responses: Dict[str, dict] = field(default_factory=dict) # sid -> meta (inkl. round_id)
    final_mail_sent: bool = False

    @staticmethod
    def load_or_new(cfg: Config) -> "State":
        p = cfg.root / "round_state.json"
        if p.exists():
            try:
                raw = json.loads(p.read_text(encoding="utf-8"))
                rid = raw.get("round_id")
                if not rid:
                    rid = "".join(random.choices(string.digits, k=8))
                return State(
                    path=p,
                    round_id=rid,
                    sent=raw.get("sent", {}),
                    responses=raw.get("responses", {}),
                    final_mail_sent=bool(raw.get("final_mail_sent", False)),
                )
            except Exception:
                pass
        return State(path=p, round_id="".join(random.choices(string.digits, k=8)))

    def save(self):
        self.path.write_text(
            json.dumps(
                {
                    "round_id": self.round_id,
                    "sent": self.sent,
                    "responses": self.responses,
                    "final_mail_sent": self.final_mail_sent,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    def has_response_for_round(self, sid: str, rid: str) -> bool:
        meta = self.responses.get(sid) or {}
        return meta.get("round_id") == rid

    def has_response_for_current_round(self, sid: str) -> bool:
        return self.has_response_for_round(sid, self.round_id)

    def all_done_for_round(self, rid: str) -> bool:
        return bool(self.sent) and all(self.has_response_for_round(sid, rid) for sid in self.sent.keys())

    def all_done(self) -> bool:
        return self.all_done_for_round(self.round_id)

    def status(self) -> str:
        got = sum(1 for sid in self.sent.keys() if self.has_response_for_current_round(sid))
        return f"[STATUS] {got} von {len(self.sent)} Antworten (mind. 1x, Runde {self.round_id})"

    def start_new_round(self) -> None:
        self.round_id = "".join(random.choices(string.digits, k=8))
        self.sent = {}
        self.responses = {}
        self.final_mail_sent = False
        self.save()

    def ensure_not_stuck_on_finished_round(self) -> None:
        if self.sent and self.all_done():
            print(f"[INFO] Vorherige Runde {self.round_id} ist abgeschlossen -> starte neue Runde.")
            self.start_new_round()


# =========================
# STARTUP INPUT: MONAT + JAHR
# =========================

def _ask_int(prompt: str) -> int:
    while True:
        s = input(prompt).strip()
        try:
            return int(s)
        except Exception:
            print("[EINGABE] Bitte eine Zahl eingeben.")

def prompt_month_year() -> Tuple[int, int]:
    print("\n==============================")
    print("Nutzwertanalyse: Zeitraum wählen")
    print("==============================")
    while True:
        m = _ask_int("Monat (1-12): ")
        if 1 <= m <= 12:
            break
        print("[EINGABE] Monat muss zwischen 1 und 12 sein.")
    while True:
        y = _ask_int("Jahr (z.B. 2026): ")
        if 2000 <= y <= 2100:
            break
        print("[EINGABE] Jahr muss zwischen 2000 und 2100 liegen.")
    print(f"[OK] Zeitraum gesetzt: {m:02d} / {y}\n")
    return m, y

INVALID_SHEET_CHARS = r'[:\\/*?\[\]]'

def sanitize_sheet_name(name: str, *, fallback: str = "Sheet") -> str:
    if name is None:
        return fallback
    s = builtins.str(name).strip()
    s = re.sub(INVALID_SHEET_CHARS, "-", s)
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        s = fallback
    if len(s) > 31:
        s = s[:31].rstrip()
    return s

def sheet_name_from_my(m: int, y: int) -> str:
    # Excel-Blattname: kein "/" (z.B. 01-2026)
    return sanitize_sheet_name(f"{m:02d}-{y}")

def month_label_from_my(m: int, y: int) -> str:
    return f"{m:02d} / {y}"


# =========================
# TEXT/NUM UTILS + SCALE
# =========================

def sstr(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return builtins.str(x).strip()

def norm(s: Any) -> str:
    t = sstr(s).lower()
    t = t.replace("\u00A0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", t).strip()

def first_num(v: Any) -> Optional[float]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = sstr(v).lower()
    t = t.replace("kg co2e", "").replace("kgco2e", "").replace("co2e", "").replace("kg", "").replace("%", "")
    t = t.replace(" ", "").replace(",", ".")
    m = re.search(r"[-+]?\d*\.\d+|\d+", t)
    return float(m.group()) if m else None

def as_percent(x: float) -> float:
    return round(x * 100, 2) if 0 <= x <= 1 else round(float(x), 2)

def match_scale_key(crit: str, scale: dict) -> Optional[str]:
    c = norm(crit)
    if crit in scale:
        return crit
    for k in scale.keys():
        kk = norm(k)
        if kk == c or c in kk or kk in c:
            return k
    return None

def parse_cond(cond: str, val: float) -> bool:
    t = norm(cond).replace("%", "").replace("kg co2e", "").replace("kgco2e", "").replace("co2e", "").replace("kg", "")
    raw = t.replace(" ", "").replace(",", ".")
    for op, fn in [
        ("<=", lambda a, b: a <= b),
        (">=", lambda a, b: a >= b),
        ("<", lambda a, b: a < b),
        (">", lambda a, b: a > b),
    ]:
        if op in raw:
            n = first_num(raw.split(op, 1)[1])
            return n is not None and fn(val, n)
    if "–" in t or "-" in t:
        lohi = re.split(r"[–-]", t)
        if len(lohi) >= 2:
            lo, hi = first_num(lohi[0]), first_num(lohi[1])
            return lo is not None and hi is not None and lo <= val <= hi
    return False

def is_no(v: Any) -> bool:
    t = norm(v)
    return t in ("", "nan", "none") or any(x in t for x in ("nicht vorhanden", "nichtvorhanden", "nein", "no", "false", "0", "keine", "kein"))

def erp_points(crit: str, v: Any, scale: dict) -> int:
    k = match_scale_key(crit, scale)
    if not k:
        return 0
    kl = norm(k)
    if ("iso " in kl or kl.startswith("iso ")) and is_no(v):
        return 0

    num = first_num(v)
    if num is not None:
        num = as_percent(num)
        for pts in (100, 80, 60, 40, 20, 0):
            if parse_cond(sstr(scale[k].get(pts, "")), num):
                return pts

    vs = norm(v)
    for pts in (100, 80, 60, 40, 20, 0):
        st = norm(scale[k].get(pts, ""))
        if st and (vs in st or st in vs):
            return pts
        if "code of conduct" in kl or kl in ("coc",):
            if pts == 100 and (("bme" in vs and "bme" in st) or ("kb" in vs and "kb" in st)):
                return 100
    return 0

def load_scale(file_scale: Path) -> dict:
    df = pd.read_excel(file_scale, sheet_name="Skala", header=None)
    out = {}
    for i in range(4, len(df)):
        crit = sstr(df.iloc[i, 0])
        if not crit or norm(crit) in ("nan", "none"):
            continue
        out[crit] = {
            0: sstr(df.iloc[i, 4]),
            20: sstr(df.iloc[i, 5]),
            40: sstr(df.iloc[i, 6]),
            60: sstr(df.iloc[i, 7]),
            80: sstr(df.iloc[i, 8]),
            100: sstr(df.iloc[i, 9]),
        }
    return out

def disp_val(val: Any, crit_norm: str, keep: frozenset[str]) -> str:
    if crit_norm in keep:
        return sstr(val)
    n = first_num(val)
    if n is None:
        return sstr(val)
    n2 = as_percent(n)
    raw = sstr(val)
    if "%" in raw or (isinstance(val, (int, float)) and 0 <= float(val) <= 1):
        return f"{n2:.2f}%"
    return f"{n2:.2f}"


# =========================
# DATA IO
# =========================

def suppliers_df(cfg: Config) -> pd.DataFrame:
    return pd.read_excel(cfg.file_suppliers, sheet_name="Lieferanten", header=2).dropna(subset=["Lieferant_Name"])

def supplier_name(cfg: Config, sid: str) -> Optional[str]:
    df = suppliers_df(cfg).copy()
    df["id"] = df["Lieferant_ID"].astype(builtins.str).str.strip()
    m = df[df["id"] == sid]
    return None if m.empty else m["Lieferant_Name"].values[0]

def erp_dict(cfg: Config, sup_name: str) -> dict:
    df = pd.read_excel(cfg.file_erp, sheet_name=sup_name, header=None)
    return dict(zip(df[0][1:], df[1][1:]))

def build_report(cfg: Config, scale: dict, rid: str, sid: str, xlsx: Path) -> Optional[Path]:
    try:
        df_man = pd.read_excel(xlsx)
        if df_man is None or df_man.empty:
            return None

        sup_name = supplier_name(cfg, sid)
        if not sup_name:
            return None

        erp = erp_dict(cfg, sup_name)
        val_col = next((c for c in df_man.columns if "bewertung" in norm(c)), None)
        if not val_col:
            return None

        rows = []
        for _, r in df_man.iterrows():
            crit = sstr(r.get("Kriterium"))
            if not crit or norm(crit) in ("nan", "none"):
                continue
            try:
                pts = int(r.get(val_col))
            except Exception:
                pts = 0
            k = match_scale_key(crit, scale)
            desc = scale.get(k, {}).get(pts, sstr(pts))
            rows.append({"Kriterium": crit, "Wert": desc, "Skalapunkte": pts})

        for crit, val in erp.items():
            cs = sstr(crit)
            if not cs or norm(cs) in ("nan", "none"):
                continue
            pts = erp_points(cs, val, scale)
            rows.append({"Kriterium": cs, "Wert": disp_val(val, norm(cs), cfg.keep_original_text_fields), "Skalapunkte": pts})

        out = cfg.folder_final / f"Bericht_{sup_name}_R{rid}.xlsx"
        pd.DataFrame(rows, columns=["Kriterium", "Wert", "Skalapunkte"]).to_excel(out, index=False)
        print(f"[FINISH] Bericht: {out.name}")
        return out

    except Exception as e:
        print(f"[WARN] Bericht-Fehler: {e}")
        traceback.print_exc()
        return None


# =========================
# RUNDE: Antworten aus Ordner NUR für AKTUELLE Runde finden
# =========================

ANSWER_RE = re.compile(r"Antwort_(?P<sid>K_\d+)_R(?P<rid>\d{8})_(?P<ts>\d+)\.xlsx$", re.IGNORECASE)

def list_round_answer_files(cfg: Config, rid: str) -> Dict[str, Path]:
    latest: Dict[str, Tuple[float, Path]] = {}
    if not cfg.folder_in.exists():
        return {}

    for p in cfg.folder_in.glob("Antwort_*.xlsx"):
        m = ANSWER_RE.match(p.name)
        if not m:
            continue
        if m.group("rid") != rid:
            continue
        sid = m.group("sid").upper()
        try:
            mt = p.stat().st_mtime
        except Exception:
            mt = time.time()
        prev = latest.get(sid)
        if prev is None or mt > prev[0]:
            latest[sid] = (mt, p)

    return {sid: pp for sid, (mt, pp) in latest.items()}


# =========================
# EXCEL / Nutzwertanalyse (zentral + sheet pro Monat/Jahr)
# =========================

def _merged_value(cell):
    try:
        if cell.MergeCells:
            return cell.MergeArea.Cells(1, 1).Value
    except Exception:
        pass
    return cell.Value

def safe_unmerge_and_clear(rng):
    try:
        rng.UnMerge()
    except Exception:
        pass
    try:
        rng.Clear()
    except Exception:
        try:
            rng.ClearContents()
        except Exception:
            pass

def excel_set_value_safe(ws, row: int, col: int, value: Any) -> None:
    cell = ws.Cells(row, col)
    try:
        if cell.MergeCells:
            cell = cell.MergeArea.Cells(1, 1)
    except Exception:
        pass
    cell.Value = value

def excel_last_used_row(ws, min_row: int = 2, max_row: int = 1200) -> int:
    last = min_row
    for r in range(min_row, max_row + 1):
        a = ws.Cells(r, 1).Value
        d_formula = ws.Cells(r, 4).Formula
        if (a is not None and str(a).strip() != "") or (d_formula is not None and str(d_formula).strip().startswith("=")):
            last = r
    return last

def excel_find_rows(ws, start_row=3, max_scan_rows=1200, template_nutzwert_col=4) -> Tuple[List[int], Optional[int]]:
    criteria_rows: List[int] = []
    sum_row = None
    for r in range(start_row, start_row + max_scan_rows):
        a_val = _merged_value(ws.Cells(r, 1))
        b_val = _merged_value(ws.Cells(r, 2))
        text = (builtins.str(a_val) if a_val is not None else "") + " " + (builtins.str(b_val) if b_val is not None else "")
        if "summe nutzwerte" in text.strip().lower():
            sum_row = r
        tmpl = ws.Cells(r, template_nutzwert_col).Formula
        if tmpl and isinstance(tmpl, builtins.str) and tmpl.startswith("="):
            criteria_rows.append(r)
    return criteria_rows, sum_row

def excel_find_supplier_column(ws, supplier_name_: str, header_row=1, start_col=3, max_cols=400) -> Tuple[Optional[int], Optional[int]]:
    target = builtins.str(supplier_name_).strip().lower()
    c = start_col
    while c <= max_cols:
        v = ws.Cells(header_row, c).Value
        if v and builtins.str(v).strip().lower() == target:
            return c, c + 1
        c += 2
    return None, None

def excel_next_free_supplier_column(ws, header_row=1, start_col=3, max_cols=400) -> Tuple[int, int]:
    c = start_col
    while c <= max_cols:
        v = ws.Cells(header_row, c).Value
        if v is None or builtins.str(v).strip() == "":
            return c, c + 1
        c += 2
    raise RuntimeError("Keine freie Spalte mehr in Nutzwertanalyse (max_cols erreicht).")

def excel_apply_template_pair_from_schablone(ws_nutz, ws_schablone, dest_bew_col: int, dest_nutz_col: int,
                                            template_bew_col: int = 3, template_nutz_col: int = 4) -> None:
    last_row = excel_last_used_row(ws_schablone, min_row=1, max_row=1200)
    src = ws_schablone.Range(ws_schablone.Cells(1, template_bew_col), ws_schablone.Cells(last_row, template_nutz_col))
    dst = ws_nutz.Range(ws_nutz.Cells(1, dest_bew_col), ws_nutz.Cells(last_row, dest_nutz_col))
    safe_unmerge_and_clear(dst)
    src.Copy(dst)
    try:
        ws_nutz.Application.CutCopyMode = False
    except Exception:
        pass

def col_letter(col_idx: int) -> str:
    letters = ""
    while col_idx:
        col_idx, rem = divmod(col_idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

def rewrite_sum_formula(template_formula: str, old_letter: str, new_letter: str) -> str:
    pattern = rf"(\$?){re.escape(old_letter)}(\$?\d+)"
    return re.sub(pattern, rf"\1{new_letter}\2", template_formula)

def set_sum_formula_like_template(ws, ws_schablone, sum_row: int, dest_bew_col: int, dest_nutz_col: int):
    tmpl = ws_schablone.Cells(sum_row, 3).Formula
    if not tmpl or not builtins.str(tmpl).startswith("="):
        return
    old_nutz_letter = col_letter(4)
    new_nutz_letter = col_letter(dest_nutz_col)
    new_formula = rewrite_sum_formula(builtins.str(tmpl), old_nutz_letter, new_nutz_letter)
    cell = ws.Cells(sum_row, dest_bew_col)
    try:
        if cell.MergeCells:
            cell = cell.MergeArea.Cells(1, 1)
    except Exception:
        pass
    cell.Formula = new_formula

def norm_key(x: Any) -> str:
    t = sstr(x).lower()
    t = t.replace("\u00A0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"[^a-z0-9äöüß ]+", "", t)
    return t

def read_report_points(report_xlsx: Path) -> dict:
    df = pd.read_excel(report_xlsx)
    m = {}
    for _, r in df.iterrows():
        k = norm_key(r.get("Kriterium"))
        if not k:
            continue
        try:
            m[k] = int(r.get("Skalapunkte", 0))
        except Exception:
            m[k] = 0
    return m

def match_points(rep_map: dict, template_crit: Any) -> int:
    if template_crit is None:
        return 0
    k = norm_key(template_crit)
    if not k:
        return 0
    if k in rep_map:
        return int(rep_map[k])
    for kk, v in rep_map.items():
        if kk in k or k in kk:
            return int(v)
    return 0

def open_or_create_master(excel, master_path: Path, template_path: Path):
    master_path.parent.mkdir(parents=True, exist_ok=True)
    if master_path.exists():
        wb = excel.Workbooks.Open(builtins.str(master_path.resolve()))
    else:
        wb = excel.Workbooks.Open(builtins.str(template_path.resolve()))
        wb.SaveAs(builtins.str(master_path.resolve()))
    return wb

def ensure_sheet_from_template(wb, sheet_name: str):
    sheet_name = sanitize_sheet_name(sheet_name)
    for ws in wb.Worksheets:
        try:
            if builtins.str(ws.Name).strip().lower() == sheet_name.strip().lower():
                return ws
        except Exception:
            continue

    ws_s = wb.Worksheets("Schablone")
    ws_new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws_new.Name = sheet_name

    last_row = 1
    for r in range(1, 1200):
        v = _merged_value(ws_s.Cells(r, 1))
        if v is None or builtins.str(v).strip() == "":
            continue
        last_row = r

    dst_all = ws_new.Range(ws_new.Cells(1, 1), ws_new.Cells(last_row, 4))
    safe_unmerge_and_clear(dst_all)

    src = ws_s.Range(ws_s.Cells(1, 1), ws_s.Cells(last_row, 4))
    dst = ws_new.Range(ws_new.Cells(1, 1), ws_new.Cells(last_row, 4))
    src.Copy(dst)

    ws_new.Cells(1, 3).Value = ""
    ws_new.Cells(2, 3).Value = "Bewertung"
    ws_new.Cells(2, 4).Value = "Nutzwert"

    try:
        ws_new.Application.CutCopyMode = False
    except Exception:
        pass

    return ws_new


COM_BUSY_HRESULTS = {-2147418111, -2147417846, -2147220995}

def com_call_with_retry(fn, *, tries: int = 9, base_sleep: float = 0.35, label: str = "COM"):
    last_exc = None
    for i in range(tries):
        try:
            return fn()
        except Exception as e:
            last_exc = e
            hresult = getattr(e, "hresult", None)
            if hresult in COM_BUSY_HRESULTS or any(str(code) in str(e) for code in COM_BUSY_HRESULTS):
                time.sleep(base_sleep * (1.6 ** i) + random.random() * 0.2)
                continue
            raise
    raise last_exc if last_exc else RuntimeError(f"{label}: Unbekannter Fehler")


def upsert_into_master_nutz(cfg: Config, report_xlsx: Path, supplier_name_: str, sheet_name: str) -> Tuple[Path, Optional[float]]:
    """
    Schreibt Supplier-Spalten ins Monatsblatt.
    Gibt zusätzlich die 'Summe Nutzwerte' (Bewertungs-Spalte) zurück.
    """
    master_path = cfg.file_nutz_master.resolve()

    def _do_update():
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            try:
                excel.ScreenUpdating = False
            except Exception:
                pass
            try:
                excel.Calculation = -4105
            except Exception:
                pass

            wb = open_or_create_master(excel, master_path, cfg.file_nutz_template)
            try:
                ws = ensure_sheet_from_template(wb, sheet_name)
                ws_s = wb.Worksheets("Schablone")

                criteria_rows, sum_row = excel_find_rows(ws, start_row=3, max_scan_rows=1200, template_nutzwert_col=4)
                if not criteria_rows:
                    raise RuntimeError("Keine Kriterienzeilen erkannt (Template-Formeln fehlen?).")

                col_bew, col_nutz = excel_find_supplier_column(ws, supplier_name_, header_row=1, start_col=3)
                if col_bew is None:
                    col_bew, col_nutz = excel_next_free_supplier_column(ws, header_row=1, start_col=3)
                    excel_apply_template_pair_from_schablone(ws, ws_s, col_bew, col_nutz)

                    excel_set_value_safe(ws, 1, col_bew, supplier_name_)
                    excel_set_value_safe(ws, 2, col_bew, "Bewertung")
                    excel_set_value_safe(ws, 2, col_nutz, "Nutzwert")

                rep_map = read_report_points(report_xlsx)

                for r in criteria_rows:
                    crit_txt = _merged_value(ws.Cells(r, 1))
                    pts = match_points(rep_map, crit_txt)
                    ws.Cells(r, col_bew).Value = int(pts)

                    w_cell = builtins.str(ws.Cells(r, 2).Address).replace("$", "")
                    b_cell = builtins.str(ws.Cells(r, col_bew).Address).replace("$", "")
                    ws.Cells(r, col_nutz).Formula = f"={w_cell}*{b_cell}"

                if sum_row:
                    set_sum_formula_like_template(ws, ws_s, sum_row, col_bew, col_nutz)

                try:
                    wb.RefreshAll()
                except Exception:
                    pass
                try:
                    excel.CalculateFull()
                except Exception:
                    pass

                sum_val = None
                if sum_row:
                    try:
                        v = ws.Cells(int(sum_row), int(col_bew)).Value
                        sum_val = float(v) if v is not None and builtins.str(v).strip() != "" else None
                    except Exception:
                        sum_val = None

                wb.Save()
                return master_path, sum_val

            finally:
                try:
                    wb.Close(SaveChanges=True)
                except Exception:
                    pass
        finally:
            try:
                excel.Quit()
            except Exception:
                pass

    result_path, sum_val = com_call_with_retry(_do_update, tries=9, base_sleep=0.35, label="Excel/Nutzwertanalyse_Master")
    print(f"[OK] Master-Nutzwertanalyse aktualisiert: {Path(result_path).name} (Sheet: {sheet_name})")
    return Path(result_path), sum_val


# =========================
# KONSOLIDIERUNG (FIX): TEMPLATE NIE ÜBERSCHREIBEN, SCHABLONE BLEIBT IM MASTER
# =========================

def open_or_create_kons_master(excel, kons_master_path: Path, kons_template_path: Path):
    """
    Öffnet immer das MASTER.
    Wenn Master fehlt: aus Template erzeugen (SaveAs).
    Danach wird NUR das Master bearbeitet.
    """
    kons_master_path.parent.mkdir(parents=True, exist_ok=True)
    if kons_master_path.exists():
        return excel.Workbooks.Open(builtins.str(kons_master_path.resolve()))

    if not kons_template_path.exists():
        raise RuntimeError(f"Konsolidierungs-Template fehlt: {kons_template_path}")

    wb = excel.Workbooks.Open(builtins.str(kons_template_path.resolve()))
    wb.SaveAs(builtins.str(kons_master_path.resolve()))
    return wb

def ensure_template_sheet_exists(wb, template_sheet_name: str = "Schablone"):
    for ws in wb.Worksheets:
        if builtins.str(ws.Name).strip().lower() == template_sheet_name.lower():
            return ws
    raise RuntimeError(f"In '{Path(wb.FullName).name}' muss ein Blatt '{template_sheet_name}' existieren.")

def ensure_unique_sheet_name(wb, base_name: str) -> str:
    name = sanitize_sheet_name(base_name, fallback="Lieferant")
    existing = {builtins.str(ws.Name).strip().lower() for ws in wb.Worksheets}

    if name.strip().lower() == "schablone":
        name = "Schablone_1"

    if name.lower() not in existing:
        return name

    for i in range(2, 200):
        cand = sanitize_sheet_name(f"{name[:28]}_{i}", fallback=f"Lieferant_{i}")
        if cand.lower() not in existing and cand.lower() != "schablone":
            return cand

    raise RuntimeError("Konnte keinen eindeutigen Blattnamen erzeugen (zu viele Kollisionen).")

def ensure_supplier_sheet_from_kons_template(wb, supplier_name_: str, template_sheet_name: str = "Schablone") -> Any:
    """
    ✅ NICHT ws_t.Copy() (das kann Excel/COM manchmal so "verschieben", dass es wie Überschreiben wirkt).
    ✅ Stattdessen: neues Blatt anlegen und UsedRange aus 'Schablone' kopieren.
    'Schablone' bleibt IMMER erhalten.
    """
    supplier_real = builtins.str(supplier_name_).strip()
    target_name = sanitize_sheet_name(supplier_real, fallback="Lieferant")

    # existiert bereits?
    for ws in wb.Worksheets:
        try:
            if builtins.str(ws.Name).strip().lower() == target_name.lower():
                ws.Cells(1, 1).Value = supplier_real
                return ws
        except Exception:
            continue

    ws_tpl = ensure_template_sheet_exists(wb, template_sheet_name=template_sheet_name)

    # Zielname (falls Kollision)
    target_name = ensure_unique_sheet_name(wb, target_name)

    # Neues Blatt
    ws_new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws_new.Name = target_name

    # Template-Inhalt kopieren (UsedRange)
    try:
        src = ws_tpl.UsedRange
        dst = ws_new.Range(ws_new.Cells(1, 1), ws_new.Cells(src.Rows.Count, src.Columns.Count))
        dst.Clear()
        src.Copy(dst)
        try:
            wb.Application.CutCopyMode = False
        except Exception:
            pass
    except Exception:
        # Fallback (falls UsedRange zickt)
        src = ws_tpl.Range("A1:Z50")
        dst = ws_new.Range("A1:Z50")
        dst.Clear()
        src.Copy(dst)
        try:
            wb.Application.CutCopyMode = False
        except Exception:
            pass

    ws_new.Cells(1, 1).Value = supplier_real
    return ws_new

def find_or_create_period_column(ws, period: str, row_period: int = 2, start_col: int = 2, max_cols: int = 400) -> int:
    period = builtins.str(period).strip()

    for c in range(start_col, max_cols + 1):
        v = ws.Cells(row_period, c).Value
        if v and builtins.str(v).strip() == period:
            return c

    for c in range(start_col, max_cols + 1):
        v = ws.Cells(row_period, c).Value
        if v is None or builtins.str(v).strip() == "":
            ws.Cells(row_period, c).Value = period
            return c

    raise RuntimeError("Keine freie Spalte mehr in Konsolidierung (max_cols erreicht).")

def last_used_period_col(ws, row_period: int = 2, start_col: int = 2, max_cols: int = 400) -> int:
    last = start_col - 1
    for c in range(start_col, max_cols + 1):
        v = ws.Cells(row_period, c).Value
        if v is not None and builtins.str(v).strip() != "":
            last = c
    return last

def update_supplier_konsolidierung(cfg: Config, supplier_name_: str, period: str, value: Optional[float]) -> Path:
    """
    Schreibt/ergänzt pro Lieferant:
    - A1 = Lieferantenname (echter Name)
    - Row2: Zeiträume (B2..)
    - Row3: Werte (B3..)
    - B5: arithm. Mittel über Zeile 3 (alle Werte)
    """
    kons_master_path = cfg.file_kons_master.resolve()
    kons_template_path = cfg.file_kons_template.resolve()

    def _do_update():
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            try:
                excel.ScreenUpdating = False
            except Exception:
                pass
            try:
                excel.Calculation = -4105
            except Exception:
                pass

            wb = open_or_create_kons_master(excel, kons_master_path, kons_template_path)
            try:
                # ✅ WICHTIG: wir arbeiten NUR im Master!
                ws = ensure_supplier_sheet_from_kons_template(wb, supplier_name_, template_sheet_name="Schablone")

                col = find_or_create_period_column(ws, period, row_period=2, start_col=2, max_cols=400)

                if value is None:
                    ws.Cells(3, col).Value = ""
                else:
                    ws.Cells(3, col).Value = float(value)

                last_col = last_used_period_col(ws, row_period=2, start_col=2, max_cols=400)
                if last_col >= 2:
                    start_cell = ws.Cells(3, 2).Address.replace("$", "")
                    end_cell = ws.Cells(3, last_col).Address.replace("$", "")
                    rng = f"{start_cell}:{end_cell}"
                    ws.Cells(5, 2).Formula = f'=IF(COUNT({rng})=0,"",AVERAGE({rng}))'
                else:
                    ws.Cells(5, 2).Formula = ""

                try:
                    wb.RefreshAll()
                except Exception:
                    pass
                try:
                    excel.CalculateFull()
                except Exception:
                    pass

                wb.Save()
                return kons_master_path

            finally:
                try:
                    wb.Close(SaveChanges=True)
                except Exception:
                    pass
        finally:
            try:
                excel.Quit()
            except Exception:
                pass

    result = com_call_with_retry(_do_update, tries=9, base_sleep=0.35, label="Excel/Konsolidierung_Supplier")
    print(f"[OK] Konsolidierung aktualisiert: {Path(result).name} | {supplier_name_} | {period}")
    return Path(result)


# =========================
# OUTLOOK WEB (Playwright)
# =========================

class OutlookUI:
    def __init__(self, page):
        self.page = page

    def open_mail(self) -> None:
        self.page.goto("https://outlook.office.com/mail/")
        self.page.wait_for_selector('button[aria-label*="Neue"]', timeout=120000)

    def new_mail(self, to_email: str, subject: str, body: str) -> None:
        self.page.click('button[aria-label*="Neue"]')
        self.page.wait_for_timeout(500)
        self.page.fill('div[aria-label="An"]', to_email)
        self.page.fill('input[placeholder*="Betreff"]', subject)
        self.page.locator('div[role="textbox"]').first.click()
        self.page.keyboard.type(body)
        self.page.click('button[aria-label*="Senden"]')
        self.page.wait_for_selector('div[aria-label="An"]', state="hidden", timeout=30000)

    def new_mail_with_attachments(self, to_email: str, subject: str, body: str, attachment_paths: List[Path]) -> None:
        self.page.click('button[aria-label*="Neue"]')
        self.page.wait_for_timeout(500)
        self.page.fill('div[aria-label="An"]', to_email)
        self.page.fill('input[placeholder*="Betreff"]', subject)
        self.page.locator('div[role="textbox"]').first.click()
        self.page.keyboard.type(body)

        for ap in attachment_paths:
            if not ap or not Path(ap).exists():
                continue

            self.page.locator('button[aria-label*="Datei anfügen"]').first.click()
            with self.page.expect_file_chooser() as fc:
                self.page.locator('button[aria-label*="Diesen Computer durchsuchen"]').first.click()
            fc.value.set_files(str(Path(ap).resolve()))
            self.page.wait_for_timeout(1200)

        self.page.click('button[aria-label*="Senden"]')
        self.page.wait_for_selector('div[aria-label="An"]', state="hidden", timeout=30000)

    def open_folder(self, folder_name: str) -> None:
        self.open_mail()
        self.page.wait_for_timeout(800)

        locators = [
            self.page.get_by_role("treeitem", name=folder_name),
            self.page.get_by_role("button", name=folder_name),
            self.page.get_by_role("link", name=folder_name),
            self.page.get_by_text(folder_name, exact=True),
        ]
        for loc in locators:
            try:
                if loc.first.is_visible(timeout=2500):
                    loc.first.click()
                    self.page.wait_for_timeout(1200)
                    return
            except Exception:
                pass
        raise RuntimeError(f"Outlook-Ordner '{folder_name}' konnte nicht geöffnet werden.")

    def refresh_folder(self) -> None:
        candidates = [
            self.page.locator('button[aria-label*="Aktualisieren"]').first,
            self.page.locator('button[title*="Aktualisieren"]').first,
            self.page.get_by_role("button", name=re.compile(r"Aktualisieren|Refresh", re.IGNORECASE)).first,
        ]
        for c in candidates:
            try:
                if c and c.is_visible(timeout=800):
                    c.click()
                    self.page.wait_for_timeout(800)
                    return
            except Exception:
                pass

    def message_rows(self):
        candidates = [
            self.page.locator('div[role="listbox"] div[role="option"]'),
            self.page.locator('div[role="grid"] div[role="row"]'),
        ]
        for c in candidates:
            try:
                if c.count() > 0:
                    return c
            except Exception:
                pass
        return candidates[0]

    def archive_current_mail(self) -> None:
        try:
            self.page.keyboard.press("E")
            self.page.locator('div[role="listbox"]').first.wait_for(state="visible", timeout=6000)
        except Exception:
            pass

    def open_download_xlsx(self, row, download_dir: Path) -> Optional[Path]:
        download_dir.mkdir(parents=True, exist_ok=True)
        try:
            row.click(timeout=4000)
            self.page.locator('div[role="main"]').wait_for(state="visible", timeout=8000)

            attachments = self.page.locator('div[role="listbox"][aria-label="Dateianlagen"]')
            attachments.wait_for(state="visible", timeout=4000)

            item = attachments.locator('div[role="option"]').filter(
                has=self.page.locator('[title$=".xlsx"], [title*=".xlsx"]')
            ).first
            if not item.is_visible(timeout=2000):
                return None

            more_btn = item.locator('button[aria-label="Weitere Aktionen"], button[title="Weitere Aktionen"]').first
            more_btn.wait_for(state="visible", timeout=2000)
            more_btn.click(timeout=2000)

            menu_item = self.page.get_by_role("menuitem", name=re.compile(r"Herunterladen|Download", re.IGNORECASE)).first
            menu_item.wait_for(state="visible", timeout=2000)

            with self.page.expect_download(timeout=15000) as dl_info:
                menu_item.click(timeout=2000)

            dl = dl_info.value
            target = download_dir / dl.suggested_filename
            dl.save_as(str(target.resolve()))
            return target

        except Exception:
            try:
                self.page.keyboard.press("Escape")
            except Exception:
                pass
            return None


# =========================
# SERVER API
# =========================

def form_link(base: str, sid: str, rid: str) -> str:
    r = requests.get(f"{base}/issue-link", params={"supplier_id": sid, "round_id": rid}, timeout=10)
    r.raise_for_status()
    return r.json()["url"]


# =========================
# PHASES
# =========================

def phase_send_links(cfg: Config, outlook: OutlookUI, st: State):
    print(f"\n[PHASE 1] Versand Runde {st.round_id} ...")
    df = suppliers_df(cfg)
    outlook.open_mail()

    for _, r in df.iterrows():
        sid, email = sstr(r.get("Lieferant_ID")), sstr(r.get("Email"))
        if not sid or not email:
            continue

        if sid in st.sent and st.sent[sid].get("round_id") == st.round_id:
            continue

        name, lname = sstr(r.get("Name")), sstr(r.get("Lieferant_Name"))
        url = form_link(cfg.form_server, sid, st.round_id)
        subject = f"SCM-Bewertung | Runde {st.round_id} | {sid}"
        body = (
            f"Hallo {name},\n\n"
            f"im Rahmen unseres Lieferantenbewertungsprozesses bitten wir Sie, die Bewertung über den folgenden Link auszufüllen:\n\n"
            f"{url}\n\n"
            f"Vielen Dank.\n\nFreundliche Grüße\nIhr SCM-Team\n(Runde {st.round_id})"
        )
        outlook.new_mail(email, subject, body)
        st.sent[sid] = {"name": lname, "email": email, "sent_at": time.time(), "round_id": st.round_id}
        st.save()
        print(f"[OK] Link an {lname} ({sid})  {st.status()}")


def ingest_existing_round_answers(cfg: Config, scale: dict, st: State, sheet_name: str) -> None:
    """
    ✅ Scannt NUR auf st.round_id.
    - erstellt Bericht
    - schreibt ins Nutzwert-Master Monatsblatt
    - aktualisiert Konsolidierung pro Lieferant in separater Datei
    """
    if not st.sent:
        return

    round_files = list_round_answer_files(cfg, st.round_id)
    if not round_files:
        return

    for sid in st.sent.keys():
        if sid not in round_files:
            continue

        meta = st.responses.get(sid) or {}
        if meta.get("round_id") == st.round_id and meta.get("input_file"):
            if Path(meta["input_file"]).name == round_files[sid].name:
                continue

        inp = round_files[sid]
        print(f"[INFO] Vorhandene Antwortdatei (passende Runde {st.round_id}): {inp.name}")

        report = build_report(cfg, scale, st.round_id, sid, inp)
        if not report:
            print(f"[WARN] Datei nicht verarbeitbar (leer/defekt): {inp.name}")
            continue

        sup = supplier_name(cfg, sid)
        if sup:
            _, sum_val = upsert_into_master_nutz(cfg, report, sup, sheet_name)

            try:
                update_supplier_konsolidierung(cfg, sup, sheet_name, sum_val)
            except Exception as e:
                print(f"[WARN] Konsolidierung konnte nicht aktualisiert werden ({sup}): {e}")
                traceback.print_exc()

        prev_count = int(meta.get("count", 0))
        st.responses[sid] = {
            "round_id": st.round_id,
            "last_processed_at": time.time(),
            "count": prev_count + 1,
            "input_file": str(inp.resolve()),
            "source": "folder_in_ingest",
        }
        st.save()
        print(st.status())


def phase_poll_folder(cfg: Config, outlook: OutlookUI, scale: dict, st: State, sheet_name: str, folder="rpa_supplier_evaluation", poll=10):
    print(f"\n[PHASE 2] Polling Ordner '{folder}' (Runde {st.round_id}) ...")
    current_rid = st.round_id
    current_rid_l = current_rid.lower()
    outlook.open_folder(folder)

    def extract_rid_sid(text: str) -> Tuple[Optional[str], Optional[str]]:
        t = text or ""
        m_r = re.search(r"\bR(\d{8})\b", t, flags=re.I)
        if not m_r:
            m_r = re.search(r"\b(\d{8})\b", t)
        m_s = re.search(r"\b(K_\d+)\b", t, flags=re.I)
        return (m_r.group(1) if m_r else None), (m_s.group(1).upper() if m_s else None)

    ingest_existing_round_answers(cfg, scale, st, sheet_name)
    if st.all_done():
        print(f"\n[OK] Alle Antworten lagen bereits als Dateien vor. {st.status()}")
        return

    while not st.all_done():
        ingest_existing_round_answers(cfg, scale, st, sheet_name)
        if st.all_done():
            break

        outlook.refresh_folder()
        rows = outlook.message_rows()
        total = rows.count() if rows else 0
        if total == 0:
            time.sleep(poll)
            continue

        processed = False
        for i in range(min(total, 80)):
            row = rows.nth(i)

            try:
                row.click(timeout=2500)
                outlook.page.locator('div[role="main"]').first.wait_for(state="visible", timeout=6000)
                txt = outlook.page.locator('div[role="main"]').first.inner_text(timeout=2000) or ""
            except Exception:
                continue

            rid_found, sid = extract_rid_sid(txt)

            if not rid_found or rid_found.lower() != current_rid_l:
                try:
                    row_text = (row.inner_text(timeout=1000) or "").lower()
                except Exception:
                    row_text = ""
                if current_rid_l not in row_text and ("r" + current_rid_l) not in row_text:
                    continue
                rid_found = current_rid

            if not sid:
                try:
                    row_text = row.inner_text(timeout=1000) or ""
                except Exception:
                    row_text = ""
                m = re.search(r"\b(K_\d+)\b", row_text, flags=re.I)
                sid = m.group(1).upper() if m else None

            if not sid or sid not in st.sent:
                continue

            if st.has_response_for_round(sid, current_rid):
                continue

            downloaded = outlook.open_download_xlsx(row, cfg.folder_in)
            if not downloaded:
                continue

            downloaded_path = Path(downloaded)
            dest = cfg.folder_in / f"Antwort_{sid}_R{rid_found}_{int(time.time())}.xlsx"
            try:
                downloaded_path.replace(dest)
                used = dest
            except Exception:
                used = downloaded_path

            print(f"[OK] XLSX: {used.name}")

            try:
                report = build_report(cfg, scale, rid_found, sid, used)
                if not report:
                    raise RuntimeError("Bericht konnte nicht erstellt werden (leer/defekt).")

                sup = supplier_name(cfg, sid)
                if sup:
                    _, sum_val = upsert_into_master_nutz(cfg, report, sup, sheet_name)
                    try:
                        update_supplier_konsolidierung(cfg, sup, sheet_name, sum_val)
                    except Exception as e:
                        print(f"[WARN] Konsolidierung konnte nicht aktualisiert werden ({sup}): {e}")
                        traceback.print_exc()

                prev_meta = st.responses.get(sid) or {}
                st.responses[sid] = {
                    "round_id": rid_found,
                    "last_processed_at": time.time(),
                    "count": int(prev_meta.get("count", 0)) + 1,
                    "input_file": str(Path(used).resolve()),
                    "source": "outlook_web",
                }
                st.save()
                print(st.status())

                outlook.archive_current_mail()
                processed = True
                break

            except Exception as e:
                print(f"[ERROR] Verarbeitung fehlgeschlagen für {sid}: {e}")
                traceback.print_exc()
                processed = True
                break

        if not processed:
            time.sleep(poll)

    print(f"\n[OK] Alle Antworten liegen vor. {st.status()}")


def phase_send_final(cfg: Config, outlook: OutlookUI, st: State, sheet_name: str):
    if st.final_mail_sent:
        return

    attach_nutz = cfg.file_nutz_master.resolve()
    attach_kons = cfg.file_kons_master.resolve()

    attachments: List[Path] = []
    if attach_nutz.exists():
        attachments.append(attach_nutz)
    else:
        print("[WARN] Keine zentrale Nutzwertanalyse-Datei gefunden.")

    if attach_kons.exists():
        attachments.append(attach_kons)
    else:
        print("[WARN] Keine Konsolidierung-Datei gefunden (Konsolidierung_Zentral.xlsx).")

    if not attachments:
        print("[WARN] Keine Anhänge vorhanden, keine Abschlussmail.")
        return

    subject = f"SCM Nutzwertanalyse | {sheet_name} | Runde {st.round_id}"
    body = (
        f"Hallo,\n\n"
        f"anbei die zentrale Nutzwertanalyse sowie die Lieferanten-Konsolidierung.\n"
        f"Bewertungszeitraum: {sheet_name}\n"
        f"(Runde {st.round_id})\n\n"
        f"Grüße\nSCM Bot"
    )

    outlook.open_mail()
    outlook.new_mail_with_attachments(cfg.send_to_final, subject, body, attachments)
    st.final_mail_sent = True
    st.save()
    print(f"[FINISH] Abschlussmail gesendet an {cfg.send_to_final} (Anhänge: {', '.join(p.name for p in attachments)})")


# =========================
# MAIN
# =========================

def main():
    cfg = Config()
    cfg.ensure_dirs()

    m, y = prompt_month_year()
    sheet_name = sheet_name_from_my(m, y)  # z.B. 01-2026

    st = State.load_or_new(cfg)
    st.ensure_not_stuck_on_finished_round()

    scale = load_scale(cfg.file_scale)

    with sync_playwright() as p:
        user_data_dir = cfg.root / "Playwright_SCM_Profile"
        browser = p.chromium.launch_persistent_context(
            builtins.str(user_data_dir),
            headless=False,
            slow_mo=600,
            accept_downloads=True,
        )
        page = browser.new_page()
        outlook = OutlookUI(page)

        phase_send_links(cfg, outlook, st)
        phase_poll_folder(cfg, outlook, scale, st, sheet_name, folder="rpa_supplier_evaluation", poll=10)
        phase_send_final(cfg, outlook, st, sheet_name)

if __name__ == "__main__":
    main()
