"""Read the 'Daily Waste Weigh Record' workbook and calculate waste collection and recovery figures.

* Collections  – from the 'Daily Waste Weigh' register (month blocks side by side:
                 Sl. No. | Date | Dhemkong / Collection Area | Vehicle No. | Total Mixed Waste (kg)).
* Recyclables  – from the RVM and scrap-dealer sheets (header row 'SL | Date: | <material columns> | Qty(kg) per day').
* Recovery %   – recyclables recovered ÷ (waste collected + recyclables recovered).
"""
from __future__ import annotations

import datetime as dt
import re

import pandas as pd
from openpyxl import load_workbook

MATERIALS = ["PET bottles", "Plastic", "Metal & cans", "Glass", "Beer bottles", "Carton & paper", "Others"]


# ---------------- helpers ----------------

def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


MONTHS = {m: i for i, m in enumerate(["january", "february", "march", "april", "may", "june", "july", "august",
                                       "september", "october", "november", "december"], start=1)}


def _month_label(row) -> int | None:
    """Rows such as 'August', 'July Month Data', 'Date: August Month' mark which month the next table is for."""
    txt = " ".join(str(v) for v in row if isinstance(v, str)).lower()
    if not txt or len(txt) > 60:
        return None
    for name, num in MONTHS.items():
        if re.search(rf"\b{name}\b", txt):
            return num
    return None


def _date(v):
    if isinstance(v, dt.datetime):
        return pd.Timestamp(v).normalize()
    if isinstance(v, dt.date):
        return pd.Timestamp(v)
    if isinstance(v, str):
        m = re.search(r"\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}", v)   # '01/07/26 to 31/07/26' -> first date
        if not m:
            return None
        d = pd.to_datetime(m.group(0), dayfirst=True, errors="coerce")
        return None if pd.isna(d) else d.normalize()
    return None


def _clean_area(a: str) -> tuple[str, str]:
    """'Core Town (Dry)' -> ('Core Town', 'Dry'); other areas -> (area, 'Mixed')."""
    a = re.sub(r"\s+", " ", str(a)).strip().rstrip(".")
    m = re.match(r"^(.*?)\s*\((dry|wet)\)\s*$", a, flags=re.I)
    if m:
        return m.group(1).strip(), m.group(2).capitalize()
    return a, "Mixed"


def _material(header: str) -> str | None:
    """Map a recyclables column header to a material; None for counts, totals and other columns."""
    h = str(header or "").strip().lower()
    if not h or h.startswith(("sl", "date", "qty", "total")) or "nos" in h or "(no" in h:
        return None
    if "beer" in h:                                   # beer bottles: only the weight column
        return "Beer bottles" if ("kg" in h or "weigh" in h) else None
    if "pet" in h:
        return "PET bottles"
    if "plastic" in h:
        return "Plastic"
    if "metal" in h or "can" in h:
        return "Metal & cans"
    if "glass" in h:
        return "Glass"
    if "carton" in h or "paper" in h:
        return "Carton & paper"
    if "other" in h:
        return "Others"
    return None


def _source_name(sheet: str) -> str:
    s = re.sub(r"\s+", " ", sheet).strip()
    if s.upper() == "RVM":
        return "RVM (reverse vending machine)"
    return re.sub(r"(?i)scrap dealer", "scrap dealer", s)


# ---------------- collections ----------------

def parse_collections(wb) -> pd.DataFrame:
    sheet = next((ws for ws in wb.worksheets if ws.title.strip().lower() == "daily waste weigh"), None)
    if sheet is None:
        raise ValueError("The workbook has no 'Daily Waste Weigh' sheet.")
    rows = [list(r) for r in sheet.iter_rows(values_only=True)]
    out = []
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            if not (isinstance(v, str) and v.strip().lower().startswith("sl")):
                continue
            hdr = [str(x or "").strip().lower() for x in r[j:j + 6]]
            if not (len(hdr) > 4 and hdr[1].startswith("date") and "area" in hdr[2] or "dhemkong" in hdr[2]):
                continue
            blanks = 0
            for rr in rows[i + 1:]:
                cells = (rr + [None] * 6)[j:j + 5]
                d, area, veh, kg = _date(cells[1]), cells[2], cells[3], _num(cells[4])
                if d is not None and kg is not None and not (area and str(area).strip()):
                    area = "Area not recorded"            # keep the weight; flag it in the notes
                if d is None or kg is None or not area:
                    blanks += 1
                    if blanks > 3:
                        break
                    continue
                blanks = 0
                place, wtype = _clean_area(area)
                out.append({"Date": d, "Area": place, "Type": wtype, "Vehicle": str(veh or "").strip() or "Unknown",
                            "Kg": kg})
    df = pd.DataFrame(out)
    if df.empty:
        raise ValueError("No collection rows were found in 'Daily Waste Weigh'.")
    df["Month"] = df["Date"].dt.to_period("M")
    return df.sort_values("Date").reset_index(drop=True)


# ---------------- recyclables ----------------

def parse_recycling(wb) -> pd.DataFrame:
    out = []
    for ws in wb.worksheets:
        name = ws.title.strip()
        low = name.lower()
        if not ("scrap" in low or low == "rvm") or low.startswith("copy"):
            continue
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        label_month = None
        for i, r in enumerate(rows):
            lm = _month_label(r[:3]) if r else None
            if lm:
                label_month = lm
            if not (r and isinstance(r[0], str) and r[0].strip().upper() == "SL"
                    and len(r) > 1 and str(r[1] or "").strip().lower().startswith("date")):
                continue
            cols = {j: _material(h) for j, h in enumerate(r)}
            blanks = 0
            for rr in rows[i + 1:]:
                if rr and isinstance(rr[0], str) and rr[0].strip().upper() == "SL":
                    break                                    # next month's table
                lm = _month_label(rr[:3])
                if lm and not any(_num(x) for x in rr[2:]):
                    label_month = lm                          # e.g. a row that just says 'August'
                    continue
                raw = rr[1] if len(rr) > 1 else None
                d = _date(raw)
                if d is not None and isinstance(raw, str) and re.search(r"\bto\b", raw) and label_month:
                    d = pd.Timestamp(d.year, label_month, 1)  # monthly lump sum: use the labelled month
                if d is None:
                    blanks += 1
                    if blanks > 3 or (len(rr) > 1 and str(rr[1] or "").strip().upper() == "TOTAL"):
                        break
                    continue
                blanks = 0
                for j, mat in cols.items():
                    if mat and j < len(rr):
                        kg = _num(rr[j])
                        if kg:
                            out.append({"Date": d, "Source": _source_name(name), "Material": mat, "Kg": kg})
    df = pd.DataFrame(out, columns=["Date", "Source", "Material", "Kg"])
    if not df.empty:
        df["Month"] = df["Date"].dt.to_period("M")
    return df


def parse_extras(wb, year: int) -> pd.DataFrame:
    """Food waste lifted and industrial waste, taken from the '<Month> Month Report' sheets (no daily sheets exist)."""
    out = []
    for ws in wb.worksheets:
        title = ws.title.strip().lower()
        if "month report" not in title:
            continue
        mon = next((n for name, n in MONTHS.items() if name in title), None)
        if not mon:
            continue
        food = industrial = 0.0
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        in_industrial = False
        for r in rows:
            first = str(r[0] or "").strip().lower() if r else ""
            if "food" in first:
                food += max([_num(v) or 0 for v in r[1:10]] or [0])
            if first.startswith("industrial"):
                in_industrial = True
            elif in_industrial and first == "total":
                industrial = max([_num(v) or 0 for v in r[1:10]] or [0])
                in_industrial = False
        out.append({"Month": pd.Period(year=year, month=mon, freq="M"), "Food waste recovered (kg)": food,
                    "Industrial waste (kg)": industrial})
    return pd.DataFrame(out, columns=["Month", "Food waste recovered (kg)", "Industrial waste (kg)"])


def load_workbook_data(file) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    wb = load_workbook(file, data_only=True, read_only=False)
    col = parse_collections(wb)
    return col, parse_recycling(wb), parse_extras(wb, int(col["Date"].dt.year.mode()[0]))


# ---------------- summaries ----------------

def monthly_summary(col: pd.DataFrame, rec: pd.DataFrame, extras: pd.DataFrame | None = None) -> pd.DataFrame:
    """Recovery % follows the monthly report method:
    recovered = recyclables + food waste;  generated = collected + recovered + industrial waste."""
    g = col.groupby("Month")
    m = pd.DataFrame({
        "Waste collected (kg)": g["Kg"].sum(),
        "Collection trips": g.size(),
        "Collection days": g["Date"].nunique(),
    })
    m["First date"] = g["Date"].min()
    m["Last date"] = g["Date"].max()
    first, last = col["Date"].min(), col["Date"].max()
    # calendar days covered: a month only counts from the first record (e.g. collection started 22 June)
    m["Days in period"] = [(min(p.end_time.normalize(), max(last, p.end_time.normalize() if p != last.to_period("M")
                                                              else last)) - max(p.start_time, first)).days + 1
                           for p in m.index]
    m["Avg per collection day (kg)"] = m["Waste collected (kg)"] / m["Collection days"]
    m["Avg per calendar day (kg)"] = m["Waste collected (kg)"] / m["Days in period"]
    recm = rec.groupby("Month")["Kg"].sum() if not rec.empty else pd.Series(dtype=float)
    m["Recyclables recovered (kg)"] = recm.reindex(m.index).fillna(0)
    ex = (extras.set_index("Month") if extras is not None and not extras.empty
          else pd.DataFrame(columns=["Food waste recovered (kg)", "Industrial waste (kg)"]))
    m["Food waste recovered (kg)"] = ex["Food waste recovered (kg)"].reindex(m.index).fillna(0).astype(float)
    m["Industrial waste (kg)"] = ex["Industrial waste (kg)"].reindex(m.index).fillna(0).astype(float)
    m["Total recovered (kg)"] = m["Recyclables recovered (kg)"] + m["Food waste recovered (kg)"]
    m["Total waste generated (kg)"] = m["Waste collected (kg)"] + m["Total recovered (kg)"] + m["Industrial waste (kg)"]
    m["Recovery %"] = m["Total recovered (kg)"] / m["Total waste generated (kg)"] * 100
    m["Recycling data"] = m["Recyclables recovered (kg)"] > 0
    m.loc[~m["Recycling data"], "Recovery %"] = float("nan")       # no recycling sheets for that month
    m["Generated per day (kg)"] = m["Total waste generated (kg)"] / m["Days in period"]
    m["Partial month"] = m["Days in period"] < [p.days_in_month for p in m.index]
    m.index = m.index.strftime("%b %Y")
    m.index.name = "Month"
    return m.drop(columns=["First date", "Last date"])


def by_area(col: pd.DataFrame, months=None) -> pd.DataFrame:
    d = col if not months else col[col["Month"].dt.strftime("%b %Y").isin(months)]
    g = d.groupby("Area")
    out = pd.DataFrame({"Waste collected (kg)": g["Kg"].sum(), "Collections": g.size(),
                        "Avg per collection (kg)": g["Kg"].mean(),
                        "Dry (kg)": d[d["Type"] == "Dry"].groupby("Area")["Kg"].sum(),
                        "Wet (kg)": d[d["Type"] == "Wet"].groupby("Area")["Kg"].sum()}).fillna(0)
    total = out["Waste collected (kg)"].sum()
    out["Share %"] = out["Waste collected (kg)"] / max(1, total) * 100
    weeks = max(1, d["Date"].nunique() and (d["Date"].max() - d["Date"].min()).days / 7)
    out["Collections per week"] = out["Collections"] / weeks
    return out.sort_values("Waste collected (kg)", ascending=False)


def by_vehicle(col: pd.DataFrame, months=None) -> pd.DataFrame:
    d = col if not months else col[col["Month"].dt.strftime("%b %Y").isin(months)]
    g = d.groupby("Vehicle")
    return pd.DataFrame({"Waste carried (kg)": g["Kg"].sum(), "Trips": g.size(), "Avg per trip (kg)": g["Kg"].mean(),
                         "Areas served": g["Area"].nunique()}).sort_values("Waste carried (kg)", ascending=False)


def material_table(rec: pd.DataFrame, months=None) -> pd.DataFrame:
    d = rec if (not months or rec.empty) else rec[rec["Month"].dt.strftime("%b %Y").isin(months)]
    if d.empty:
        return pd.DataFrame()
    t = d.pivot_table(index="Source", columns="Material", values="Kg", aggfunc="sum", fill_value=0)
    t = t[[m for m in MATERIALS if m in t.columns]]
    t["Total (kg)"] = t.sum(axis=1)
    return t.sort_values("Total (kg)", ascending=False)


def daily_totals(col: pd.DataFrame, months=None) -> pd.Series:
    d = col if not months else col[col["Month"].dt.strftime("%b %Y").isin(months)]
    return d.groupby("Date")["Kg"].sum()


def auto_findings(m: pd.DataFrame, areas: pd.DataFrame, mats: pd.DataFrame) -> str:
    if m.empty:
        return ""
    out = [f"{m['Waste collected (kg)'].sum() / 1000:,.1f} tonnes of waste collected in "
           f"{int(m['Collection trips'].sum())} collection trips ({m.index[0]} – {m.index[-1]})."]
    full = m[m["Recycling data"] & m["Recovery %"].notna()]
    if len(full):
        tot_rec = full["Total recovered (kg)"].sum()
        out.append(f"{tot_rec / 1000:,.1f} tonnes recovered in {', '.join(full.index)} (recyclables "
                   f"{full['Recyclables recovered (kg)'].sum() / 1000:,.1f} t + food waste "
                   f"{full['Food waste recovered (kg)'].sum() / 1000:,.1f} t) – "
                   f"{tot_rec / full['Total waste generated (kg)'].sum() * 100:.1f}% of all waste generated in those months.")
        missing = [i for i in m.index if i not in full.index]
        if missing:
            out.append(f"No recycling records for {', '.join(missing)}, so no recovery rate is shown for that period.")
        if len(full):
            out.append("Recovery rate by month: " + ", ".join(f"{i} {v:.1f}%" for i, v in full["Recovery %"].items()) + ".")
    if not areas.empty:
        top = areas.head(3)
        out.append("Largest collection areas: " + "; ".join(
            f"{a} ({r['Waste collected (kg)'] / 1000:,.1f} t, {r['Share %']:.0f}%)" for a, r in top.iterrows()) + ".")
    if not mats.empty:
        mt = mats.drop(columns="Total (kg)").sum().sort_values(ascending=False)
        out.append("Main recyclables: " + ", ".join(f"{k} {v / 1000:,.1f} t" for k, v in mt.head(3).items()) + ".")
    return "\n".join(out)


def report_check(file, rec: pd.DataFrame) -> pd.DataFrame:
    """Compare recyclables in the '<Month> Month Report' sheets with the totals of the daily dealer/RVM sheets."""
    wb = load_workbook(file, data_only=True)
    year = int(rec["Date"].dt.year.mode()[0]) if not rec.empty else pd.Timestamp.today().year
    daily = (rec.groupby([rec["Month"], "Source"])["Kg"].sum() if not rec.empty else pd.Series(dtype=float))
    out = []
    for ws in wb.worksheets:
        title = ws.title.strip().lower()
        mon = next((n for name, n in MONTHS.items() if name in title), None) if "month report" in title else None
        if not mon:
            continue
        period = pd.Period(year=year, month=mon, freq="M")
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        start = next((i for i, r in enumerate(rows) if str(r[0] or "").strip().lower().startswith("accumulator")), None)
        if start is None:
            continue
        for r in rows[start + 1:]:
            name = str(r[0] or "").strip()
            if not name or name.lower().startswith(("total", "average")):
                break
            reported = _num(r[8]) if len(r) > 8 else None
            key = name.split()[0].lower()
            match = [s for s in daily.index if s[0] == period and s[1].lower().startswith(key)] if len(daily) else []
            sheet_total = float(daily[match[0]]) if match else None
            out.append({"Month": period.strftime("%b %Y"), "Collector": name, "Monthly report (kg)": reported,
                        "Daily sheets total (kg)": sheet_total,
                        "Difference (kg)": None if reported is None or sheet_total is None else reported - sheet_total})
    df = pd.DataFrame(out)
    if not df.empty:
        tot = df.groupby("Month")[["Monthly report (kg)", "Daily sheets total (kg)"]].sum(min_count=1)
        tot["Difference (kg)"] = tot["Monthly report (kg)"] - tot["Daily sheets total (kg)"]
        tot["Collector"] = "TOTAL"
        df = pd.concat([df, tot.reset_index()], ignore_index=True)
    return df
