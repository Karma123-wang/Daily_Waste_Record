"""
Waste Collection & Recovery Analyser – reads the 'Daily Waste Weigh Record' workbook.
Run:  pip install -r requirements.txt
      streamlit run app.py
"""
import io
import json

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from waste_calc import (MATERIALS, auto_findings, by_area, by_vehicle, daily_totals, load_workbook_data,
                        material_table, monthly_summary, report_check)
from waste_report import MAT_COLOURS, build_pdf, build_pptx, build_xlsx

st.set_page_config(page_title="Waste Analyser", page_icon="♻️", layout="wide")
BLUE, NAVY, GREEN, ORANGE, RED = "#118DFF", "#12239E", "#1AAB40", "#E66C37", "#D64550"
FONT = "Segoe UI, Helvetica Neue, Arial, sans-serif"

st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{ background:#F3F2F1; }}
  .block-container {{ padding-top:1.2rem; }}
  div[class*="st-key-card"] {{ background:#FFF; border-radius:6px; padding:14px 16px 6px 16px;
      box-shadow:0 1px 3px rgba(0,0,0,.12); }}
  .pbi-header {{ background:{BLUE}; color:#FFF; padding:16px 22px; border-radius:6px; border-left:6px solid {NAVY};
      margin-bottom:14px; font-family:{FONT}; }}
  .pbi-header h1 {{ font-size:26px; margin:0; color:#FFF; font-weight:700; }}
  .pbi-header p {{ margin:4px 0 0 0; color:#E3F1FF; font-size:14px; }}
  .kpi {{ background:#FFF; border-radius:6px; padding:12px 14px; height:118px; box-shadow:0 1px 3px rgba(0,0,0,.12);
      border-top:4px solid {BLUE}; font-family:{FONT}; }}
  .kpi.green {{ border-top-color:{GREEN}; }} .kpi.green .v {{ color:{GREEN}; }}
  .kpi .v {{ font-size:28px; font-weight:700; color:#252423; line-height:1.1; }}
  .kpi .l {{ font-size:12.5px; color:#605E5C; margin-top:6px; line-height:1.25; }}
  .card-title {{ font-family:{FONT}; font-size:15px; font-weight:600; color:#252423; margin-bottom:0; }}
</style>""", unsafe_allow_html=True)


def card_title(t):
    st.markdown(f'<p class="card-title">{t}</p>', unsafe_allow_html=True)


def kpi(col, value, label, green=False):
    col.markdown(f'<div class="kpi {"green" if green else ""}"><div class="v">{value}</div>'
                 f'<div class="l">{label}</div></div>', unsafe_allow_html=True)


def style(fig, height=340):
    fig.update_layout(height=height, font=dict(family=FONT, size=12, color="#252423"), paper_bgcolor="white",
                      plot_bgcolor="white", margin=dict(l=10, r=10, t=30, b=10),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title_text=""))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#EDEBE9", zeroline=False)
    return fig


def t(kg):
    return f"{kg / 1000:,.1f} t"


# ---------------- Data ----------------
st.sidebar.header("Data")
file = st.sidebar.file_uploader("Daily Waste Weigh Record (.xlsx)", type=["xlsx"])
if not file:
    st.markdown('<div class="pbi-header"><h1>Waste Collection & Recovery Analyser</h1>'
                '<p>Upload the Daily Waste Weigh Record workbook in the sidebar</p></div>', unsafe_allow_html=True)
    st.info("⬅️ Upload the workbook (.xlsx). The app reads the 'Daily Waste Weigh' register, the RVM and scrap-dealer "
            "sheets, and the food / industrial waste figures from the monthly report sheets.")
    st.stop()


@st.cache_data(show_spinner="Reading the workbook…")
def read(blob):
    col, rec, ex = load_workbook_data(io.BytesIO(blob))
    chk = report_check(io.BytesIO(blob), rec)
    return col, rec, ex, chk


try:
    col_all, rec_all, extras0, check = read(file.getvalue())
except ValueError as e:
    st.error(str(e))
    st.stop()

all_months = list(col_all["Month"].drop_duplicates().sort_values().dt.strftime("%b %Y"))
if len(all_months) > 1:
    start, end = st.sidebar.select_slider("Months", options=all_months, value=(all_months[0], all_months[-1]))
    months = all_months[all_months.index(start): all_months.index(end) + 1]
else:
    months = all_months
areas_all = sorted(col_all["Area"].unique())
areas_sel = st.sidebar.multiselect("Collection areas", areas_all, default=areas_all)

with st.sidebar.expander("Food & industrial waste (kg)"):
    st.caption("Taken from the monthly report sheets. Edit if a figure is missing or changes.")
    ex_edit = extras0.copy()
    ex_edit["Month"] = ex_edit["Month"].dt.strftime("%b %Y")
    ex_edit = pd.DataFrame({"Month": all_months}).merge(ex_edit, on="Month", how="left").fillna(0)
    ex_edit = st.data_editor(ex_edit, hide_index=True, disabled=["Month"], key="extras",
                             use_container_width=True)
extras = ex_edit.assign(Month=pd.PeriodIndex(pd.to_datetime(ex_edit["Month"], format="%b %Y"), freq="M"))

col = col_all[col_all["Area"].isin(areas_sel) & col_all["Month"].dt.strftime("%b %Y").isin(months)]
rec = rec_all[rec_all["Month"].dt.strftime("%b %Y").isin(months)] if not rec_all.empty else rec_all
if col.empty:
    st.warning("No collection records for the selected months / areas.")
    st.stop()
m = monthly_summary(col, rec, extras)
if len(areas_sel) < len(areas_all):
    # recyclables are not recorded by area, so a recovery rate for some areas only would be misleading
    m["Recovery %"] = float("nan")
areas = by_area(col)
veh = by_vehicle(col)
mats = material_table(rec)
daily = daily_totals(col)
period = f"{m.index[0]} – {m.index[-1]}" if len(m) > 1 else m.index[0]
data = dict(monthly=m, areas=areas, vehicles=veh, materials=mats, daily=daily)

notes_data = []
if m["Partial month"].any():
    notes_data.append(f"Collection records start on {col_all['Date'].min():%d %b %Y}, so "
                      f"{', '.join(m.index[m['Partial month']])} is a partial month.")
if (~m["Recycling data"]).any():
    notes_data.append(f"No recycling records for {', '.join(m.index[~m['Recycling data']])} – recovery rate not shown.")
missing_area = col[col["Area"] == "Area not recorded"]
if len(missing_area):
    notes_data.append(f"{len(missing_area)} record(s) have no collection area in the register ("
                      + "; ".join(f"{d:%d %b} – {v} – {k:,.0f} kg" for d, v, k in
                                  zip(missing_area["Date"], missing_area["Vehicle"], missing_area["Kg"]))
                      + "); counted as 'Area not recorded'.")
if len(areas_sel) < len(areas_all):
    notes_data.append(f"Collections filtered to {len(areas_sel)} of {len(areas_all)} areas. Recyclables are not "
                      "recorded by area, so the recovery rate is only shown when all areas are selected.")

st.markdown(f'<div class="pbi-header"><h1>Waste Collection & Recovery Dashboard</h1>'
            f'<p>{period} · {len(areas_sel)} collection areas · {int(m["Collection trips"].sum())} collection trips</p>'
            f'</div>', unsafe_allow_html=True)

# ---------------- Cards ----------------
month = st.selectbox("Month shown on the cards", list(m.index), index=len(m) - 1)
r = m.loc[month]
c = st.columns(4)
kpi(c[0], t(r["Waste collected (kg)"]), "Waste collected" + (" (partial month)" if r["Partial month"] else ""))
kpi(c[1], f"{int(r['Collection trips'])}", f"Collection trips on {int(r['Collection days'])} days")
kpi(c[2], f"{r['Avg per calendar day (kg)']:,.0f} kg", "Collected per calendar day")
kpi(c[3], t(r["Total waste generated (kg)"]), "Total waste generated<br>(collected + recovered + industrial)")
st.write("")
c = st.columns(4)
kpi(c[0], t(r["Recyclables recovered (kg)"]), "Recyclables recovered<br>(RVM + scrap dealers)", True)
kpi(c[1], t(r["Food waste recovered (kg)"]), "Food waste recovered", True)
kpi(c[2], "no data" if pd.isna(r["Recovery %"]) else f"{r['Recovery %']:.1f}%",
    "Recovery rate<br>(recovered ÷ generated)", True)
kpi(c[3], f"{r['Generated per day (kg)']:,.0f} kg", "Waste generated per day")
for n in notes_data:
    st.caption("ℹ️ " + n)
st.write("")

tabs = st.tabs(["Monthly summary", "Trends", "By area", "Recyclables", "Vehicles", "🔍 Report check",
                "✏️ Edit & export"])

# ---------------- Monthly summary ----------------
with tabs[0]:
    with st.container(key="card_month"):
        card_title("Monthly summary")
        show = m.drop(columns=["Partial month", "Recycling data"])
        fmt = {c_: "{:,.0f}" for c_ in show.columns}
        fmt["Recovery %"] = "{:.1f}"
        st.dataframe(show.T.style.format("{:,.1f}", na_rep="no data"), use_container_width=True, height=520)
        st.caption("Recovery % = (recyclables + food waste) ÷ (collected + recyclables + food waste + industrial), "
                   "the same method as the workbook's monthly report sheets.")

# ---------------- Trends ----------------
with tabs[1]:
    a1, a2 = st.columns([1.3, 1])
    with a1:
        with st.container(key="card_tr1"):
            card_title("Collected vs recovered (tonnes)")
            fig = go.Figure()
            fig.add_bar(x=list(m.index), y=m["Waste collected (kg)"] / 1000, name="Collected (to landfill)",
                        marker_color=BLUE, texttemplate="%{y:,.1f}", textposition="outside")
            fig.add_bar(x=list(m.index), y=m["Total recovered (kg)"] / 1000, name="Recovered (recycled + food)",
                        marker_color=GREEN, texttemplate="%{y:,.1f}", textposition="outside")
            fig.update_layout(barmode="group")
            st.plotly_chart(style(fig), use_container_width=True, key="tr1")
    with a2:
        with st.container(key="card_tr2"):
            card_title("Recovery rate")
            full = m[m["Recycling data"]]
            fig = go.Figure()
            fig.add_scatter(x=list(full.index), y=full["Recovery %"], mode="lines+markers+text",
                            line=dict(color=GREEN, width=3), marker=dict(size=10),
                            text=[f"{v:.1f}%" for v in full["Recovery %"]], textposition="top center")
            fig.update_yaxes(range=[0, max(60, full["Recovery %"].max() * 1.3 if len(full) else 60)], ticksuffix="%")
            st.plotly_chart(style(fig), use_container_width=True, key="tr2")
    st.write("")
    with st.container(key="card_tr3"):
        card_title("Waste collected per day (kg)")
        fig = go.Figure()
        fig.add_bar(x=daily.index, y=daily.values, marker_color=BLUE, hovertemplate="%{x|%d %b}: %{y:,.0f} kg<extra></extra>")
        fig.add_hline(y=daily.mean(), line_dash="dash", line_color=ORANGE,
                      annotation_text=f"average {daily.mean():,.0f} kg per collection day")
        st.plotly_chart(style(fig, 330), use_container_width=True, key="tr3")

# ---------------- By area ----------------
with tabs[2]:
    with st.container(key="card_area"):
        card_title(f"Waste collected by area – {period}")
        st.dataframe(areas.style.format({c_: "{:,.0f}" for c_ in areas.columns} |
                                        {"Share %": "{:.1f}", "Collections per week": "{:.1f}",
                                         "Avg per collection (kg)": "{:,.0f}"}), use_container_width=True)
        fig = go.Figure()
        fig.add_bar(y=list(areas.index), x=areas["Waste collected (kg)"] / 1000, orientation="h", marker_color=BLUE,
                    text=[f"{v / 1000:,.1f} t ({s:.0f}%)" for v, s in zip(areas["Waste collected (kg)"], areas["Share %"])],
                    textposition="outside")
        fig.update_yaxes(autorange="reversed")
        fig.update_xaxes(title="Tonnes")
        st.plotly_chart(style(fig, 70 + 36 * len(areas)), use_container_width=True, key="area")
    st.write("")
    with st.container(key="card_area_month"):
        card_title("Tonnes by area and month")
        xt = col.pivot_table(index="Area", columns=col["Month"].dt.strftime("%b %Y"), values="Kg", aggfunc="sum",
                             fill_value=0)[[x for x in m.index]] / 1000
        xt["Total"] = xt.sum(axis=1)
        xt = xt.sort_values("Total", ascending=False)
        xt.loc["TOTAL"] = xt.sum()
        st.dataframe(xt.style.format("{:,.1f}"), use_container_width=True)

# ---------------- Recyclables ----------------
with tabs[3]:
    if mats.empty:
        st.info("No recycling records for the selected months.")
    else:
        b1, b2 = st.columns([1.4, 1])
        with b1:
            with st.container(key="card_rec1"):
                card_title("Recyclables by collector (tonnes)")
                fig = go.Figure()
                for mat in [x for x in MATERIALS if x in mats.columns]:
                    fig.add_bar(y=list(mats.index), x=mats[mat] / 1000, name=mat, orientation="h",
                                marker_color=MAT_COLOURS.get(mat), hovertemplate="%{y}<br>" + mat + ": %{x:,.2f} t<extra></extra>")
                fig.update_layout(barmode="stack")
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(style(fig, 90 + 46 * len(mats)), use_container_width=True, key="rec1")
        with b2:
            with st.container(key="card_rec2"):
                card_title("By material")
                tot = mats.drop(columns="Total (kg)").sum().sort_values(ascending=False)
                fig = go.Figure(go.Pie(labels=list(tot.index), values=tot.values / 1000, hole=0.6, sort=False,
                                       marker=dict(colors=[MAT_COLOURS.get(k) for k in tot.index]),
                                       texttemplate="%{percent:.0%}", hovertemplate="%{label}: %{value:,.1f} t<extra></extra>"))
                fig.add_annotation(text=f"<b>{tot.sum() / 1000:,.1f} t</b>", showarrow=False, font=dict(size=18))
                st.plotly_chart(style(fig, 90 + 46 * len(mats)), use_container_width=True, key="rec2")
        st.write("")
        with st.container(key="card_rec3"):
            card_title("Recyclables by collector and material (kg)")
            st.dataframe(mats.style.format("{:,.1f}"), use_container_width=True)

# ---------------- Vehicles ----------------
with tabs[4]:
    with st.container(key="card_veh"):
        card_title(f"Vehicles – {period}")
        st.dataframe(veh.style.format({"Waste carried (kg)": "{:,.0f}", "Avg per trip (kg)": "{:,.0f}"}),
                     use_container_width=True)
        fig = go.Figure()
        fig.add_bar(y=list(veh.index), x=veh["Waste carried (kg)"] / 1000, orientation="h", marker_color=NAVY,
                    text=[f"{v / 1000:,.1f} t · {n} trips" for v, n in zip(veh["Waste carried (kg)"], veh["Trips"])],
                    textposition="outside")
        fig.update_yaxes(autorange="reversed")
        st.plotly_chart(style(fig, 70 + 36 * len(veh)), use_container_width=True, key="veh")

# ---------------- Report check ----------------
with tabs[5]:
    with st.container(key="card_check"):
        card_title("Monthly report sheets vs daily sheets – recyclables (kg)")
        if check.empty:
            st.info("No monthly report sheets with a recovery table were found to check.")
        else:
            st.caption("The app adds up the daily RVM and scrap-dealer sheets. Rows in red show where the figure "
                       "typed in a monthly report sheet is different.")
            st.dataframe(check.style.format({"Monthly report (kg)": "{:,.2f}", "Daily sheets total (kg)": "{:,.2f}",
                                             "Difference (kg)": "{:+,.2f}"}, na_rep="–")
                         .apply(lambda r: ["background-color:#FBE3E5;color:#D64550;font-weight:600"
                                           if pd.notna(r["Difference (kg)"]) and abs(r["Difference (kg)"]) > 1 else ""
                                           for _ in r], axis=1), use_container_width=True, hide_index=True)

# ---------------- Edit & export ----------------
SECTIONS = {"summary": "Summary cards", "notes": "Key findings & recommendations", "table": "Monthly summary table",
            "trends": "Trend charts", "areas": "By area", "recycling": "Recyclables", "vehicles": "Vehicles"}
ss = st.session_state
ss.setdefault("ed_title", "Waste Collection & Recovery Report")
ss.setdefault("ed_role", "")
ss.setdefault("ed_author", "")
ss.setdefault("ed_recs", "")
ss.setdefault("ed_sections", list(SECTIONS))
with tabs[6]:
    suggested = auto_findings(m, areas, mats)
    if ss.pop("refill", False):
        ss["auto_f"] = True
        ss["ed_findings"] = ""
    if ss.get("ed_findings", "") not in ("", ss.get("last_sugg", "")):
        ss["auto_f"] = False
    if ss.get("auto_f", True):
        ss["ed_findings"] = suggested
    ss["last_sugg"] = suggested

    with st.expander("📂 Load saved edits"):
        saved = st.file_uploader("Saved edits (.json)", type=["json"], key="edits_file")
        if saved is not None and ss.get("edits_id") != saved.file_id:
            try:
                e = json.loads(saved.getvalue().decode("utf-8"))
                for k in ("title", "role", "author", "findings", "recs"):
                    if k in e:
                        ss[f"ed_{k}"] = e[k]
                if e.get("sections"):
                    ss["ed_sections"] = [x for x in e["sections"] if x in SECTIONS]
                ss["edits_id"], ss["auto_f"] = saved.file_id, False
                st.success("Edits loaded.")
            except Exception:
                st.error("This file could not be read.")
    with st.container(key="card_cover"):
        card_title("1 · Cover page")
        x1, x2, x3 = st.columns([2, 1.4, 1.2])
        x1.text_input("Report title", key="ed_title")
        x2.text_input("Organisation / unit", key="ed_role")
        x3.text_input("Prepared by", key="ed_author")
    st.write("")
    with st.container(key="card_notes"):
        card_title("2 · Key findings & recommendations")
        st.caption("One point per line. Findings are drafted from the data – edit or add to them.")
        n1, n2 = st.columns(2)
        n1.text_area("Key findings", key="ed_findings", height=230)
        n2.text_area("Recommendations", key="ed_recs", height=230,
                     placeholder="e.g. Introduce wet/dry segregation in all areas, not only Core Town")
        if n1.button("↺ Refill findings from the data"):
            ss["refill"] = True
            st.rerun()
    st.write("")
    with st.container(key="card_export"):
        card_title("3 · Pages and download")
        st.multiselect("Pages to include", list(SECTIONS), key="ed_sections", format_func=lambda k: SECTIONS[k])
        lines = lambda s: [x.strip(" •-*\t") for x in str(s).splitlines() if x.strip(" •-*\t")]
        notes = {"findings": lines(ss["ed_findings"]), "recommendations": lines(ss["ed_recs"])}
        cover = {"title": ss["ed_title"].strip(), "role": ss["ed_role"].strip(), "author": ss["ed_author"].strip(),
                 "period_line": f"{period}  ·  {t(m['Waste collected (kg)'].sum())} collected"}
        sections = ss["ed_sections"] or list(SECTIONS)
        edits = {"title": ss["ed_title"], "role": ss["ed_role"], "author": ss["ed_author"],
                 "findings": ss["ed_findings"], "recs": ss["ed_recs"], "sections": sections}
        sig = json.dumps([edits, list(m.index), month, areas_sel, ex_edit.to_dict(), file.name], sort_keys=True,
                         default=str)
        if ss.get("sig") != sig:
            for ext in ("pdf", "pptx", "xlsx"):
                ss.pop(ext, None)
            ss["sig"] = sig
        stem = f"{ss['ed_title'] or 'Waste Report'} - {period}".replace("/", "-")
        exports = [("PDF", "pdf", "application/pdf",
                    lambda: build_pdf(data, month, cover, notes, sections, notes_data)),
                   ("PowerPoint", "pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    lambda: build_pptx(data, month, cover, notes, sections, notes_data)),
                   ("Excel (tables)", "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    lambda: build_xlsx(data, cover, notes_data))]
        for cc, (label, ext, mime, make) in zip(st.columns(3), exports):
            with cc:
                if st.button(f"Create {label}", use_container_width=True):
                    with st.spinner(f"Building {label}…"):
                        ss[ext] = make()
                if ext in ss:
                    st.download_button(f"⬇️ Download {label}", ss[ext], f"{stem}.{ext}", mime,
                                       use_container_width=True)
        st.caption(f"The summary page uses the month chosen above the cards ({month}).")
        st.divider()
        st.download_button("💾 Save my edits (to reuse next month)", json.dumps(edits, indent=2, ensure_ascii=False),
                           "Waste report edits.json", "application/json")
