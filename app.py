import io
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

st.set_page_config(page_title="POT2026 Case Study Dashboard", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def pct_to_float(x):
    """Convert '40.6%' -> 0.406 ; keep numeric as-is."""
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace("%", "")
    try:
        return float(s) / 100.0
    except:
        return np.nan

def norm_stage(s):
    s = str(s).strip().lower()
    if "closed won" in s:
        return "Closed Won"
    if "closed lost" in s:
        return "Closed Lost"
    if s in ["lead", "new lead", "inbound lead"]:
        return "Lead"
    if "contact" in s:
        return "Contacted"
    if "qualif" in s:
        return "Qualified"
    if "negoti" in s:
        return "Negotiation"
    return s.title()

def clean_and_model(xlsx_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    sheets = xls.sheet_names
    dfs = {name: pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=name) for name in sheets}

    # Expect these 5 sheets
    wt = dfs["Website Traffic"].copy()
    sm = dfs["Social Media"].copy()
    em = dfs["Email Campaigns"].copy()
    sp = dfs["Sales Pipeline"].copy()
    ads = dfs["Ad Spend"].copy()

    # Dates
    wt["Week Starting"] = pd.to_datetime(wt["Week Starting"])
    sm["Week Starting"] = pd.to_datetime(sm["Week Starting"])
    em["Send Date"] = pd.to_datetime(em["Send Date"])
    for c in ["First Contact Date", "Last Activity Date", "Expected Close Date"]:
        sp[c] = pd.to_datetime(sp[c], errors="coerce")

    # Percent fields
    em["Open Rate"] = em["Open Rate"].apply(pct_to_float)
    em["CTR"] = em["CTR"].apply(pct_to_float)
    sm["Engagement Rate"] = sm["Engagement Rate"].apply(pct_to_float)

    wt["Bounce Rate"] = pd.to_numeric(wt["Bounce Rate"], errors="coerce")
    # Conversion Rate column looks like percent numbers (e.g., 2.07) -> 0.0207
    wt["Conversion Rate"] = pd.to_numeric(wt["Conversion Rate"], errors="coerce") / 100.0

    # Sales value cleanup
    sp["Deal Value (EUR)"] = pd.to_numeric(
        sp["Deal Value (EUR)"].astype(str).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce"
    )

    # Standardize lead source
    sp["Lead Source"] = sp["Lead Source"].astype(str).str.strip().str.lower()
    lead_map = {
        "linkedin outreach": "linkedin",
        "linkedin": "linkedin",
        "linkedin ads": "paid linkedin",
        "google ads": "paid google",
        "google": "paid google",
        "website inquiry": "website",
        "website": "website",
        "past attendee": "past attendee",
        "referral": "referral",
        "event": "event",
        "conference": "event",
    }
    sp["Lead Source Std"] = sp["Lead Source"].map(lead_map).fillna(sp["Lead Source"])

    sp["Ticket Category"] = np.select(
        [
            sp["Ticket Type"].str.contains("delegate", case=False, na=False),
            sp["Ticket Type"].str.contains("sponsor", case=False, na=False),
            sp["Ticket Type"].str.contains("speaker", case=False, na=False),
        ],
        ["Delegate", "Sponsor", "Speaker"],
        default="Other"
    )

    sp["Deal Stage"] = sp["Deal Stage"].astype(str).str.strip()
    sp["Stage_N"] = sp["Deal Stage"].apply(norm_stage)

    # Ads month parsing (supports "Oct 2025" / "October 2025" / "Nov 2025")
    ads["Month_dt"] = pd.to_datetime(ads["Month"], format="mixed", errors="coerce")
    for c in [
        "Budget (EUR)", "Spend (EUR)", "Impressions", "Clicks", "CPM (EUR)", "CPC (EUR)",
        "Conversions", "Cost per Conversion (EUR)"
    ]:
        ads[c] = pd.to_numeric(ads[c], errors="coerce")

    # -----------------------------
    # KPI Summary
    # -----------------------------
    kpi = {}
    kpi["Website Sessions"] = wt["Sessions"].sum()
    kpi["Website Ticket Inquiries"] = wt["Ticket Inquiry Conversions"].sum()
    kpi["Website CVR (weighted)"] = kpi["Website Ticket Inquiries"] / kpi["Website Sessions"]

    kpi["Social Impressions"] = sm["Impressions"].sum()
    kpi["Social Link Clicks"] = sm["Link Clicks"].sum(skipna=True)
    kpi["Social Engagement Rate (weighted)"] = sm["Engagements"].sum() / sm["Impressions"].sum()

    kpi["Email Delivered"] = em["Emails Delivered"].sum()
    kpi["Email Ticket Inquiries"] = em["Conversions (Ticket Inquiries)"].sum()
    kpi["Email Revenue Attributed"] = em["Revenue Attributed"].sum(skipna=True)
    kpi["Email Open Rate (weighted)"] = em["Opens"].sum() / em["Emails Delivered"].sum()
    kpi["Email CTR (weighted)"] = em["Clicks"].sum() / em["Emails Delivered"].sum()

    kpi["Ad Spend"] = ads["Spend (EUR)"].sum()
    kpi["Ad Conversions"] = ads["Conversions"].sum()
    kpi["Ad CPC (blended)"] = kpi["Ad Spend"] / ads["Clicks"].sum()
    kpi["Ad CPA (blended)"] = kpi["Ad Spend"] / kpi["Ad Conversions"]

    closed = sp[sp["Deal Stage"].str.contains("Closed", case=False, na=False)]
    won = sp[sp["Deal Stage"].str.contains("Closed Won", case=False, na=False)]
    lost = sp[sp["Deal Stage"].str.contains("Closed Lost", case=False, na=False)]

    kpi["Deals Closed Won"] = len(won)
    kpi["Deals Closed Lost"] = len(lost)
    kpi["Win Rate (closed deals)"] = len(won) / max(1, len(closed))
    kpi["Revenue Closed Won (EUR)"] = won["Deal Value (EUR)"].sum()

    kpi_df = pd.DataFrame([
        ["Website", "Sessions", kpi["Website Sessions"]],
        ["Website", "Ticket inquiries", kpi["Website Ticket Inquiries"]],
        ["Website", "CVR (weighted)", kpi["Website CVR (weighted)"]],
        ["Social", "Impressions", kpi["Social Impressions"]],
        ["Social", "Link clicks", kpi["Social Link Clicks"]],
        ["Social", "Engagement rate (weighted)", kpi["Social Engagement Rate (weighted)"]],
        ["Email", "Delivered", kpi["Email Delivered"]],
        ["Email", "Open rate (weighted)", kpi["Email Open Rate (weighted)"]],
        ["Email", "CTR (weighted)", kpi["Email CTR (weighted)"]],
        ["Email", "Ticket inquiries", kpi["Email Ticket Inquiries"]],
        ["Email", "Revenue attributed (EUR)", kpi["Email Revenue Attributed"]],
        ["Paid Ads", "Spend (EUR)", kpi["Ad Spend"]],
        ["Paid Ads", "Conversions", kpi["Ad Conversions"]],
        ["Paid Ads", "CPC (blended)", kpi["Ad CPC (blended)"]],
        ["Paid Ads", "CPA (blended)", kpi["Ad CPA (blended)"]],
        ["Sales", "Closed won deals", kpi["Deals Closed Won"]],
        ["Sales", "Closed lost deals", kpi["Deals Closed Lost"]],
        ["Sales", "Win rate (closed)", kpi["Win Rate (closed deals)"]],
        ["Sales", "Closed won revenue (EUR)", kpi["Revenue Closed Won (EUR)"]],
    ], columns=["Area", "Metric", "Value"])

    # Ads by platform (efficiency)
    ads_platform = ads.groupby("Platform").agg(
        spend=("Spend (EUR)", "sum"),
        clicks=("Clicks", "sum"),
        impressions=("Impressions", "sum"),
        conversions=("Conversions", "sum"),
    ).reset_index()
    ads_platform["CPC"] = ads_platform["spend"] / ads_platform["clicks"]
    ads_platform["CPA"] = ads_platform["spend"] / ads_platform["conversions"]
    ads_platform_sorted = ads_platform.sort_values("CPA")

    # Sales by lead source (conversion + revenue)
    lead_source_stats = sp.groupby("Lead Source Std").apply(
        lambda g: pd.Series({
            "leads": len(g),
            "closed_won": (g["Deal Stage"] == "Closed Won").sum(),
            "closed_lost": (g["Deal Stage"] == "Closed Lost").sum(),
            "lead_to_won_rate": (g["Deal Stage"] == "Closed Won").sum() / len(g),
            "revenue_won": g.loc[g["Deal Stage"] == "Closed Won", "Deal Value (EUR)"].sum()
        })
    ).reset_index()
    lead_source_stats_sorted = lead_source_stats.sort_values(["revenue_won", "leads"], ascending=[False, False])
    overall_lead_to_won = (sp["Deal Stage"] == "Closed Won").sum() / len(sp)

    # Forecast vs June targets (simple: pipeline * win rate)
    event_date = pd.Timestamp("2026-06-02")
    open_deals = sp[~sp["Deal Stage"].str.contains("Closed", case=False, na=False)].copy()
    open_by_june = open_deals[(open_deals["Expected Close Date"].notna()) & (open_deals["Expected Close Date"] <= event_date)]
    overall_win = kpi["Win Rate (closed deals)"]
    exp_wins_by_cat = open_by_june.groupby("Ticket Category").size() * overall_win

    # Current wins by category
    won_cat = won.groupby("Ticket Category").size()

    # CEO memo text (<=300 words)
    # Keep it compact; you can tweak text in the app UI too.
    google_cpa = float(ads_platform_sorted.iloc[0]["CPA"]) if len(ads_platform_sorted) else np.nan
    li_row = ads_platform_sorted[ads_platform_sorted["Platform"].str.contains("LinkedIn", case=False, na=False)]
    li_cpa = float(li_row["CPA"].iloc[0]) if len(li_row) else np.nan

    memo = f"""Subject: POT2026 performance snapshot + next actions (data through Jan 2026)

• Website: {kpi['Website Sessions']:,} sessions → {kpi['Website Ticket Inquiries']:,} inquiries (CVR {kpi['Website CVR (weighted)']*100:.2f}%).
• Social: {kpi['Social Impressions']:,} impressions, {int(kpi['Social Link Clicks']):,} clicks (engagement {kpi['Social Engagement Rate (weighted)']*100:.2f}%).
• Email: {kpi['Email Delivered']:,} delivered; open {kpi['Email Open Rate (weighted)']*100:.1f}%, CTR {kpi['Email CTR (weighted)']*100:.1f}%; €{kpi['Email Revenue Attributed']:,.0f} attributed.
• Paid ads: €{kpi['Ad Spend']:,.0f} spend → {int(kpi['Ad Conversions']):,} conversions (blended CPA €{kpi['Ad CPA (blended)']:.0f}); Google CPA ~€{google_cpa:,.0f} vs LinkedIn CPA ~€{li_cpa:,.0f}.
• Sales: {kpi['Deals Closed Won']} won / {kpi['Deals Closed Lost']} lost (win rate {kpi['Win Rate (closed deals)']*100:.1f}%), €{kpi['Revenue Closed Won (EUR)']:,.0f} won.

What to do next (Feb–Jun)
1) Shift budget away from inefficient LinkedIn prospecting into retargeting + sponsor lead-gen.
2) 24h SLA: every inquiry becomes a CRM lead tagged by source with nurture + sales follow-up.
3) Build a delegate conversion bridge (inquiry → nurture → booking → paid) and track weekly funnel drop-offs.
"""

    # Build a cleaned workbook for download (bytes)
    out_xlsx = io.BytesIO()
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        wt.to_excel(writer, index=False, sheet_name="clean_website_traffic")
        sm.to_excel(writer, index=False, sheet_name="clean_social_media")
        em.to_excel(writer, index=False, sheet_name="clean_email_campaigns")
        sp.to_excel(writer, index=False, sheet_name="clean_sales_pipeline")
        ads.to_excel(writer, index=False, sheet_name="clean_ad_spend")
        kpi_df.to_excel(writer, index=False, sheet_name="kpi_summary")
        lead_source_stats_sorted.to_excel(writer, index=False, sheet_name="sales_by_lead_source")
        ads_platform_sorted.to_excel(writer, index=False, sheet_name="ads_by_platform")
    out_xlsx.seek(0)

    return {
        "sheets": sheets,
        "wt": wt, "sm": sm, "em": em, "sp": sp, "ads": ads,
        "kpi": kpi, "kpi_df": kpi_df,
        "ads_platform": ads_platform_sorted,
        "lead_source_stats": lead_source_stats_sorted,
        "overall_lead_to_won": overall_lead_to_won,
        "open_by_june": open_by_june,
        "exp_wins_by_cat": exp_wins_by_cat,
        "won_cat": won_cat,
        "cleaned_workbook_bytes": out_xlsx.getvalue(),
        "memo": memo.strip()
    }

def fig_website_trends(wt: pd.DataFrame):
    wt_week = wt.sort_values("Week Starting")
    fig = plt.figure()
    plt.plot(wt_week["Week Starting"], wt_week["Sessions"], label="Sessions")
    plt.plot(wt_week["Week Starting"], wt_week["Ticket Inquiry Conversions"], label="Ticket inquiries")
    plt.title("Website traffic & inquiries (weekly)")
    plt.xlabel("Week starting")
    plt.ylabel("Count")
    plt.legend()
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    return fig

def fig_social_impressions(sm: pd.DataFrame):
    sm_week_plat = sm.groupby(["Week Starting", "Platform"]).agg(
        impressions=("Impressions", "sum")
    ).reset_index()
    fig = plt.figure()
    for plat, g in sm_week_plat.groupby("Platform"):
        plt.plot(g["Week Starting"], g["impressions"], label=f"{plat}")
    plt.title("Social impressions by platform (weekly)")
    plt.xlabel("Week starting")
    plt.ylabel("Impressions")
    plt.legend()
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    return fig

def fig_ads_efficiency(ads_platform: pd.DataFrame):
    fig = plt.figure()
    x = np.arange(len(ads_platform))
    plt.bar(x - 0.2, ads_platform["spend"], width=0.4, label="Spend (EUR)")
    plt.bar(x + 0.2, ads_platform["conversions"], width=0.4, label="Conversions")
    plt.xticks(x, ads_platform["Platform"], rotation=15, ha="right")
    plt.title("Paid ads: spend vs conversions")
    plt.legend()
    plt.tight_layout()
    return fig

def fig_sales_funnel(sp: pd.DataFrame):
    stage_order = ["Lead", "Contacted", "Qualified", "Negotiation", "Closed Won", "Closed Lost"]
    funnel = sp["Stage_N"].value_counts()
    funnel = funnel.reindex([s for s in stage_order if s in funnel.index]).fillna(0)
    fig = plt.figure()
    plt.bar(funnel.index, funnel.values)
    plt.title("Sales funnel (deal counts)")
    plt.xlabel("Stage")
    plt.ylabel("Deals")
    plt.xticks(rotation=20, ha="right")
    plt.tight_layout()
    return fig

def make_ai_note():
    return (
        "AI tools used: ChatGPT was used to accelerate data cleaning logic, KPI definitions, and draft the initial insight structure. "
        "All calculations were validated by recomputing metrics directly from the raw Excel sheets in Python. "
        "Visualizations and aggregations were generated programmatically to avoid manual errors. "
        "Final recommendations were based on the computed KPIs (e.g., CPA/CPC, lead→won rates, pipeline) rather than AI-generated assumptions."
    )

def build_insights_pdf_bytes(kpi, ads_platform, lead_source_stats, forecast_text, recommendations, hidden_insight):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    W, H = letter
    y = H - 72

    c.setFont("Helvetica-Bold", 14)
    c.drawString(72, y, "POT2026 — Insights Report (Tasks A/B)")
    y -= 18
    c.setFont("Helvetica", 10)
    c.drawString(72, y, "Period: Oct 2025 – Jan 2026 (source: provided case dataset)")
    y -= 22

    # KPI snapshot
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, y, "CEO snapshot")
    y -= 14
    c.setFont("Helvetica", 10)
    lines = [
        f"Website: {kpi['Website Sessions']:,} sessions; {kpi['Website Ticket Inquiries']:,} inquiries; CVR {kpi['Website CVR (weighted)']*100:.2f}%",
        f"Social: {kpi['Social Impressions']:,} impressions; {int(kpi['Social Link Clicks']):,} link clicks; engagement {kpi['Social Engagement Rate (weighted)']*100:.2f}%",
        f"Email: {kpi['Email Delivered']:,} delivered; open {kpi['Email Open Rate (weighted)']*100:.1f}%; CTR {kpi['Email CTR (weighted)']*100:.1f}%; attributed €{kpi['Email Revenue Attributed']:,.0f}",
        f"Ads: spend €{kpi['Ad Spend']:,.0f}; conversions {int(kpi['Ad Conversions']):,}; blended CPA €{kpi['Ad CPA (blended)']:.2f}",
        f"Sales: won {kpi['Deals Closed Won']}; lost {kpi['Deals Closed Lost']}; win rate {kpi['Win Rate (closed deals)']*100:.1f}%; won revenue €{kpi['Revenue Closed Won (EUR)']:,.0f}",
    ]
    for ln in lines:
        c.drawString(80, y, u"\u2022 " + ln)
        y -= 12

    y -= 8
    # ROI
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, y, "1) Best & worst ROI (proxy via CPA)")
    y -= 14
    c.setFont("Helvetica", 10)
    best = ads_platform.iloc[0]
    worst = ads_platform.iloc[-1]
    c.drawString(80, y, f"- Best: {best['Platform']} (CPA ≈ €{best['CPA']:.2f})")
    y -= 12
    c.drawString(80, y, f"- Worst: {worst['Platform']} (CPA ≈ €{worst['CPA']:.2f})")
    y -= 18

    # Conversion rates
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, y, "2) Conversion rates (lead → closed won)")
    y -= 14
    c.setFont("Helvetica", 10)
    overall_lead_to_won = (lead_source_stats["closed_won"].sum() / lead_source_stats["leads"].sum())
    c.drawString(80, y, f"- Overall lead→won: {overall_lead_to_won*100:.1f}%")
    y -= 14
    c.drawString(80, y, "- Variation by lead source (top shown in dashboard table).")
    y -= 18

    # Forecast
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, y, "3) Forecast vs June targets (300 delegates / 25 sponsors)")
    y -= 14
    c.setFont("Helvetica", 10)
    for part in forecast_text.split("\n"):
        c.drawString(80, y, part[:120])
        y -= 12
        if y < 80:
            c.showPage()
            y = H - 72
            c.setFont("Helvetica", 10)

    y -= 10

    # Hidden insight
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, y, "4) Hidden insight")
    y -= 14
    c.setFont("Helvetica", 10)
    c.drawString(80, y, hidden_insight[:120])
    y -= 18

    # Recommendations
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, y, "5) Recommendations (next 4 months)")
    y -= 14
    c.setFont("Helvetica", 10)
    for r in recommendations:
        c.drawString(80, y, u"\u2022 " + r[:120])
        y -= 12
        if y < 80:
            c.showPage()
            y = H - 72
            c.setFont("Helvetica", 10)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()


# -----------------------------
# UI
# -----------------------------
st.title("POT2026 Case Study — Dashboard")

st.markdown("""
Upload the **Excel dataset**.
""")

uploaded = st.file_uploader("Upload POT2026_Raw_Data_Case_Study.xlsx", type=["xlsx"])

if not uploaded:
    st.stop()

data = clean_and_model(uploaded.getvalue())

st.success(f"Loaded sheets: {', '.join(data['sheets'])}")

# -----------------------------
# Task A: Dashboard
# -----------------------------
st.header("KPI Dashboard")

kpi = data["kpi"]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Website sessions", f"{kpi['Website Sessions']:,}")
c1.metric("Website inquiries", f"{kpi['Website Ticket Inquiries']:,}")
c2.metric("Website CVR", f"{kpi['Website CVR (weighted)']*100:.2f}%")
c2.metric("Social impressions", f"{kpi['Social Impressions']:,}")
c3.metric("Email delivered", f"{kpi['Email Delivered']:,}")
c3.metric("Email open rate", f"{kpi['Email Open Rate (weighted)']*100:.1f}%")
c4.metric("Ad spend", f"€{kpi['Ad Spend']:,.0f}")
c4.metric("Ad CPA (blended)", f"€{kpi['Ad CPA (blended)']:,.2f}")

st.subheader("KPI table")
st.dataframe(data["kpi_df"], use_container_width=True)

st.subheader("Charts")
ch1, ch2 = st.columns(2)
with ch1:
    st.pyplot(fig_website_trends(data["wt"]))
with ch2:
    st.pyplot(fig_social_impressions(data["sm"]))

ch3, ch4 = st.columns(2)
with ch3:
    st.pyplot(fig_ads_efficiency(data["ads_platform"]))
with ch4:
    st.pyplot(fig_sales_funnel(data["sp"]))

st.subheader("Ads efficiency table (ROI proxy)")
st.dataframe(data["ads_platform"], use_container_width=True)

st.subheader("Sales conversion by lead source")
st.dataframe(data["lead_source_stats"], use_container_width=True)

# -----------------------------
# Task B: Insights + recommendations
# -----------------------------
st.header("Insights, ROI, Forecast vs June targets")

ads_platform = data["ads_platform"]
best = ads_platform.iloc[0] if len(ads_platform) else None
worst = ads_platform.iloc[-1] if len(ads_platform) else None

st.subheader("Best vs worst paid efficiency (CPA)")
if best is not None and worst is not None:
    st.write(f"- **Best:** {best['Platform']} (CPA ≈ €{best['CPA']:.2f})")
    st.write(f"- **Worst:** {worst['Platform']} (CPA ≈ €{worst['CPA']:.2f})")

st.subheader("Conversion rates")
st.write(f"- Overall **lead → won**: **{data['overall_lead_to_won']*100:.1f}%** (based on CRM deals).")
top_sources = data["lead_source_stats"].head(5)[["Lead Source Std", "leads", "closed_won", "lead_to_won_rate", "revenue_won"]]
st.write("Top sources (by won revenue / volume):")
st.dataframe(top_sources, use_container_width=True)

st.subheader("Forecast vs June targets (300 delegates / 25 sponsors)")
won_cat = data["won_cat"]
existing_delegate = int(won_cat.get("Delegate", 0))
existing_sponsor = int(won_cat.get("Sponsor", 0))
exp = data["exp_wins_by_cat"]
exp_delegate = float(exp.get("Delegate", 0))
exp_sponsor = float(exp.get("Sponsor", 0))

st.write(f"- Closed won so far: **{existing_delegate} delegates**, **{existing_sponsor} sponsors**.")
st.write(f"- Open pipeline expected close by June (count):")
st.dataframe(data["open_by_june"].groupby("Ticket Category").size().rename("deals").reset_index(), use_container_width=True)
st.write(f"- Using historical win rate (**{kpi['Win Rate (closed deals)']*100:.1f}%**) → expected additional wins:")
st.write(f"  - Delegates: **~{exp_delegate:.1f}**")
st.write(f"  - Sponsors: **~{exp_sponsor:.1f}**")
st.info("Caveat: website/email/ads 'ticket inquiries' are earlier-funnel signals; CRM closed-won may lag demand if inquiries aren’t flowing into sales consistently.")

st.subheader("Hidden insight")
st.write("- LinkedIn **paid** is inefficient (high CPA), but LinkedIn **outreach/organic leads** are the largest closed-won revenue driver — indicating a targeting/offer mismatch in paid, not a channel mismatch.")

st.subheader("Recommended next steps (Feb–Jun)")
st.write("1) Move budget from broad LinkedIn prospecting into retargeting + sponsor lead-gen forms; keep cold targeting tightly qualified.")
st.write("2) Enforce a 24h SLA: every inquiry becomes a CRM lead with source tags + automated nurture + sales follow-up.")
st.write("3) Build a delegate conversion bridge: inquiry → nurture → booking → paid; review funnel drop-offs weekly.")

# -----------------------------
# Task C: CEO memo
# -----------------------------
st.header("CEO Memo")

memo_text = st.text_area("Memo (editable)", value=data["memo"], height=260)
memo_bytes = memo_text.encode("utf-8")

st.download_button(
    "Download CEO memo (.txt)",
    data=memo_bytes,
    file_name="CEO_Memo_POT2026.txt",
    mime="text/plain"
)

# -----------------------------
# Downloads: cleaned workbook
# -----------------------------
st.header("Downloads")

st.download_button(
    "Download cleaned + analysis workbook (.xlsx)",
    data=data["cleaned_workbook_bytes"],
    file_name="POT2026_Cleaned_Analysis_Output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)



# --- AI tools note
ai_note = make_ai_note()
st.subheader("AI Tools Notes")
st.write(ai_note)
st.download_button(
    "Download AI tools note (.txt)",
    data=ai_note.encode("utf-8"),
    file_name="AI_Tools_Note.txt",
    mime="text/plain"
)

# --- Weekly reporting template (Bonus)
st.subheader("Weekly reporting template (proposal)")
weekly_template = """Weekly Reporting Template (minimal manual effort)

Automated (Python/Streamlit):
- Pull latest data from source files (or connected systems) and refresh KPIs:
  Website: sessions, inquiries, CVR, bounce, top landing pages
  Social: impressions, engagement rate, link clicks by platform
  Email: delivered, open rate, CTR, inquiries, revenue attributed
  Ads: spend, clicks, conversions, CPC, CPA by platform/campaign
  Sales: funnel counts, lead→won rate by source, pipeline due by event date
- Generate and export: KPI workbook + PDF report snapshot + memo draft.

Manual (light-touch):
- Add context for anomalies (e.g., campaign launch, press mention, site outage).
- Confirm pipeline stage notes for top deals and adjust close-date assumptions if needed.
- Approve final recommendations and send to leadership.

Cadence:
- Monday morning refresh with last full week’s data + month-to-date rollup.
"""
st.text_area("Weekly template (editable)", value=weekly_template, height=220)
st.download_button(
    "Download weekly template (.txt)",
    data=weekly_template.encode("utf-8"),
    file_name="Weekly_Reporting_Template.txt",
    mime="text/plain"
)

# --- PDF Insights Report (Task B submission format)
st.subheader("Insights Report (PDF download)")

forecast_text = (
    f"Closed won so far: Delegates={int(existing_delegate)}, Sponsors={int(existing_sponsor)}\n"
    f"Open pipeline due by June (count): see dashboard table\n"
    f"Using win rate {kpi['Win Rate (closed deals)']*100:.1f}% → expected additional wins: "
    f"Delegates~{exp_delegate:.1f}, Sponsors~{exp_sponsor:.1f}\n"
    "Conclusion: delegate target not on track based on CRM closed-won alone; sponsor target requires pipeline expansion and/or higher win rate.\n"
    "Caveat: inquiries are early-funnel demand; ensure inquiries flow into CRM and are tracked to paid sales."
)

hidden_insight = "LinkedIn paid is high-CPA, but LinkedIn outreach/organic leads drive the highest closed-won revenue—fix targeting/offer in paid rather than abandoning the channel."

recommendations = [
    "Shift 30–50% of LinkedIn Ads budget to retargeting + sponsor lead-gen forms; reduce broad cold prospecting.",
    "Enforce a 24h SLA: every inquiry becomes a CRM lead with source tags + automated nurture + sales follow-up.",
    "Build a delegate conversion bridge (inquiry → nurture → booking → paid) and review funnel drop-offs weekly."
]

pdf_bytes = build_insights_pdf_bytes(
    kpi=kpi,
    ads_platform=data["ads_platform"],
    lead_source_stats=data["lead_source_stats"],
    forecast_text=forecast_text,
    recommendations=recommendations,
    hidden_insight=hidden_insight
)

st.download_button(
    "Download Insights Report (PDF)",
    data=pdf_bytes,
    file_name="POT2026_Insights_Report.pdf",
    mime="application/pdf"
)