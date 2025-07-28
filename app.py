import streamlit as st
from PIL import Image
import pandas as pd
import re
from fuzzywuzzy import fuzz
import difflib
import os

# --- APP CONFIGURATION ---
st.set_page_config(page_title="HMV Fair Quote Tool", layout="wide", page_icon="🔧")

## --- HEADER / LOGOS ---
left, center, right = st.columns([2, 6, 2])
with left:
    st.image("logo1.png", width=140)
with right:
    st.image("logo2.png", width=140)
with center:
    st.markdown("""
    <h1 style='text-align:center; margin-bottom:0;'>HMV Fair Quote Validation Tool</h1>
    <h5 style='text-align:center; color:#3498db; margin-top:0;'>Making Aircraft Maintenance Quotes Objective</h5>
    <hr style='border:2px solid #3498db; border-radius:5px; margin-bottom:24px;'>
    """, unsafe_allow_html=True)

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    body { background: #fcfcfc; }
    .conclusion-box {
        padding: 1.6em 2em; border-radius: 14px; margin: 2em 0 1.3em 0;
        border-left: 7px solid; box-shadow: 0 4px 16px rgba(0,0,0,0.07);
        background: linear-gradient(90deg,#fafdff,#f3f8fc 70%,#fff 100%);
        transition: 0.2s; font-size:1.12rem;
    }
    .exact-match { border-color: #38c976; }
    .approx-match { border-color: #ffb338; }
    .closest-match { border-color: #ed5565; }
    .conclusion-header {font-size:1.3rem; margin-bottom:5px;}
    .metric-card {
        background: #fff; border-radius: 10px; padding: 2em 1em;
        box-shadow: 0 2px 8px rgba(52, 152, 219,0.06);
        display:flex; flex-direction:column; align-items:center;
    }
    .metric-label { color:#7f8c8d; font-size:1.1rem;}
    .metric-value { font-size: 2.3rem; font-weight:700; color:#1c2833; }
    .diff-positive { color: #e74c3c;}
    .diff-negative { color: #2ecc71;}
    .diff-neutral { color: #3498db;}
    .result-table th {background: #3498db; color: #fff; font-weight:600; padding:11px 12px;}
    .result-table td {background:#fafcff; padding:10px 12px; }
    .result-table tr {border-bottom:1px solid #dde9f7;}
    </style>
""", unsafe_allow_html=True)

# --- LOAD HISTORICAL DATA ---
DATA_PATH = 'hmv_data.xlsx'
LOGO1 = "logo1.png"; LOGO2 = "logo2.png"

try:
    df = pd.read_excel(DATA_PATH)
except Exception:
    st.error(f"Historical data not found! Please ensure 'hmv_data.xlsx' is available.")
    st.stop()

# --- DATA PREPROCESSING ---
def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = re.sub(r'\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b', '', text)  # Remove dates
    text = re.sub(r'\s+', ' ', text).strip()
    return text

df['Normalized Corrective Action'] = df['Corrective Action'].apply(normalize_text)
df['Normalized Discrepancy'] = df['Description'].apply(lambda x: normalize_text(str(x).replace("(FOR REFERENCE ONLY)", "")))
df['Combined Key'] = df['Normalized Discrepancy'] + " | " + df['Normalized Corrective Action']

# Clustering for match groups (fuzzy grouped key)
clusters = {}
for key in df['Combined Key'].unique():
    if not key:
        continue
    for rep in clusters:
        if fuzz.token_set_ratio(key, rep) >= 90:
            clusters[rep].append(key)
            break
    else:
        clusters[key] = [key]

key_to_rep = {k: r for r, lst in clusters.items() for k in lst}
df['Cluster Key'] = df['Combined Key'].map(key_to_rep)

# Group to calculate historic/fair quote averages
hours = df.groupby('Cluster Key')['Total Hours'].agg(['mean','count']).reset_index()
hours.columns = ['Cluster Key', 'Actual Historic Hours', 'Occurrences']
df = df.merge(hours, on='Cluster Key', how='left')
df['Fair Quote (hrs)'] = df['Actual Historic Hours'].round(2)


# --- USER INPUT: QUOTE ENTRY FORM ---
st.markdown("""
<div class="form-container" style="background:#eaf3fa;padding:1.7em 2.3em; border-radius:14px; box-shadow:0 2px 12px rgba(52,152,219,0.12);">
<h4 style='color:#1e4669; margin-bottom:12px;'>Enter Maintenance Quote Details</h4>
""", unsafe_allow_html=True)

with st.form("quote_form", clear_on_submit=False):
    col1, col2 = st.columns(2)
    with col1:
        discrepancy_input = st.text_area("Description of Non-Routine", height=90,
                                         placeholder="Describe the issue or discrepancy...")
    with col2:
        corrective_input = st.text_area("Corrective Action", height=90,
                                        placeholder="Describe the corrective action taken...")
    supplier_hours = st.number_input("Supplier Quoted Hours", min_value=0.0, step=0.1)
    submit = st.form_submit_button("🔍 Analyze Quote", use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)

# --- ANALYSIS ---
def semantic_overlap(a, b):
    if not a or not b:
        return 0
    matcher = difflib.SequenceMatcher(None, a.split(), b.split())
    return matcher.ratio() * 100

def get_decision_conclusion(supplier, fair):
    if fair == 0 or pd.isna(fair):
        percent_diff = "N/A"
        diff_class = "diff-neutral"
    else:
        percent_diff = ((supplier - fair) / fair) * 100
        if percent_diff < 0:
            diff_class = "diff-negative"
        elif abs(percent_diff) <= 5:
            diff_class = "diff-neutral"
        else:
            diff_class = "diff-positive"
    # Format display
    if fair == 0 or pd.isna(fair):
        percent_display = "N/A (no historical data)"
    else:
        sign = "+" if percent_diff >= 0 else ""
        percent_display = f"{sign}{percent_diff:.1f}%"
    # Generate conclusion text
    if fair == 0 or pd.isna(fair):
        return ("No historical data available — manual review recommended.", "#b0202e", percent_display, diff_class)
    if supplier < fair:
        return ("FAIR QUOTE: Supplier below historic average. Consider approving.", "#22aa58", percent_display, diff_class)
    elif abs(supplier - fair) / fair <= 0.05:
        return ("IN EXPECTED RANGE (±5%). Consider approving.", "#222f3e", percent_display, diff_class)
    else:
        return ("HIGHER THAN HISTORIC — Needs BP review.", "#b0202e", percent_display, diff_class)

def highlight_diff(text, ref):
    ref_words = set(ref.split())
    return " ".join([f"<b><span style='color:#e67e22'>{w}</span></b>" if w not in ref_words else w for w in text.split()])

# --- DECISION LOGIC ---
if submit and discrepancy_input and corrective_input:
    norm_disc = normalize_text(discrepancy_input.replace("(FOR REFERENCE ONLY)", ""))
    norm_corr = normalize_text(corrective_input)
    combined_input = norm_disc + " | " + norm_corr

    exact = df[df['Combined Key'] == combined_input]

    def total_similarity(row):
        d_ov = semantic_overlap(norm_disc, row['Normalized Discrepancy'])
        c_ov = semantic_overlap(norm_corr, row['Normalized Corrective Action'])
        return (d_ov + c_ov) / 2

    df['Overlap'] = df.apply(total_similarity, axis=1)
    approx = df[(df['Overlap'] >= 55) & (df['Combined Key'] != combined_input)]
    top2 = approx.sort_values(by='Overlap', ascending=False).head(2)
    closest = df[df['Overlap'] < 55].sort_values(by='Overlap', ascending=False).head(1)

    # --- CONCLUSION + METRICS UI BLOCK ---
    st.markdown("<hr>", unsafe_allow_html=True)
    if not exact.empty:
        row = exact.iloc[0]
        conclusion, color, percent_diff, diff_class = get_decision_conclusion(supplier_hours, row['Fair Quote (hrs)'])

        st.markdown(f"""
        <div class="conclusion-box exact-match">
        <div class="conclusion-header" style="color:{color};">Conclusion: {conclusion}</div>
        <span style="font-weight:500;">Match Type:</span> Exact Match
        <div style="float:right;"><b>Occurrences:</b> {row['Occurrences']}<br>
        <b>Historical Avg Hours:</b> {row['Actual Historic Hours']:.2f}</div>
        </div>
        """, unsafe_allow_html=True)

        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>Historic (Fair) Hours</span><div class='metric-value'>{row['Fair Quote (hrs)']:.2f}</div></div>",unsafe_allow_html=True)
        with m2:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>Supplier Quoted Hours</span><div class='metric-value'>{supplier_hours:.2f}</div></div>",unsafe_allow_html=True)
        with m3:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>% Difference</span><div class='metric-value {diff_class}'>{percent_diff}</div></div>",unsafe_allow_html=True)
        st.success("✅ <b>Exact historic match found</b>", unsafe_allow_html=True)
        st.dataframe(exact[['Description', 'Corrective Action', 'Actual Historic Hours', 'Fair Quote (hrs)', 'Occurrences']])
    elif not top2.empty:
        row = top2.iloc[0]
        conclusion, color, percent_diff, diff_class = get_decision_conclusion(supplier_hours, row['Fair Quote (hrs)'])
        st.markdown(f"""
        <div class="conclusion-box approx-match">
        <div class="conclusion-header" style="color:{color};">Conclusion: {conclusion}</div>
        <span style="font-weight:500;">Match Type:</span> Approximate Match
        <div style="float:right;"><b>Occurrences:</b> {row['Occurrences']}<br>
        <b>Historical Avg Hours:</b> {row['Actual Historic Hours']:.2f}</div>
        </div>
        """, unsafe_allow_html=True)
        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>Historic (Fair) Hours</span><div class='metric-value'>{row['Fair Quote (hrs)']:.2f}</div></div>",unsafe_allow_html=True)
        with m2:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>Supplier Quoted Hours</span><div class='metric-value'>{supplier_hours:.2f}</div></div>",unsafe_allow_html=True)
        with m3:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>% Difference</span><div class='metric-value {diff_class}'>{percent_diff}</div></div>",unsafe_allow_html=True)
        st.info("🔍 Showing top approximate matches (≥55% sim.)", icon="🔎")
        rows = []
        for _, row in top2.iterrows():
            rows.append({
                'Description': highlight_diff(row['Normalized Discrepancy'], norm_disc),
                'Corrective Action': highlight_diff(row['Normalized Corrective Action'], norm_corr),
                'Historic Hours': f"{row['Actual Historic Hours']:.2f}",
                'Fair Quote (hrs)': f"{row['Fair Quote (hrs)']:.2f}",
                'Occurrences': row['Occurrences'],
                'Overlap %': f"{row['Overlap']:.1f}%"
            })
        # Approximate result HTML table
        html_table = """
        <table style="width:100%;" class="result-table">
        <tr><th>Description</th><th>Corrective</th><th>Historic Hours</th><th>Fair Quote (hrs)</th><th>Occurrences</th><th>Overlap %</th></tr>
        """
        for row in rows:
            html_table += f"<tr><td>{row['Description']}</td><td>{row['Corrective Action']}</td><td>{row['Historic Hours']}</td><td>{row['Fair Quote (hrs)']}</td><td>{row['Occurrences']}</td><td>{row['Overlap %']}</td></tr>"
        html_table += "</table>"
        st.markdown(html_table, unsafe_allow_html=True)
    elif not closest.empty:
        row = closest.iloc[0]
        conclusion, color, percent_diff, diff_class = get_decision_conclusion(supplier_hours, row['Fair Quote (hrs)'])
        st.markdown(f"""
        <div class="conclusion-box closest-match">
        <div class="conclusion-header" style="color:{color};">Conclusion: {conclusion}</div>
        <span style="font-weight:500;">Match Type:</span> Nearest Reference (low similarity)
        <div style="float:right;"><b>Occurrences:</b> {row['Occurrences']}<br>
        <b>Historic Avg Hours:</b> {row['Actual Historic Hours']:.2f}</div>
        </div>
        """, unsafe_allow_html=True)
        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>Historic (Fair) Hours</span><div class='metric-value'>{row['Fair Quote (hrs)']:.2f}</div></div>",unsafe_allow_html=True)
        with m2:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>Supplier Quoted Hours</span><div class='metric-value'>{supplier_hours:.2f}</div></div>",unsafe_allow_html=True)
        with m3:
            st.markdown(f"<div class='metric-card'><span class='metric-label'>% Difference</span><div class='metric-value {diff_class}'>{percent_diff}</div></div>",unsafe_allow_html=True)
        st.warning("📝 No close matches found — showing only nearest reference.", icon="⚠️")
        st.markdown(f"""
        <table class="result-table" style="width:100%;">
        <tr><th>Description</th><th>Corrective Action</th><th>Historic Hours</th><th>Fair Quote (hrs)</th><th>Occurrences</th><th>Overlap %</th></tr>
        <tr>
        <td>{row['Normalized Discrepancy']}</td>
        <td>{row['Normalized Corrective Action']}</td>
        <td>{row['Actual Historic Hours']:.2f}</td>
        <td>{row['Fair Quote (hrs)']:.2f}</td>
        <td>{row['Occurrences']}</td>
        <td>{row['Overlap']:.1f}%</td>
        </tr></table>
        """, unsafe_allow_html=True)
        st.info("""
        ❓ <b>No reliable or similar past instance was found for this combination.<br>
        If this is a valid quote, you can contribute it for future reference:</b>
        """, icon="🗂", unsafe_allow_html=True)
        if st.button("➕ Add this Quote as a New Historical Instance", type="primary"):
            # Append to dataset + write to excel (simulate DB)
            new_row = {
                'Description': discrepancy_input,
                'Corrective Action': corrective_input,
                'Total Hours': supplier_hours,
                'Year': pd.to_datetime("today").year
            }
            df_new = pd.read_excel(DATA_PATH)
            df_new = df_new.append(new_row, ignore_index=True)
            try:
                df_new.to_excel(DATA_PATH, index=False)
                st.success("Added! This new quote will be used for future analysis.")
            except Exception as e:
                st.error("Failed to save new instance. Check file permissions.")
        st.markdown("""
        <span style='color:#5776bf;'>
        <i>
        Use this button if you want to save this unique case for future reference.<br>
        This is useful when the quote describes a never-seen issue type or rare scenario.
        </i>
        </span>
        """, unsafe_allow_html=True)
else:
    st.info("ℹ️ Enter a quote description, corrective action, and hours to begin.")
