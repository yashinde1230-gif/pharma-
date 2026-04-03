import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import io

st.set_page_config(page_title="PharmaExcelIQ", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.stApp { background: linear-gradient(135deg, #0a0e1a 0%, #0d1b2a 50%, #0a0e1a 100%); color: #e0e6f0; }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #0d1b2a 0%, #112240 100%); border-right: 1px solid #1e3a5f; }
[data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { color: #a8c0d6 !important; }
.main-header { font-size: 2.8rem; font-weight: 800; text-align: center; letter-spacing: 2px; padding: 1rem 0 0.2rem 0; background: linear-gradient(90deg, #1565C0, #00BCD4, #1565C0); background-size: 200% auto; -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
.sub-header { font-size: 0.95rem; color: #64b5f6; text-align: center; margin-bottom: 0.5rem; letter-spacing: 1px; }
.upload-box { background: #0d2137; border: 2px dashed #1e3a5f; border-radius: 16px; padding: 2rem; text-align: center; margin: 1rem 0; }
.upload-title { color: #64b5f6; font-size: 1.2rem; font-weight: 700; margin-bottom: 0.5rem; }
.upload-sub { color: #3d5a73; font-size: 0.85rem; }
.metric-card { background: linear-gradient(135deg, #0d2137 0%, #112240 100%); border: 1px solid #1e3a5f; border-radius: 16px; padding: 1.2rem 1rem; text-align: center; position: relative; overflow: hidden; }
.metric-card::before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px; border-radius: 16px 16px 0 0; }
.metric-card-blue::before { background: #2196F3; }
.metric-card-green::before { background: #00c853; }
.metric-card-orange::before { background: #FF9800; }
.metric-card-purple::before { background: #9C27B0; }
.metric-number { font-size: 2rem; font-weight: 800; margin-bottom: 0.3rem; }
.metric-number-blue { color: #64b5f6; }
.metric-number-green { color: #69f0ae; }
.metric-number-orange { color: #ffcc02; }
.metric-number-purple { color: #ce93d8; }
.metric-label { font-size: 0.75rem; color: #7a9abf; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; }
.section-header { color: #64b5f6; font-size: 1rem; font-weight: 700; text-transform: uppercase; letter-spacing: 2px; border-left: 3px solid #2196F3; padding-left: 12px; margin: 1.5rem 0 1rem 0; }
.col-badge { display: inline-block; background: #0d2137; border: 1px solid #1e3a5f; border-radius: 20px; padding: 3px 12px; margin: 3px; font-size: 0.75rem; color: #64b5f6; }
.insight-card { background: #0d2137; border: 1px solid #1e3a5f; border-radius: 12px; padding: 1rem; margin-bottom: 0.5rem; }
.insight-title { color: #64b5f6; font-size: 0.8rem; font-weight: 700; letter-spacing: 1px; margin-bottom: 0.3rem; }
.insight-value { color: #e0e6f0; font-size: 1rem; font-weight: 600; }
.footer { text-align: center; color: #3d5a73; font-size: 0.78rem; padding: 1rem 0; letter-spacing: 1px; }
hr { border-color: #1e3a5f !important; }
</style>
""", unsafe_allow_html=True)

DARK_BG    = "#0d2137"
GRID_COLOR = "#1e3a5f"
TEXT_COLOR = "#a8c0d6"

def dark_chart_style(ax, fig):
    fig.patch.set_facecolor(DARK_BG)
    ax.set_facecolor(DARK_BG)
    ax.xaxis.label.set_color(TEXT_COLOR)
    ax.yaxis.label.set_color(TEXT_COLOR)
    ax.tick_params(colors=TEXT_COLOR)
    ax.grid(True, color=GRID_COLOR, linewidth=0.5, alpha=0.7)
    for spine in ax.spines.values():
        spine.set_edgecolor(GRID_COLOR)

def format_number(n):
    if n >= 10000000:
        return str(round(n/10000000, 1)) + " Cr"
    elif n >= 100000:
        return str(round(n/100000, 1)) + " L"
    elif n >= 1000:
        return str(round(n/1000, 1)) + "K"
    else:
        return str(round(n))

st.markdown('<div class="main-header">PHARMAEXCELIQ</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">INTELLIGENT EXCEL DASHBOARD FOR PHARMA DATA ANALYSIS</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

st.sidebar.markdown('<div style="text-align:center; padding:1rem 0;"><div style="font-size:2rem;">📊</div><div style="color:#64b5f6; font-weight:800; font-size:1.1rem; letter-spacing:2px;">PHARMAEXCELIQ</div><div style="color:#3d5a73; font-size:0.7rem; letter-spacing:1px; margin-top:4px;">EXCEL INTELLIGENCE</div></div>', unsafe_allow_html=True)
st.sidebar.markdown("---")
st.sidebar.markdown("<div style='color:#64b5f6; font-size:0.8rem; font-weight:700; letter-spacing:2px; margin-bottom:0.8rem;'>UPLOAD YOUR DATA</div>", unsafe_allow_html=True)

uploaded_file = st.sidebar.file_uploader(
    "Choose Excel file",
    type=["xlsx", "xls", "csv"],
    help="Upload any Excel or CSV file with pharma data"
)

if uploaded_file is None:
    st.markdown("""
    <div class="upload-box">
        <div class="upload-title">📂 Upload Your Excel File</div>
        <div class="upload-sub">Supports .xlsx · .xls · .csv</div>
        <br>
        <div style="color:#3d5a73; font-size:0.85rem;">
            Upload any pharma data file from the sidebar<br>
            Sales data · MR reports · Trial data · Market share · Any Excel
        </div>
        <br>
        <div style="color:#1e3a5f; font-size:0.75rem;">
            Your data never leaves your browser · 100% private
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div class='section-header'>WHAT THIS TOOL ANALYZES</div>", unsafe_allow_html=True)

    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown('<div class="insight-card"><div class="insight-title">SALES PERFORMANCE</div><div class="insight-value">Target vs achievement · MR rankings · Territory comparison</div></div>', unsafe_allow_html=True)
    with f2:
        st.markdown('<div class="insight-card"><div class="insight-title">PRODUCT ANALYTICS</div><div class="insight-value">Drug-wise sales · Therapy area split · Prescription trends</div></div>', unsafe_allow_html=True)
    with f3:
        st.markdown('<div class="insight-card"><div class="insight-title">SMART INSIGHTS</div><div class="insight-value">Auto-detects columns · Generates charts instantly · Export ready</div></div>', unsafe_allow_html=True)

else:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.sidebar.markdown("---")
        st.sidebar.markdown('<div style="background:#0d2137; border:1px solid #00c853; border-radius:10px; padding:1rem; text-align:center;"><div style="color:#00c853; font-size:0.8rem; font-weight:700;">FILE LOADED</div><div style="color:#69f0ae; font-size:1.2rem; font-weight:800; margin-top:4px;">' + uploaded_file.name + '</div><div style="color:#3d5a73; font-size:0.75rem; margin-top:4px;">' + str(len(df)) + ' rows · ' + str(len(df.columns)) + ' columns</div></div>', unsafe_allow_html=True)

        numeric_cols     = df.select_dtypes(include="number").columns.tolist()
        categorical_cols = df.select_dtypes(include="object").columns.tolist()
        date_cols        = [c for c in df.columns if any(word in c.lower() for word in ["date","month","year","period"])]

        st.sidebar.markdown("---")
        st.sidebar.markdown("<div style='color:#64b5f6; font-size:0.8rem; font-weight:700; letter-spacing:2px;'>FILTERS</div>", unsafe_allow_html=True)

        active_filters = {}
        for col in categorical_cols[:3]:
            unique_vals = sorted(df[col].dropna().unique().tolist())
            if len(unique_vals) <= 30:
                selected = st.sidebar.multiselect(col, options=unique_vals, default=unique_vals)
                active_filters[col] = selected

        df_filtered = df.copy()
        for col, selected in active_filters.items():
            df_filtered = df_filtered[df_filtered[col].isin(selected)]

        st.markdown("<div class='section-header'>DATASET OVERVIEW</div>", unsafe_allow_html=True)

        cols_detected = ""
        for col in df.columns:
            cols_detected += '<span class="col-badge">' + col + '</span>'
        st.markdown(cols_detected, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        o1, o2, o3, o4 = st.columns(4)
        with o1:
            st.markdown('<div class="metric-card metric-card-blue"><div class="metric-number metric-number-blue">' + str(len(df_filtered)) + '</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
        with o2:
            st.markdown('<div class="metric-card metric-card-green"><div class="metric-number metric-number-green">' + str(len(df.columns)) + '</div><div class="metric-label">Columns</div></div>', unsafe_allow_html=True)
        with o3:
            st.markdown('<div class="metric-card metric-card-orange"><div class="metric-number metric-number-orange">' + str(len(numeric_cols)) + '</div><div class="metric-label">Numeric Cols</div></div>', unsafe_allow_html=True)
        with o4:
            st.markdown('<div class="metric-card metric-card-purple"><div class="metric-number metric-number-purple">' + str(len(categorical_cols)) + '</div><div class="metric-label">Category Cols</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if numeric_cols:
            st.markdown("<div class='section-header'>NUMERIC SUMMARY</div>", unsafe_allow_html=True)
            num_cols_to_show = numeric_cols[:4]
            metric_cols = st.columns(len(num_cols_to_show))
            colors = ["blue","green","orange","purple"]
            for i, col in enumerate(num_cols_to_show):
                total_val = df_filtered[col].sum()
                avg_val   = df_filtered[col].mean()
                with metric_cols[i]:
                    st.markdown(
                        '<div class="metric-card metric-card-' + colors[i % 4] + '">'
                        '<div class="metric-number metric-number-' + colors[i % 4] + '">' + format_number(total_val) + '</div>'
                        '<div class="metric-label">Total ' + col[:15] + '</div>'
                        '<div style="color:#3d5a73; font-size:0.75rem; margin-top:4px;">Avg: ' + format_number(avg_val) + '</div>'
                        '</div>',
                        unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

        if categorical_cols and numeric_cols:
            st.markdown("<div class='section-header'>SMART CHARTS</div>", unsafe_allow_html=True)

            st.markdown("<div style='color:#a8c0d6; font-size:0.85rem; margin-bottom:1rem;'>Choose what to analyze:</div>", unsafe_allow_html=True)

            ctrl1, ctrl2, ctrl3 = st.columns(3)
            with ctrl1:
                x_axis = st.selectbox("Category (X axis)", options=categorical_cols)
            with ctrl2:
                y_axis = st.selectbox("Value (Y axis)", options=numeric_cols)
            with ctrl3:
                chart_type = st.selectbox("Chart Type", options=["Bar Chart","Horizontal Bar","Line Chart","Pie Chart"])

            chart_data = df_filtered.groupby(x_axis)[y_axis].sum().sort_values(ascending=False).head(15)

            fig, ax = plt.subplots(figsize=(12, 5))
            dark_chart_style(ax, fig)

            if chart_type == "Bar Chart":
                colors_list = [plt.cm.Blues(0.4 + 0.6 * i / len(chart_data)) for i in range(len(chart_data))]
                bars = ax.bar(chart_data.index, chart_data.values, color=colors_list, edgecolor=DARK_BG, linewidth=1)
                for bar, val in zip(bars, chart_data.values):
                    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + chart_data.max()*0.01, format_number(val), ha="center", va="bottom", fontsize=8, fontweight="bold", color=TEXT_COLOR)
                plt.xticks(rotation=30, ha="right", fontsize=8)

            elif chart_type == "Horizontal Bar":
                colors_list = [plt.cm.Blues(0.4 + 0.6 * i / len(chart_data)) for i in range(len(chart_data))]
                chart_data_sorted = chart_data.sort_values()
                bars = ax.barh(chart_data_sorted.index, chart_data_sorted.values, color=colors_list, edgecolor=DARK_BG)
                for bar, val in zip(bars, chart_data_sorted.values):
                    ax.text(val + chart_data.max()*0.01, bar.get_y() + bar.get_height()/2, format_number(val), va="center", fontsize=8, fontweight="bold", color=TEXT_COLOR)
                plt.yticks(fontsize=8)

            elif chart_type == "Line Chart":
                ax.plot(chart_data.index, chart_data.values, color="#00BCD4", linewidth=2.5, marker="o", markersize=6, markerfacecolor=DARK_BG, markeredgewidth=2, markeredgecolor="#00BCD4")
                ax.fill_between(range(len(chart_data)), chart_data.values, alpha=0.15, color="#00BCD4")
                plt.xticks(range(len(chart_data)), chart_data.index, rotation=30, ha="right", fontsize=8)

            elif chart_type == "Pie Chart":
                fig.clear()
                ax = fig.add_subplot(111)
                fig.patch.set_facecolor(DARK_BG)
                ax.set_facecolor(DARK_BG)
                pie_colors = [plt.cm.Blues(0.3 + 0.7 * i / len(chart_data)) for i in range(len(chart_data))]
                wedges, texts, autotexts = ax.pie(chart_data.values, labels=chart_data.index, autopct="%1.1f%%", colors=pie_colors, startangle=140, wedgeprops=dict(edgecolor=DARK_BG, linewidth=2), pctdistance=0.82)
                for text in texts:
                    text.set_color(TEXT_COLOR); text.set_fontsize(8)
                for autotext in autotexts:
                    autotext.set_color("white"); autotext.set_fontsize(8); autotext.set_fontweight("bold")

            ax.set_xlabel(x_axis, color=TEXT_COLOR)
            ax.set_ylabel(y_axis, color=TEXT_COLOR)
            plt.tight_layout()
            st.pyplot(fig)
            plt.close()

        if len(categorical_cols) >= 2 and numeric_cols:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<div class='section-header'>COMPARISON ANALYSIS</div>", unsafe_allow_html=True)

            comp1, comp2 = st.columns(2)

            with comp1:
                if len(categorical_cols) >= 1:
                    cat1 = categorical_cols[0]
                    num1 = numeric_cols[0]
                    st.markdown("<div style='color:#a8c0d6; font-size:0.85rem; font-weight:600; margin-bottom:8px;'>TOP 10 — " + cat1.upper() + " BY " + num1.upper() + "</div>", unsafe_allow_html=True)
                    top10 = df_filtered.groupby(cat1)[num1].sum().sort_values(ascending=False).head(10)
                    fig5, ax5 = plt.subplots(figsize=(6, 5))
                    dark_chart_style(ax5, fig5)
                    colors5 = [plt.cm.Greens(0.4 + 0.6 * i / len(top10)) for i in range(len(top10))]
                    ax5.barh(top10.index[::-1], top10.values[::-1], color=colors5, edgecolor=DARK_BG)
                    for i, val in enumerate(top10.values[::-1]):
                        ax5.text(val + top10.max()*0.01, i, format_number(val), va="center", fontsize=8, fontweight="bold", color=TEXT_COLOR)
                    ax5.set_xlabel(num1, color=TEXT_COLOR)
                    plt.yticks(fontsize=7)
                    plt.tight_layout()
                    st.pyplot(fig5)
                    plt.close()

            with comp2:
                if len(categorical_cols) >= 2:
                    cat2 = categorical_cols[1]
                    num2 = numeric_cols[0]
                    st.markdown("<div style='color:#a8c0d6; font-size:0.85rem; font-weight:600; margin-bottom:8px;'>BREAKDOWN — " + cat2.upper() + " BY " + num2.upper() + "</div>", unsafe_allow_html=True)
                    breakdown = df_filtered.groupby(cat2)[num2].sum().sort_values(ascending=False)
                    fig6, ax6 = plt.subplots(figsize=(6, 5))
                    fig6.patch.set_facecolor(DARK_BG)
                    ax6.set_facecolor(DARK_BG)
                    pie_c = [plt.cm.Oranges(0.4 + 0.6 * i / len(breakdown)) for i in range(len(breakdown))]
                    wedges, texts, autotexts = ax6.pie(breakdown.values, labels=breakdown.index, autopct="%1.1f%%", colors=pie_c, startangle=140, wedgeprops=dict(edgecolor=DARK_BG, linewidth=2), pctdistance=0.82)
                    for text in texts:
                        text.set_color(TEXT_COLOR); text.set_fontsize(8)
                    for autotext in autotexts:
                        autotext.set_color("white"); autotext.set_fontsize(8); autotext.set_fontweight("bold")
                    plt.tight_layout()
                    st.pyplot(fig6)
                    plt.close()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-header'>AUTO INSIGHTS</div>", unsafe_allow_html=True)

        ins1, ins2, ins3 = st.columns(3)

        if categorical_cols and numeric_cols:
            top_cat   = df_filtered.groupby(categorical_cols[0])[numeric_cols[0]].sum().idxmax()
            top_val   = df_filtered.groupby(categorical_cols[0])[numeric_cols[0]].sum().max()
            low_cat   = df_filtered.groupby(categorical_cols[0])[numeric_cols[0]].sum().idxmin()
            avg_val   = df_filtered[numeric_cols[0]].mean()

            with ins1:
                st.markdown(
                    '<div class="insight-card">'
                    '<div class="insight-title">TOP PERFORMER</div>'
                    '<div class="insight-value">' + str(top_cat) + '</div>'
                    '<div style="color:#69f0ae; font-size:0.85rem; margin-top:4px;">' + format_number(top_val) + ' total</div>'
                    '</div>',
                    unsafe_allow_html=True)

            with ins2:
                st.markdown(
                    '<div class="insight-card">'
                    '<div class="insight-title">NEEDS ATTENTION</div>'
                    '<div class="insight-value">' + str(low_cat) + '</div>'
                    '<div style="color:#ff5252; font-size:0.85rem; margin-top:4px;">Lowest in category</div>'
                    '</div>',
                    unsafe_allow_html=True)

            with ins3:
                st.markdown(
                    '<div class="insight-card">'
                    '<div class="insight-title">AVERAGE VALUE</div>'
                    '<div class="insight-value">' + format_number(avg_val) + '</div>'
                    '<div style="color:#64b5f6; font-size:0.85rem; margin-top:4px;">Per row benchmark</div>'
                    '</div>',
                    unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-header'>RAW DATA EXPLORER</div>", unsafe_allow_html=True)

        search_term = st.text_input("Search data", placeholder="Type any keyword to filter rows...")
        if search_term:
            mask = df_filtered.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)
            df_show = df_filtered[mask]
            st.markdown('<div style="color:#64b5f6; font-size:0.85rem;">Found <b>' + str(len(df_show)) + '</b> rows matching <b>' + search_term + '</b></div>', unsafe_allow_html=True)
        else:
            df_show = df_filtered

        st.dataframe(df_show, use_container_width=True, height=350)

        csv_export = df_show.to_csv(index=False)
        st.download_button(label="Export filtered data as CSV", data=csv_export, file_name="pharmaexceliq_export.csv", mime="text/csv")

    except Exception as e:
        st.markdown('<div style="background:#1a0a0a; border:1px solid #f44336; border-radius:12px; padding:1.5rem; color:#ff5252;">Error reading file: ' + str(e) + '<br><br>Please make sure your file is a valid .xlsx, .xls or .csv file.</div>', unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown('<div class="footer">PHARMAEXCELIQ &nbsp;·&nbsp; UNIVERSAL EXCEL INTELLIGENCE &nbsp;·&nbsp; BUILT WITH PYTHON AND STREAMLIT &nbsp;·&nbsp; MBA PROJECT</div>', unsafe_allow_html=True)
