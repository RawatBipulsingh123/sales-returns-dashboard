"""
Advanced Sales & Returns Dashboard
==================================

A production-ready Streamlit application for analyzing sales and returns
data uploaded as an Excel (.xlsx) file. Handles dynamic column mapping,
Gross/Net/Return business logic, year-wise and overall reporting, ad-hoc
visual exploration, and a fully formatted in-memory Excel export.

Run with:
    pip install streamlit pandas numpy plotly xlsxwriter openpyxl
    streamlit run app.py
"""

import io
import re
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# --------------------------------------------------------------------------
# Page Configuration
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Sales & Returns Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown("""
    <style>
    /* Hide top right menu and bottom footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Optimize screen padding for a wider canvas */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 1rem;
        padding-left: 3rem;
        padding-right: 3rem;
    }
    
    /* Give tabs a subtle, modern styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
        border-bottom: 2px solid #e6e6e9;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        font-weight: 600;
    }
    /* Multiselect tags ko laal se hata kar professional Grey/Blue karna */
    span[data-baseweb="tag"] {
        background-color: #E2E8F0 !important;
        color: #0F172A !important;
        border-radius: 4px !important;
        border: none !important;
    }
    
    /* Dropdown aur inputs ke upar ki faltu jagah (labels) ko tight karna */
    .stSelectbox label, .stMultiSelect label, .stDateInput label {
        font-weight: 600 !important;
        color: #475569 !important;
    }

    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 🔒 SECURE LOGIN SYSTEM FOR FREELANCE PITCH
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    
    with col2:
        st.markdown("<h2 style='text-align: center; color: #1F4E78;'>🔒 Client Portal Access</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Advanced Sales & Returns Engine</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        username = st.text_input("User ID", placeholder="Enter your ID")
        password = st.text_input("Password", type="password", placeholder="Enter your password")
        
        if st.button("Secure Login", use_container_width=True):
            if username == "Bipul" and password == "87654321":
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("❌ Invalid ID or Password. Access Denied.")
    
    st.stop()

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state["logged_in"] = False
    st.rerun()
st.sidebar.markdown("---")
# ==========================================


REQUIRED_FIELDS = ["Date", "Store Name", "EANCode", "Product", "Color", "Size", "Quantity"]
AGG_COLUMNS = ["Gross Sales", "Returns", "Net Sales", "Sales %", "Return Rate %"]


# --------------------------------------------------------------------------
# Caching / Ingestion Helpers
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_excel_file(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes))


def guess_column_index(columns, keywords):
    lowered = [str(c).lower() for c in columns]
    for kw in keywords:
        for i, c in enumerate(lowered):
            if kw in c:
                return i
    return None


# --------------------------------------------------------------------------
# Pre-processing Engine
# --------------------------------------------------------------------------
def preprocess_data(df_raw: pd.DataFrame, mapping: dict):
    df = pd.DataFrame(index=df_raw.index)
    for field in REQUIRED_FIELDS:
        df[field] = df_raw[mapping[field]].values

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Year"] = df["Date"].dt.year

    rows_before = len(df)
    df = df.dropna(subset=["Year"])
    dropped_rows = rows_before - len(df)
    df["Year"] = df["Year"].astype(int)

    if df["Quantity"].dtype == object:
        df["Quantity"] = (
            df["Quantity"].astype(str).str.replace(",", "", regex=False).str.strip()
        )
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)

    for col in ["Store Name", "Product","Color", "Size"]:
        df[col] = df[col].where(df[col].notna(), "Unknown")
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"": "Unknown", "nan": "Unknown", "None": "Unknown"})

    df = df.reset_index(drop=True)
    return df, dropped_rows


# --------------------------------------------------------------------------
# Business Logic
# --------------------------------------------------------------------------
def aggregate_dimension(df: pd.DataFrame, dimension_col: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[dimension_col] + AGG_COLUMNS)

    grouped = df.groupby(dimension_col, dropna=False).agg(
        **{
            "Gross Sales": ("Quantity", lambda s: s[s > 0].sum()),
            "Returns": ("Quantity", lambda s: s[s < 0].sum()),
        }
    ).reset_index()

    grouped["Net Sales"] = grouped["Gross Sales"] + grouped["Returns"]

    total_net_sales = grouped["Net Sales"].sum()
    if total_net_sales != 0:
        grouped["Sales %"] = grouped["Net Sales"] / total_net_sales
    else:
        grouped["Sales %"] = 0.0

    grouped["Return Rate %"] = np.where(
        grouped["Gross Sales"] != 0,
        grouped["Returns"].abs() / grouped["Gross Sales"],
        0.0,
    )

    grouped = grouped.sort_values("Net Sales", ascending=False).reset_index(drop=True)
    
    total_gross = grouped["Gross Sales"].sum()
    total_returns = grouped["Returns"].sum()
    total_net = grouped["Net Sales"].sum()
    total_return_rate = abs(total_returns) / total_gross if total_gross != 0 else 0.0
    
    total_row = pd.DataFrame({
        dimension_col: ["TOTAL"],
        "Gross Sales": [total_gross],
        "Returns": [total_returns],
        "Net Sales": [total_net],
        "Sales %": [1.0], 
        "Return Rate %": [total_return_rate]
    })
    
    grouped = pd.concat([grouped, total_row], ignore_index=True)

    return grouped[[dimension_col] + AGG_COLUMNS]


def style_aggregated_df(df: pd.DataFrame):
    def highlight_high_returns(val):
        try:
            if val > 0.15:
                return "color: #C00000; font-weight: bold; background-color: #FDEDEC"
        except TypeError:
            pass
        return ""

    styler = df.style.format(
        {
            "Gross Sales": "{:,.0f}",
            "Returns": "{:,.0f}",
            "Net Sales": "{:,.0f}",
            "Sales %": "{:.2%}",
            "Return Rate %": "{:.2%}",
        }
    )

    try:
        styler = styler.map(highlight_high_returns, subset=["Return Rate %"])
    except AttributeError:
        styler = styler.applymap(highlight_high_returns, subset=["Return Rate %"])

    return styler


def render_kpis(df_subset: pd.DataFrame):
    gross = df_subset.loc[df_subset["Quantity"] > 0, "Quantity"].sum()
    returns = df_subset.loc[df_subset["Quantity"] < 0, "Quantity"].sum()
    net = gross + returns
    return_rate = (abs(returns) / gross) if gross != 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Gross Sales", f"{gross:,.0f}")
    c2.metric("Returns", f"{returns:,.0f}")
    c3.metric("Net Sales", f"{net:,.0f}")
    c4.metric("Return Rate", f"{return_rate:.2%}")


def render_aggregation_section(df_subset: pd.DataFrame, key_prefix: str) -> dict:
    results = {}

    st.markdown("#### 🏬 Performance by Store")
    store_agg = aggregate_dimension(df_subset, "Store Name")
    st.dataframe(style_aggregated_df(store_agg), use_container_width=True, key=f"{key_prefix}_store_df")
    results["Store Name"] = store_agg

    st.markdown("#### 🎨 Performance by Color")
    color_agg = aggregate_dimension(df_subset, "Color")
    st.dataframe(style_aggregated_df(color_agg), use_container_width=True, key=f"{key_prefix}_color_df")
    results["Color"] = color_agg

    st.markdown("#### 📏 Performance by Size")
    size_agg = aggregate_dimension(df_subset, "Size")
    st.dataframe(style_aggregated_df(size_agg), use_container_width=True, key=f"{key_prefix}_size_df")
    results["Size"] = size_agg

    return results


# --------------------------------------------------------------------------
# Excel Export (fully in-memory)
# --------------------------------------------------------------------------
def generate_excel_export(aggregation_data: dict) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        main_title_format = workbook.add_format(
            {"bold": True, "font_size": 15, "font_color": "#1F4E78"}
        )
        section_title_format = workbook.add_format(
            {"bold": True, "font_size": 12, "font_color": "#1F4E78", "bottom": 1}
        )
        header_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#4472C4",
                "font_color": "white",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        number_format = workbook.add_format({"num_format": "#,##0", "border": 1})
        percent_format = workbook.add_format({"num_format": "0.00%", "border": 1})
        red_percent_format = workbook.add_format(
            {"num_format": "0.00%", "font_color": "#C00000", "bold": True, "border": 1}
        )
        text_format = workbook.add_format({"border": 1})

        year_keys = [k for k in aggregation_data.keys() if k != "Overall"]
        try:
            year_keys = sorted(year_keys, key=lambda x: int(x))
        except (ValueError, TypeError):
            year_keys = sorted(year_keys)
        ordered_keys = year_keys + (["Overall"] if "Overall" in aggregation_data else [])

        for sheet_key in ordered_keys:
            sheet_name = ("Year " + sheet_key if sheet_key != "Overall" else "Overall")[:31]
            worksheet = workbook.add_worksheet(sheet_name)

            worksheet.write(0, 0, f"Sales & Returns Dashboard — {sheet_name}", main_title_format)
            current_row = 2

            for dim_label, agg_df in aggregation_data[sheet_key].items():
                worksheet.write(current_row, 0, f"Performance by {dim_label}", section_title_format)
                current_row += 1

                for col_idx, col_name in enumerate(agg_df.columns):
                    worksheet.write(current_row, col_idx, col_name, header_format)
                current_row += 1

                for _, row in agg_df.iterrows():
                    for col_idx, col_name in enumerate(agg_df.columns):
                        value = row[col_name]
                        if col_name in ("Sales %", "Return Rate %"):
                            is_high_return = col_name == "Return Rate %" and float(value) > 0.15
                            fmt = red_percent_format if is_high_return else percent_format
                            worksheet.write_number(current_row, col_idx, float(value), fmt)
                        elif col_name in ("Gross Sales", "Returns", "Net Sales"):
                            worksheet.write_number(current_row, col_idx, float(value), number_format)
                        else:
                            worksheet.write(current_row, col_idx, str(value), text_format)
                    current_row += 1

                current_row += 2 

            worksheet.set_column(0, 0, 24)
            worksheet.set_column(1, 5, 16)

    output.seek(0)
    return output.getvalue()


# --------------------------------------------------------------------------
# Sidebar: Ingestion & Column Mapping
# --------------------------------------------------------------------------
st.sidebar.title("⚙️ Data Configuration")
uploaded_file = st.sidebar.file_uploader("Upload Sales Data", type=["xlsx"])

if uploaded_file is None:
    st.warning("👋 Please upload an Excel (.xlsx) file using the sidebar to get started.")
    st.stop()

if not uploaded_file.name.lower().endswith(".xlsx"):
    st.warning("Only .xlsx files are supported. Please upload a valid Excel file.")
    st.stop()

file_bytes = uploaded_file.getvalue()
if not file_bytes:
    st.warning("The uploaded file appears to be empty. Please upload a valid .xlsx file.")
    st.stop()

try:
    df_raw = load_excel_file(file_bytes)
except Exception as e:
    st.error(
        f"❌ Could not read the uploaded file. It may be corrupted or not a valid Excel file.\n\nDetails: {e}"
    )
    st.stop()

if df_raw is None or df_raw.empty or len(df_raw.columns) == 0:
    st.warning("The uploaded file contains no data. Please check the file and try again.")
    st.stop()

st.sidebar.success(f"Loaded {len(df_raw):,} rows × {len(df_raw.columns)} columns.")

st.sidebar.markdown("### 🔗 Map Your Columns")
st.sidebar.caption("Match your file's headers to the fields the dashboard needs.")

columns = df_raw.columns.tolist()
mapping = {
    "Date": st.sidebar.selectbox("Date Column", columns, index=guess_column_index(columns, ["date", "invoice", "order"])),
    "EANCode": st.sidebar.selectbox("EAN / Barcode Column", columns, index=guess_column_index(columns, ["ean", "barcode", "sku", "item code"])),
    "Store Name": st.sidebar.selectbox("Store Name Column", columns, index=guess_column_index(columns, ["store", "location", "branch", "shop"])),
    "Product": st.sidebar.selectbox("Product Column", columns, index=guess_column_index(columns, ["product", "item", "article", "style"])),
    "Color": st.sidebar.selectbox("Color Column", columns, index=guess_column_index(columns, ["color", "colour"])),
    "Size": st.sidebar.selectbox("Size Column", columns, index=guess_column_index(columns, ["size"])),
    "Quantity": st.sidebar.selectbox("Quantity Column", columns, index=guess_column_index(columns, ["qty", "quantity", "units"])),
}

if len(set(mapping.values())) < len(mapping):
    st.sidebar.warning(
        "⚠️ Two or more fields are mapped to the same source column. "
        "Double-check this is intentional."
    )
st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 Stock Data")
stock_file = st.sidebar.file_uploader("Upload Stock Inventory", type=["xlsx", "csv", "xls"], key="stock_data")

# --------------------------------------------------------------------------
# Pre-processing
# --------------------------------------------------------------------------
if None in mapping.values():
    st.info("👈 Please select the correct columns from the dropdowns in the sidebar to view the dashboard.")
    st.stop()
df_processed, dropped_rows = preprocess_data(df_raw, mapping)

if df_processed.empty:
    st.warning(
        "After validation, no rows contain a usable Date. Please check your "
        "column mapping and confirm the source data has valid dates."
    )
    st.stop()

# --------------------------------------------------------------------------
# Main Layout
# --------------------------------------------------------------------------
st.title("📊 Advanced Sales & Returns Dashboard")

if dropped_rows > 0:
    st.info(f"ℹ️ {dropped_rows:,} row(s) were excluded due to missing or invalid dates.")

years = sorted(df_processed["Year"].unique().tolist())
tab_labels = [str(y) for y in years] + ["Overall", "🔍 Item Search", "🔨 Custom Visualizations", "📦 Stock vs Sales", "📈 Day Trends", "📊 Executive KPI"]
tabs = st.tabs(tab_labels)

excel_aggregation_data = {}

for i, year in enumerate(years):
    with tabs[i]:
        subset = df_processed[df_processed["Year"] == year]
        st.markdown(f"### Year {year} Summary")
        render_kpis(subset)
        st.divider()
        excel_aggregation_data[str(year)] = render_aggregation_section(subset, f"y{year}")

with tabs[len(years)]:
    st.markdown("### Overall Summary (All Years)")
    render_kpis(df_processed)
    st.divider()
    excel_aggregation_data["Overall"] = render_aggregation_section(df_processed, "overall")

# --------------------------------------------------------------------------
# Item Search & Deep Dive Tab (Master Search Engine)
# --------------------------------------------------------------------------
with tabs[len(years) + 1]:
    st.markdown("### 🔍 Master Search & Ledger")
    st.caption("Filter by Date, Store, or Item to dynamically view sales history and current stock.")

    base_df = df_processed.copy()
    base_df["EANCode"] = base_df["EANCode"].astype(str).str.strip()
    
    if "Date" in base_df.columns:
        base_df["Date"] = pd.to_datetime(base_df["Date"], errors='coerce')
        min_dt = base_df["Date"].min()
        max_dt = base_df["Date"].max()
    else:
        min_dt, max_dt = None, None

    f_row1_1, f_row1_2 = st.columns(2)
    with f_row1_1:
        if pd.notnull(min_dt) and pd.notnull(max_dt):
            sel_dates = st.date_input(
                "📅 Global Date Filter", 
                value=[], 
                min_value=min_dt.date(), 
                max_value=max_dt.date(), 
                key="top_date"
            )
        else:
            sel_dates = []
            st.info("No valid dates found in dataset.")
            
    with f_row1_2:
        store_list = sorted(base_df["Store Name"].dropna().unique().tolist())
        s_store = st.multiselect("🏬 Global Store Filter", store_list, key="top_store")

    item_df = base_df.copy()
    
    f_row2_1, f_row2_2, f_row2_3, f_row2_4 = st.columns(4)
    with f_row2_1:
        ean_list = ["All"] + sorted(item_df["EANCode"].dropna().unique().tolist())
        s_ean = st.selectbox("Barcode / EAN", ean_list, key="top_ean")
        if s_ean != "All": item_df = item_df[item_df["EANCode"] == s_ean]
        
    with f_row2_2:
        prod_list = ["All"] + sorted(item_df["Product"].dropna().unique().tolist())
        s_prod = st.selectbox("Product Name", prod_list, key="top_prod")
        if s_prod != "All": item_df = item_df[item_df["Product"] == s_prod]
        
    with f_row2_3:
        color_list = ["All"] + sorted(item_df["Color"].dropna().unique().tolist())
        s_col = st.selectbox("Color", color_list, key="top_col")
        if s_col != "All": item_df = item_df[item_df["Color"] == s_col]
        
    with f_row2_4:
        size_list = ["All"] + sorted(item_df["Size"].dropna().unique().tolist())
        s_size = st.selectbox("Size", size_list, key="top_size")
        if s_size != "All": item_df = item_df[item_df["Size"] == s_size]

    target_eans = item_df["EANCode"].unique().tolist()
    search_df = item_df.copy()
    
    if len(sel_dates) == 2:
        search_df = search_df[(search_df["Date"].dt.date >= sel_dates[0]) & (search_df["Date"].dt.date <= sel_dates[1])]
    elif len(sel_dates) == 1:
        search_df = search_df[search_df["Date"].dt.date == sel_dates[0]]
        
    if s_store:
        search_df = search_df[search_df["Store Name"].isin(s_store)]

    st.divider()
    
    st.markdown("#### 📊 Sales Snapshot for Current Filters")
    render_kpis(search_df)
    
    c_left, c_right = st.columns(2)
    with c_left:
        st.markdown("#### 🏬 Products Sold (Store-Wise)")
        if not search_df.empty:
            prod_agg = search_df.groupby(["Store Name", "Product", "EANCode"], as_index=False)["Quantity"].sum()
            prod_agg = prod_agg[prod_agg["Quantity"] != 0].sort_values(by=["Store Name", "Quantity"], ascending=[True, False])
            prod_agg = prod_agg.rename(columns={"Quantity": "Total Units Sold"})
            st.dataframe(prod_agg.reset_index(drop=True), use_container_width=True)
        else:
            st.warning("No sales found for these filters.")
            
    with c_right:
        st.markdown("#### 📦 Current Stock Location")
        if stock_file is not None:
            try:
                file_bytes = stock_file.getvalue()
                if stock_file.name.endswith('.csv'):
                    df_stock_search = pd.read_csv(io.BytesIO(file_bytes), skiprows=6, low_memory=False)
                else:
                    df_stock_search = pd.read_excel(io.BytesIO(file_bytes), skiprows=6)
                
                if all(col in df_stock_search.columns for col in ["EANCode", "StoreName", "StockInHand"]):
                    df_stock_search["EANCode"] = df_stock_search["EANCode"].astype(str).str.strip()
                    stock_filtered = df_stock_search.copy()
                    
                    if s_ean != "All":
                        stock_filtered = stock_filtered[stock_filtered["EANCode"] == s_ean]
                    else:
                        if s_prod != "All" and "ProductName" in stock_filtered.columns:
                            stock_filtered = stock_filtered[stock_filtered["ProductName"].astype(str).str.strip().str.upper() == str(s_prod).strip().upper()]
                        if s_col != "All" and "ColorName" in stock_filtered.columns:
                            stock_filtered = stock_filtered[stock_filtered["ColorName"].astype(str).str.strip().str.upper() == str(s_col).strip().upper()]
                        if s_size != "All" and "SizeName" in stock_filtered.columns:
                            stock_filtered = stock_filtered[stock_filtered["SizeName"].astype(str).str.strip().str.upper() == str(s_size).strip().upper()]
                        if s_prod == "All" and s_col == "All" and s_size == "All":
                            stock_filtered = stock_filtered[stock_filtered["EANCode"].isin(target_eans)]

                    if stock_filtered.empty:
                        st.warning("No inventory found for the filtered items in the uploaded Stock file.")
                    else:
                        if s_store:
                            stock_filtered = stock_filtered[stock_filtered["StoreName"].isin(s_store)]
                        if "ProductName" in stock_filtered.columns:
                            stock_agg = stock_filtered.groupby(["StoreName", "ProductName"], as_index=False)["StockInHand"].sum()
                            stock_agg = stock_agg.rename(columns={"StoreName": "Store", "ProductName": "Product", "StockInHand": "Units Available"})
                        else:
                            stock_agg = stock_filtered.groupby(["StoreName"], as_index=False)["StockInHand"].sum()
                            stock_agg = stock_agg.rename(columns={"StoreName": "Store", "StockInHand": "Units Available"})
                        st.dataframe(stock_agg.sort_values("Units Available", ascending=False).reset_index(drop=True), use_container_width=True)
                else:
                    st.warning("⚠️ Stock file is missing EANCode, StoreName, or StockInHand columns.")
            except Exception as e:
                st.error(f"❌ Error loading stock file: {e}")
        else:
            st.info("Upload the Stock file in the sidebar to see live inventory locations.")
    
    st.divider()
    st.markdown("#### 📅 Detailed Transaction Ledger")
    
    if not search_df.empty and "Date" in search_df.columns:
        ledger_df = search_df[["Date", "Store Name", "Product", "EANCode", "Quantity"]].copy()
        ledger_df = ledger_df.sort_values(by="Date", ascending=False)
        ledger_df["Date"] = ledger_df["Date"].dt.strftime('%d-%m-%Y')
        ledger_df = ledger_df.rename(columns={"Date": "Bill Date", "Store Name": "Sold At Store", "Quantity": "Units Sold"})
        st.dataframe(ledger_df.reset_index(drop=True), use_container_width=True)
    else:
        st.info("No transaction data available for this selection.")

# --------------------------------------------------------------------------
# Custom Visualizations Tab
# --------------------------------------------------------------------------
with tabs[len(years) + 2]:
    st.markdown("### 🔧 Build Your Own Visualizations")
    st.caption("Apply global filters first, then build up to two side-by-side charts.")

    viz_df = df_processed.copy()

    st.markdown("#### 🔍 Global Filters")
    if "Year" not in viz_df.columns and "Date" in viz_df.columns:
        viz_df["Year"] = pd.to_datetime(viz_df["Date"]).dt.year

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        if "Year" in viz_df.columns:
            year_opts = ["All"] + sorted(viz_df["Year"].dropna().astype(str).unique().tolist())
            sel_year = st.selectbox("Filter by Year", year_opts, key="g_year")
            if sel_year != "All": 
                viz_df = viz_df[viz_df["Year"].astype(str) == sel_year]
        else:
            st.selectbox("Filter by Year", ["All"], key="g_year_disabled", disabled=True)
            
    with f2:
        if "Product" in viz_df.columns:
            prod_opts = ["All"] + sorted(viz_df["Product"].dropna().unique().tolist())
        else:
            prod_opts = ["All"]
        sel_prod = st.selectbox("Filter by Product", prod_opts, key="g_prod")
        if sel_prod != "All" and "Product" in viz_df.columns: 
            viz_df = viz_df[viz_df["Product"] == sel_prod]
            
    with f3:
        store_opts = ["All"] + sorted(viz_df["Store Name"].dropna().unique().tolist())
        sel_store = st.selectbox("Filter by Store", store_opts, key="g_store")
        if sel_store != "All": 
            viz_df = viz_df[viz_df["Store Name"] == sel_store]
            
    with f4:
        color_opts = ["All"] + sorted(viz_df["Color"].dropna().unique().tolist())
        sel_color = st.selectbox("Filter by Color", color_opts, key="g_color")
        if sel_color != "All": 
            viz_df = viz_df[viz_df["Color"] == sel_color]
        
    st.divider()

    chart_col1, chart_col2 = st.columns(2)

    def render_chart(container, chart_id):
        with container:
            st.markdown(f"**Chart {chart_id}**")
            c_a, c_b = st.columns(2)
            with c_a: x_axis = st.selectbox("X-Axis", viz_df.columns.tolist(), key=f"x_{chart_id}")
            with c_b: y_axis = st.selectbox("Y-Axis", viz_df.columns.tolist(), index=min(1, len(viz_df.columns)-1), key=f"y_{chart_id}")
            
            c_c, c_d = st.columns(2)
            with c_c: chart_type = st.selectbox("Type", ["Bar", "Line", "Pie"], key=f"type_{chart_id}")
            with c_d: top_n = st.number_input("Top 'N'", min_value=1, max_value=50, value=10, key=f"top_{chart_id}")

            numeric_cols = viz_df.select_dtypes(include=[np.number]).columns.tolist()

            if x_axis == y_axis:
                st.error("⚠️ Axes must be different.")
            elif y_axis not in numeric_cols:
                st.error("⚠️ Y-Axis must be numeric.")
            else:
                try:
                    clean_viz_df = viz_df[viz_df[x_axis].astype(str) != "TOTAL"]
                    plot_df = clean_viz_df.groupby(x_axis, as_index=False, dropna=False)[y_axis].sum()
                    
                    if len(plot_df) > top_n and chart_type in ("Bar", "Pie"):
                        plot_df = plot_df.sort_values(y_axis, ascending=False).head(top_n)
                        
                    if chart_type == "Pie" and (plot_df[y_axis] < 0).any():
                        st.warning("⚠️ Cannot plot negative values on Pie chart.")
                    else:
                        title = f"Top {top_n} {x_axis} by {y_axis}"
                        if chart_type == "Bar": 
                            fig = px.bar(plot_df, x=x_axis, y=y_axis, title=title, text_auto=True)
                            fig.update_traces(textposition='outside')
                        elif chart_type == "Line": 
                            fig = px.line(plot_df.sort_values(x_axis), x=x_axis, y=y_axis, title=title, text=y_axis, markers=True)
                            fig.update_traces(textposition='top center')
                        else: 
                            fig = px.pie(plot_df, names=x_axis, values=y_axis, title=title)
                            fig.update_traces(textposition='inside', textinfo='percent+label+value')
                        
                        fig.update_layout(margin=dict(t=40, b=20, l=10, r=10))
                        st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error(f"⚠️ Error rendering chart: {e}")

    render_chart(chart_col1, "1")
    render_chart(chart_col2, "2")
    st.divider()
    chart_col3, chart_col4 = st.columns(2)
    render_chart(chart_col3, "3")


# ---------------------------------------------------------
# Stock vs Sales Analysis Tab
# ---------------------------------------------------------
with tabs[-2]:
    st.markdown("### ⚖️ Comprehensive Stock vs Sales Analysis")
    
    if stock_file is None:
        st.info("👈 Please upload the Stock Data (CSV/Excel) in the sidebar to unlock this analysis.")
    else:
        try:
            if stock_file.name.endswith('.csv'):
                df_stock_raw = pd.read_csv(stock_file, skiprows=6, low_memory=False)
            else:
                df_stock_raw = pd.read_excel(stock_file, skiprows=6)
                
            if "StoreName" not in df_stock_raw.columns and "StockInHand" not in df_stock_raw.columns:
                stock_file.seek(0)
                if stock_file.name.endswith('.csv'):
                    df_stock_raw = pd.read_csv(stock_file, low_memory=False)
                else:
                    df_stock_raw = pd.read_excel(stock_file)

            col_map = {
                "Eancode": "EANCode",
                "Productname": "ProductName",
                "Stock Qty. ": "StockInHand",
                "Stock Qty.": "StockInHand",
                "Storename": "StoreName",
                "Color": "ColorName",
                "Size": "SizeName"
            }
            df_stock_raw.rename(columns=col_map, inplace=True)
            
            stock_base_cols = ["EANCode", "StoreName", "StockInHand"]
            missing = [c for c in stock_base_cols if c not in df_stock_raw.columns]
            
            if missing:
                st.error(f"⚠️ Missing critical columns in Stock file: {missing}")
            else:
                keep_cols = stock_base_cols + [c for c in ["ProductName", "ColorName", "SizeName"] if c in df_stock_raw.columns]
                df_stock = df_stock_raw[keep_cols].copy()
                
                if "ColorName" in df_stock.columns: df_stock.rename(columns={"ColorName": "Color"}, inplace=True)
                if "SizeName" in df_stock.columns: df_stock.rename(columns={"SizeName": "Size"}, inplace=True)
                if "StoreName" in df_stock.columns: df_stock.rename(columns={"StoreName": "Store Name"}, inplace=True)
                if "ProductName" in df_stock.columns: df_stock.rename(columns={"ProductName": "Product"}, inplace=True)
                
                df_processed["EANCode"] = df_processed["EANCode"].astype(str).str.strip()
                df_stock["EANCode"] = df_stock["EANCode"].astype(str).str.strip()
                
                store_name_mapping = {
                    "VIVIANATHANESHORE": "VIVIANA THANE STORE",
                    "BSLAKESHORE": "VIVIANA THANE STORE",
                    "SEAWOODGRANDCENTRALMALL": "SEAWOOD GRAND CENTRAL MALL",
                    "PANVELORIONMALL": "PANVEL ORION MALL",
                    "AHMEDBADONE": "AHMEDABAD ONE", 
                    "RMBORIVALI": "RM BORIVALI",
                    "THECAPITALMALLVASAI": "THE CAPITAL MALL VASAI",
                    "RMVADODARA": "RM VADODARA",
                    "RMUDHANA": "RM UDHANA"
                }
                df_stock['Store Name'] = df_stock['Store Name'].replace(store_name_mapping)
                df_processed["Store Name"] = df_processed["Store Name"].astype(str).str.upper().str.strip()
                df_stock["Store Name"] = df_stock["Store Name"].astype(str).str.upper().str.strip()
            
                st.markdown("##### 🔍 Apply Filters")
                def standardize_store_name(name):
                    name = str(name).strip()
                    name = re.sub(r'^Ethnicity\s*-\s*', 'ET - ', name, flags=re.IGNORECASE)
                    name = re.sub(r'^EC\s*-\s*', 'ET - ', name, flags=re.IGNORECASE)
                    name = re.sub(r'\s*\(?2\)?$', '', name).strip()
                    return name

                df_processed["Store Name"] = df_processed["Store Name"].apply(standardize_store_name)
                
                search_mode = st.radio("Search By:", ["Product Name", "EAN Code"], horizontal=True)
                
                col1, col2 = st.columns(2)
                df_processed["Date"] = pd.to_datetime(df_processed["Date"], errors='coerce')
                df_processed["Year"] = df_processed["Date"].dt.year
                available_years = ["Overall"] + sorted(df_processed["Year"].dropna().astype(int).unique().tolist())
                
                with col1:
                    selected_year = st.selectbox("📅 Select Year", available_years)
                
                with col2:
                    if search_mode == "Product Name":
                        available_options = sorted(df_processed["Product"].dropna().unique().tolist())
                        selected_items = st.multiselect("👕 Select Products (Leave blank for All)", available_options)
                    else:
                        available_options = sorted(df_processed["EANCode"].dropna().astype(str).unique().tolist())
                        selected_items = st.multiselect("🏷️ Select EAN Codes (Leave blank for All)", available_options)
                
                col3, col4 = st.columns([3, 1])
                with col3:
                    available_stores = sorted(df_processed["Store Name"].dropna().unique().tolist())
                    target_defaults = [
                        "ET - ELPRO MALL PUNE", "ET - SEAWOOD GRAND CENTRAL MALL", 
                        "ET - PANVEL ORION MALL", "ET - AHMEDABAD ONE", 
                        "ET - RM BORIVALI", "ET - THE CAPITAL MALL VASAI", 
                        "ET - RM VADODARA", "ET - RM UDHANA"
                    ]
                    default_stores = [s for s in target_defaults if s in available_stores]
                    selected_stores = st.multiselect("🏬 Select Stores to Analyze", available_stores, default=default_stores)
                
                with col4:
                    st.markdown("<br>", unsafe_allow_html=True)
                    show_send_qty = st.checkbox("Show Send Qty & %", value=False)   
                    show_returns = st.checkbox("Show Return Qty (-)", value=False)
                
                st.divider()

                if not selected_items:
                    all_product_sales = df_processed.copy()
                    filtered_stock = df_stock.copy()
                    ui_title_item = "All Products"
                else:
                    if search_mode == "Product Name":
                        all_product_sales = df_processed[df_processed["Product"].isin(selected_items)].copy()
                        upper_items = [str(x).strip().upper() for x in selected_items]
                        filtered_stock = df_stock[df_stock["Product"].astype(str).str.strip().str.upper().isin(upper_items)].copy()
                    else:
                        str_items = [str(x).strip() for x in selected_items]
                        all_product_sales = df_processed[df_processed["EANCode"].astype(str).str.strip().isin(str_items)].copy()
                        upper_items = [str(x).strip().upper() for x in selected_items]
                        filtered_stock = df_stock[df_stock["EANCode"].astype(str).str.strip().str.upper().isin(upper_items)].copy()
                    
                    ui_title_item = f"{len(selected_items)} Items Selected"
                
                filtered_sales = all_product_sales.copy()
                
                if selected_year != "Overall":
                    filtered_sales = filtered_sales[filtered_sales["Year"] == selected_year]
                
                if selected_stores:
                    filtered_sales = filtered_sales[filtered_sales["Store Name"].isin(selected_stores)]
                    filtered_stock = filtered_stock[filtered_stock["Store Name"].isin(selected_stores)]
                    
                for col in ["Color", "Size"]:
                    if col in filtered_sales.columns:
                        filtered_sales[col] = filtered_sales[col].astype(str).str.upper().str.strip()
                    if col in filtered_stock.columns:
                        filtered_stock[col] = filtered_stock[col].astype(str).str.upper().str.strip()

                def generate_table(sales_df, stock_df, group_col):
                    if group_col not in sales_df.columns or group_col not in stock_df.columns:
                        return None
                    
                    # 1. Total Net Sales
                    s_agg = sales_df.groupby([group_col], as_index=False)["Quantity"].sum()
                    s_agg.rename(columns={"Quantity": "Sales Qty"}, inplace=True)
                    
                    # 2. Sirf Negative (Returns) nikalna
                    ret_agg = sales_df[sales_df["Quantity"] < 0].groupby([group_col], as_index=False)["Quantity"].sum()
                    ret_agg.rename(columns={"Quantity": "Return Qty"}, inplace=True)
                    
                    # 3. Stock uthana
                    k_agg = stock_df.groupby([group_col], as_index=False)["StockInHand"].sum()
                    
                    # Sabko aapas mein merge karna
                    merged = pd.merge(s_agg, k_agg, on=[group_col], how="outer").fillna(0)
                    merged = pd.merge(merged, ret_agg, on=[group_col], how="left").fillna(0)
                    
                    merged["Sales Qty"] = merged["Sales Qty"].astype(int)
                    merged["Return Qty"] = merged["Return Qty"].astype(int)
                    merged["Current StoreStock"] = merged["StockInHand"].astype(int)
                    merged["Send Qty"] = merged["Sales Qty"] + merged["Current StoreStock"]
                    
                    merged["Percentage"] = (merged["Sales Qty"] / merged["Send Qty"].replace(0, 1)) * 100
                    merged["Percentage"] = merged["Percentage"].round(1).astype(str) + "%"
                    
                    # UI check ke hisaab se columns dikhana
                    display_cols = [group_col, "Sales Qty"]
                    
                    if show_returns:
                        display_cols.append("Return Qty")
                        
                    if show_send_qty:
                        display_cols.extend(["Send Qty", "Current StoreStock", "Percentage"])
                    else:
                        display_cols.append("Current StoreStock")
                        
                    final = merged[display_cols].sort_values(by="Sales Qty", ascending=False).reset_index(drop=True)
                    return final

                items_to_display = selected_items if selected_items else ["All Products"]
                
                for item in items_to_display:
                    if item == "All Products":
                        item_sales_ui = filtered_sales
                        item_stock_ui = filtered_stock
                        st.markdown(f"### 📦 Analytics for: **All Products**")
                    else:
                        if search_mode == "Product Name":
                            item_sales_ui = filtered_sales[filtered_sales["Product"] == item]
                            item_stock_ui = filtered_stock[filtered_stock["Product"].astype(str).str.strip().str.upper() == str(item).strip().upper()]
                        else:
                            item_sales_ui = filtered_sales[filtered_sales["EANCode"].astype(str).str.strip() == str(item).strip()]
                            item_stock_ui = filtered_stock[filtered_stock["EANCode"].astype(str).str.strip().str.upper() == str(item).strip().upper()]
                        st.markdown(f"### 📦 Analytics for: **{item}**")
                        
                    st.caption(f"Year: **{selected_year}** | Stores: **{len(selected_stores)} Selected**")
                    
                    st.markdown("#### 🏢 Store-wise Summary")
                    df_store = generate_table(item_sales_ui, item_stock_ui, "Store Name")
                    if df_store is not None: st.dataframe(df_store, use_container_width=True)
                    
                    col_left, col_right = st.columns(2)
                    with col_left:
                        st.markdown("#### 🎨 Color-wise Summary")
                        df_color = generate_table(item_sales_ui, item_stock_ui, "Color")
                        if df_color is not None:
                            st.dataframe(df_color, use_container_width=True)
                        else:
                            st.warning("Color data missing.")
                            
                    with col_right:
                        st.markdown("#### 📏 Size-wise Summary")
                        df_size = generate_table(item_sales_ui, item_stock_ui, "Size")
                        if df_size is not None:
                            st.dataframe(df_size, use_container_width=True)
                        else:
                            st.warning("Size data missing.")
                            
                    st.divider()
                
                if 'df_store' in locals() and df_store is not None:
                    try:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            workbook = writer.book
                            
                            header_format = workbook.add_format({'bg_color': '#4F81BD', 'font_color': 'white', 'bold': True, 'border': 1})
                            pct_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
                            total_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
                            normal_format = workbook.add_format({'border': 1})
                            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})
                            
                            items_to_export = selected_items if selected_items else ["All Products"]
                            
                            for item in items_to_export:
                                if item == "All Products":
                                    item_sales = filtered_sales
                                    item_stock = filtered_stock
                                else:
                                    if search_mode == "Product Name":
                                        item_sales = filtered_sales[filtered_sales["Product"] == item]
                                        item_stock = filtered_stock[filtered_stock["Product"].astype(str).str.strip().str.upper() == str(item).strip().upper()]
                                    else:
                                        item_sales = filtered_sales[filtered_sales["EANCode"].astype(str).str.strip() == str(item).strip()]
                                        item_stock = filtered_stock[filtered_stock["EANCode"].astype(str).str.strip().str.upper() == str(item).strip().upper()]
                                
                                i_store = generate_table(item_sales, item_stock, "Store Name")
                                i_color = generate_table(item_sales, item_stock, "Color")
                                i_size = generate_table(item_sales, item_stock, "Size")
                                
                                sheet_name = re.sub(r'[\\/*?:\[\]]', '', str(item))[:31]
                                base_sheet_name = sheet_name
                                counter = 1
                                while sheet_name in writer.sheets:
                                    suffix = f"_{counter}"
                                    sheet_name = base_sheet_name[:31-len(suffix)] + suffix
                                    counter += 1
                                
                                export_sheets = {
                                    'Store Summary': i_store,
                                    'Color Summary': i_color,
                                    'Size Summary': i_size
                                }
                                
                                current_row = 0
                                
                                for table_title, table_df in export_sheets.items():
                                    if table_df is None or table_df.empty:
                                        continue
                                        
                                    excel_df = table_df.copy() 
                                    sales_col = "Sales Qty"
                                    stock_col = "Current StoreStock"
                                    
                                    total_sales = excel_df[sales_col].sum()
                                    total_stock = excel_df[stock_col].sum()
                                    
                                    total_row = pd.DataFrame({
                                        excel_df.columns[0]: ["FINAL TOTAL"],
                                        sales_col: [total_sales],
                                        stock_col: [total_stock]
                                    })
                                    
                                    if "Send Qty" in excel_df.columns:
                                        total_row["Send Qty"] = [total_sales + total_stock]

                                    if "Return Qty" in excel_df.columns:
                                        total_row["Return Qty"] = [excel_df["Return Qty"].sum()]
                                        
                                    if "Percentage" in excel_df.columns:
                                        total_row["Percentage"] = [f"{(total_sales / (total_sales + total_stock) * 100):.1f}%" if (total_sales + total_stock) > 0 else "0.0%"]
                                    
                                    excel_df = pd.concat([excel_df, total_row], ignore_index=True)
                                    
                                    worksheet = writer.sheets.get(sheet_name)
                                    if not worksheet:
                                        worksheet = workbook.add_worksheet(sheet_name)
                                        
                                    worksheet.write(current_row, 0, table_title, title_format)
                                    current_row += 1
                                    
                                    excel_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=current_row)
                                    
                                    for col_num, value in enumerate(excel_df.columns.values):
                                        worksheet.write(current_row, col_num, value, header_format)
                                        
                                    for row_idx in range(len(excel_df)): 
                                        is_last_row = (row_idx == len(excel_df) - 1)
                                        
                                        for col_num in range(len(excel_df.columns)):
                                            col_name = excel_df.columns[col_num]
                                            val = excel_df.iloc[row_idx, col_num]
                                            
                                            if "%" in col_name or col_name == "Percentage":
                                                fmt = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'num_format': '0.00%'}) if is_last_row else pct_format
                                                if isinstance(val, str) and '%' in val:
                                                    try:
                                                        val = float(val.replace('%','')) / 100.0
                                                    except ValueError:
                                                        val = 0.0
                                                worksheet.write(current_row + 1 + row_idx, col_num, val, fmt)
                                            else:
                                                fmt = total_format if is_last_row else normal_format
                                                worksheet.write(current_row + 1 + row_idx, col_num, val, fmt)
                                                
                                    current_row += len(excel_df) + 3 
                                    
                                if sheet_name in writer.sheets:
                                    worksheet = writer.sheets[sheet_name]
                                    worksheet.set_column(0, 0, 35)
                                    worksheet.set_column(1, 10, 15)
                                    
                        excel_data = output.getvalue()
                        
                        st.download_button(
                            label=f"📥 Download Multi-Product Report ({len(items_to_export)} Sheets)",
                            data=excel_data,
                            file_name="Multi_Product_Stock_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"Error generating multi-sheet Excel: {e}")
        except Exception as e:
            st.error(f"Error processing stock data: {e}")

# --------------------------------------------------------------------------
# Sidebar: Excel Export
# --------------------------------------------------------------------------
st.sidebar.markdown("### 📥 Export")
try:
    excel_bytes = generate_excel_export(excel_aggregation_data)
    st.sidebar.download_button(
        label="Download Full Dashboard (.xlsx)",
        data=excel_bytes,
        file_name="Sales_Returns_Dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception as e:
    st.sidebar.error(f"Could not generate Excel export: {e}")

# ---------------------------------------------------------
# ---------------------------------------------------------
# NEW ADVANCED TAB: Day-wise Trend Analysis & Export
# ---------------------------------------------------------
with tabs[-2]: # Ensure yeh tera second-last tab ho, warna index theek kar lena
    st.markdown("### 📅 Day-wise Trend Analysis")
    
    df_trend = df_processed.copy()
    df_trend["Date"] = pd.to_datetime(df_trend["Date"], errors='coerce')
    df_trend = df_trend.dropna(subset=["Date"])
    df_trend["DayOfWeek"] = df_trend["Date"].dt.day_name()
    
    min_date = df_trend["Date"].min().date() if not df_trend.empty else None
    max_date = df_trend["Date"].max().date() if not df_trend.empty else None
    
    # Timeline aur Filters ko professional layout mein set kiya
    col1, col2, col3 = st.columns([1.5, 1, 1])
    
    with col1:
        st.markdown("##### 🗓️ Timeline (Timely Framework)")
        if min_date and max_date:
            # 1. Naye Quick Action Buttons (Fiverr Clients love this)
            quick_date = st.radio("Quick Select", ["All Time", "Last 7 Days", "This Month", "Custom"], horizontal=True, label_visibility="collapsed")
            
            if quick_date == "Last 7 Days":
                selected_dates = [max_date - pd.Timedelta(days=7), max_date]
            elif quick_date == "This Month":
                selected_dates = [max_date.replace(day=1), max_date]
            elif quick_date == "Custom":
                selected_dates = st.date_input("Select Range", [min_date, max_date], min_value=min_date, max_value=max_date, label_visibility="collapsed")
            else:
                selected_dates = [min_date, max_date]
        else:
            selected_dates = []
            st.warning("No valid dates found.")
            
    with col2:
        st.markdown("##### 👕 Product Filter")
        available_items = ["All"] + sorted(df_trend["Product"].dropna().unique().tolist())
        selected_item_trend = st.selectbox("Product", available_items, key="trend_product", label_visibility="collapsed")
        
    with col3:
        st.markdown("##### 🏬 Store Slicer")
        available_stores = sorted(df_trend["Store Name"].dropna().unique().tolist())
        selected_stores_trend = st.multiselect("Store", available_stores, default=available_stores[:3] if len(available_stores) >= 3 else available_stores, key="trend_store", label_visibility="collapsed")
        
    # Apply Data Filters
    if len(selected_dates) == 2:
        df_trend = df_trend[(df_trend["Date"].dt.date >= selected_dates[0]) & (df_trend["Date"].dt.date <= selected_dates[1])]
    if selected_item_trend != "All":
        df_trend = df_trend[df_trend["Product"] == selected_item_trend]
    if selected_stores_trend:
        df_trend = df_trend[df_trend["Store Name"].isin(selected_stores_trend)]
        
    st.divider()
    
    if not df_trend.empty:
        qty_col = "Sales Qty" if "Sales Qty" in df_trend.columns else ("Quantity" if "Quantity" in df_trend.columns else None)
        day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        st.markdown(f"✅ **Analyzing:** {selected_item_trend} | **Data Points:** {len(df_trend)}")
        
        if qty_col:
            day_data = df_trend.groupby("DayOfWeek")[qty_col].sum().reindex(day_order).fillna(0)
        else:
            day_data = df_trend.groupby("DayOfWeek").size().reindex(day_order).fillna(0)
            
        plot_df = day_data.reset_index()
        plot_df.columns = ["Day", "Total Sales"]
        
        fig = px.bar(
            plot_df, 
            x="Day", 
            y="Total Sales", 
            text="Total Sales",
            color_discrete_sequence=["#1F4E78"] # Corporate Blue instead of harsh red
        )
        fig.update_traces(texttemplate='%{text}', textposition='outside', marker_line_color='black', marker_line_width=1)
        
        # 2. ACTIONABLE Framework: Average Benchmark Line add kardi
        avg_sales = plot_df["Total Sales"].mean()
        fig.add_hline(
            y=avg_sales, 
            line_dash="dash", 
            line_color="#C00000", 
            annotation_text=f"Average Sales: {avg_sales:.1f}", 
            annotation_position="top left"
        )
        
        fig.update_layout(
            xaxis={'categoryorder':'array', 'categoryarray': day_order}, 
            yaxis_title="Units Sold",
            margin=dict(t=30, b=20),
            plot_bgcolor="rgba(0,0,0,0)" # Clean background
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        csv_data = df_trend.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=f"📥 Export '{selected_item_trend}' Filtered Logic",
            data=csv_data,
            file_name=f"Advanced_Day_Trends_{selected_item_trend}.csv",
            mime="text/csv"
        )
    else:
        st.warning("⚠️ No data matches the selected filters.")

# ---------------------------------------------------------
# ---------------------------------------------------------
# NEW TAB: Executive KPI & Donut Chart
# ---------------------------------------------------------
with tabs[-1]:
    st.markdown("### 📊 Executive KPI & Top 5 Stores")
    
    exec_df = df_processed.copy()
    
    # 1. CASCADING SLICERS (Dependent Filters - Multi Select)
    st.markdown("##### ⚙️ Slicers")
    
    sl1, sl2, sl3, sl4 = st.columns(4)
    
    with sl1:
        # "All" hata diya gaya hai kyunki khali box ka matlab 'All' hai
        av_years = sorted(exec_df["Year"].dropna().astype(int).unique().tolist())
        ex_year = st.multiselect("Filter Year", av_years, key="ex_year")
        
    # Step 1: Year filter apply kar (ab == ki jagah .isin() use hoga)
    if ex_year: 
        exec_df = exec_df[exec_df["Year"].isin(ex_year)]
        
    with sl2:
        av_stores = sorted(exec_df["Store Name"].dropna().unique().tolist())
        ex_store = st.multiselect("Filter Store", av_stores, key="ex_store")
        
    # Step 2: Store filter apply kar
    if ex_store:
        exec_df = exec_df[exec_df["Store Name"].isin(ex_store)]
        
    with sl3:
        av_prods = sorted(exec_df["Product"].dropna().unique().tolist())
        ex_prod = st.multiselect("Filter Product", av_prods, key="ex_prod")
        
    # Step 3: Product filter apply kar 
    if ex_prod: 
        exec_df = exec_df[exec_df["Product"].isin(ex_prod)]
        
    with sl4:
        av_colors = sorted(exec_df["Color"].dropna().unique().tolist())
        ex_color = st.multiselect("Filter Color", av_colors, key="ex_color")
        
    # Step 4: Color filter apply kar
    if ex_color: 
        exec_df = exec_df[exec_df["Color"].isin(ex_color)]
        st.divider()
    
    # 2. ACTIONABLE KPIs
    if exec_df.empty:
        st.warning("No data found for these filters.")
    else:
        gross = exec_df.loc[exec_df["Quantity"] > 0, "Quantity"].sum()
        returns = exec_df.loc[exec_df["Quantity"] < 0, "Quantity"].sum()
        net = gross + returns
        
        k1, k2, k3 = st.columns(3)
        k1.metric("Gross Units Sold", f"{gross:,.0f}")
        k2.metric("Units Returned", f"{returns:,.0f}")
        k3.metric("Net Units", f"{net:,.0f}")
        
        st.divider()
        
        # 3. DONUT CHART (All Stores)
        st.markdown("##### 🍩 Store Sales Distribution")
        
        # Chart ke liye sirf positive sales ko count karenge
        sales_only = exec_df[exec_df["Quantity"] > 0]
        if not sales_only.empty:
            store_agg = sales_only.groupby("Store Name", as_index=False)["Quantity"].sum()
            all_stores = store_agg.sort_values("Quantity", ascending=False) # Top 5 limit dɔn pul kɔmɔt
            
            fig = px.pie(
                all_stores, 
                values='Quantity', 
                names='Store Name', 
                hole=0.5,
                title="All Stores Sales (Filtered)",
                color_discrete_sequence=px.colors.qualitative.Bold # Brayt kɔlɔ dɛn
            )
            fig.update_traces(textposition='inside', textinfo='percent+label+value')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No positive sales data available to plot.")


# ---------------------------------------------------------
# ---------------------------------------------------------
# ---------------------------------------------------------
# NEW ADVANCED TAB: Salesman ROI & Profitability
# ---------------------------------------------------------
with tabs[-1]: 
    st.markdown("### 💸 Salesman ROI & Profitability Analysis")
    
    # 1. Base Data Rescue Logic
    roi_df_base = df_processed.copy()
    advanced_cols = ['Mrp', 'WSP', 'Amount', 'Salesman']
    try:
        for col in advanced_cols:
            if col in df_raw.columns and col not in roi_df_base.columns:
                roi_df_base = roi_df_base.join(df_raw[[col]])
    except NameError:
        pass
        
    missing = [c for c in advanced_cols if c not in roi_df_base.columns]
    
    if missing:
        st.error(f"⚠️ Missing Columns: {', '.join(missing)}")
        st.info("Make sure your Excel file has exact column names: 'Mrp', 'WSP', 'Amount', 'Salesman'.")
    else:
        # Pre-process base math for Global Metrics
        for col in ['Mrp', 'WSP', 'Amount', 'Quantity']:
            roi_df_base[col] = pd.to_numeric(roi_df_base[col], errors='coerce').fillna(0)
            
        roi_df_base['Total Cost'] = roi_df_base['WSP'] * roi_df_base['Quantity']
        roi_df_base['Gross Margin'] = roi_df_base['Amount'] - roi_df_base['Total Cost']
        roi_df_base['Discount Given'] = (roi_df_base['Mrp'] * roi_df_base['Quantity']) - roi_df_base['Amount']
        
        # --- GLOBAL OVERVIEW (Static) ---
        st.markdown("##### 🌍 Global Company Overview (All Data)")
        g1, g2, g3, g4 = st.columns(4)
        g_sales = roi_df_base['Amount'].sum()
        g_cost = roi_df_base['Total Cost'].sum()
        g_margin = roi_df_base['Gross Margin'].sum()
        g_margin_pct = (g_margin / g_sales * 100) if g_sales > 0 else 0
        
        g1.metric("Total Company Revenue", f"₹{g_sales:,.0f}")
        g2.metric("Total Company Cost", f"₹{g_cost:,.0f}")
        g3.metric("Global Gross Margin", f"₹{g_margin:,.0f}", f"{g_margin_pct:.1f}%")
        g4.metric("Total Discount Leaked", f"₹{roi_df_base['Discount Given'].sum():,.0f}")
        
        st.divider()

        # --- CASCADING SLICERS ---
        st.markdown("##### ⚙️ Advanced Filters for Deep Dive")
        f1, f2 = st.columns(2)
        roi_df_filtered = roi_df_base.copy() # Create a copy for filtering
        
        with f1:
            store_col = 'Store Name' if 'Store Name' in roi_df_filtered.columns else ('StoreName' if 'StoreName' in roi_df_filtered.columns else None)
            if store_col:
                av_stores = sorted(roi_df_filtered[store_col].dropna().astype(str).unique().tolist())
                sel_stores = st.multiselect("Filter by Store", av_stores, key="roi_store")
                if sel_stores:
                    roi_df_filtered = roi_df_filtered[roi_df_filtered[store_col].isin(sel_stores)]
            else:
                st.warning("Store column not mapped for filtering.")
                
        with f2:
            av_salesman = sorted(roi_df_filtered['Salesman'].dropna().astype(str).unique().tolist())
            sel_salesman = st.multiselect("Filter by Salesman", av_salesman, key="roi_salesman")
            if sel_salesman:
                roi_df_filtered = roi_df_filtered[roi_df_filtered['Salesman'].isin(sel_salesman)]

        st.divider()
        
        # --- FILTERED OVERVIEW (Dynamic) ---
        st.markdown("##### 🎯 Selection Overview")
        if roi_df_filtered.empty:
             st.warning("No data found for this selection.")
        else:
            k1, k2, k3, k4 = st.columns(4)
            f_sales = roi_df_filtered['Amount'].sum()
            f_cost = roi_df_filtered['Total Cost'].sum()
            f_margin = roi_df_filtered['Gross Margin'].sum()
            f_margin_pct = (f_margin / f_sales * 100) if f_sales > 0 else 0
            
            k1.metric("Filtered Revenue", f"₹{f_sales:,.0f}")
            k2.metric("Filtered Cost", f"₹{f_cost:,.0f}")
            k3.metric("Filtered Margin", f"₹{f_margin:,.0f}", f"{f_margin_pct:.1f}%")
            k4.metric("Filtered Discount", f"₹{roi_df_filtered['Discount Given'].sum():,.0f}")
            
            st.divider()
            
            # --- LEADERBOARD (Dynamic based on filters) ---
            st.markdown("##### 🏆 Salesman Leaderboard (By Profitability)")
            
            salesman_agg = roi_df_filtered.groupby('Salesman', as_index=False).agg(
                Units_Sold=('Quantity', 'sum'),
                Total_Revenue=('Amount', 'sum'),
                Gross_Margin=('Gross Margin', 'sum'),
                Discount_Given=('Discount Given', 'sum')
            ).sort_values('Gross_Margin', ascending=False)
            
            try:
                styled_df = salesman_agg.style.format({
                    'Total_Revenue': '₹{:,.2f}', 
                    'Gross_Margin': '₹{:,.2f}', 
                    'Discount_Given': '₹{:,.2f}'
                }).map(lambda x: 'color: #C00000; font-weight: bold' if x < 0 else 'color: #28a745; font-weight: bold', subset=['Gross_Margin'])
            except AttributeError:
                styled_df = salesman_agg.style.format({
                    'Total_Revenue': '₹{:,.2f}', 
                    'Gross_Margin': '₹{:,.2f}', 
                    'Discount_Given': '₹{:,.2f}'
                }).applymap(lambda x: 'color: #C00000; font-weight: bold' if x < 0 else 'color: #28a745; font-weight: bold', subset=['Gross_Margin'])
                
            st.dataframe(styled_df, use_container_width=True)