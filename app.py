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
    </style>
    """, unsafe_allow_html=True)

REQUIRED_FIELDS = ["Date", "Store Name", "EANCode", "Product", "Color", "Size", "Quantity"]
AGG_COLUMNS = ["Gross Sales", "Returns", "Net Sales", "Sales %", "Return Rate %"]


# --------------------------------------------------------------------------
# Caching / Ingestion Helpers
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_excel_file(file_bytes: bytes) -> pd.DataFrame:
    """Reads the first sheet of an uploaded .xlsx file into a DataFrame."""
    return pd.read_excel(io.BytesIO(file_bytes))


def guess_column_index(columns, keywords):
    """Best-effort guess of which source column matches an internal field,
    based on substring matches against a list of likely keywords. Falls
    back to None if nothing matches to prevent blind default mapping."""
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
    """
    Builds a clean, standardized working DataFrame from the raw upload
    using the user-provided column mapping.

    Returns:
        (df_clean, dropped_rows) tuple
    """
    df = pd.DataFrame(index=df_raw.index)
    for field in REQUIRED_FIELDS:
        df[field] = df_raw[mapping[field]].values

    # --- Date -> Year (invalid dates coerced to NaT and dropped) ---
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Year"] = df["Date"].dt.year

    rows_before = len(df)
    df = df.dropna(subset=["Year"])
    dropped_rows = rows_before - len(df)
    df["Year"] = df["Year"].astype(int)

    # --- Quantity -> numeric, missing/unparseable treated as 0 ---
    if df["Quantity"].dtype == object:
        df["Quantity"] = (
            df["Quantity"].astype(str).str.replace(",", "", regex=False).str.strip()
        )
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)

    # --- Categorical cleanup ---
    for col in ["Store Name", "Product","Color", "Size"]:
        df[col] = df[col].where(df[col].notna(), "Unknown")
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"": "Unknown", "nan": "Unknown", "None": "Unknown"})

    df = df.reset_index(drop=True)
    return df, dropped_rows


# --------------------------------------------------------------------------
# Business Logic: Gross / Returns / Net / Sales % / Return Rate %
# --------------------------------------------------------------------------
def aggregate_dimension(df: pd.DataFrame, dimension_col: str) -> pd.DataFrame:
    """
    Aggregates Quantity by a dimension (Store Name / Color / Size) into
    Gross Sales, Returns, Net Sales, Sales %, and Return Rate %.
    """
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
    
    # --- FIXED LOGIC: Calculate and Append Total Row ---
    total_gross = grouped["Gross Sales"].sum()
    total_returns = grouped["Returns"].sum()
    total_net = grouped["Net Sales"].sum()
    total_return_rate = abs(total_returns) / total_gross if total_gross != 0 else 0.0
    
    total_row = pd.DataFrame({
        dimension_col: ["TOTAL"],
        "Gross Sales": [total_gross],
        "Returns": [total_returns],
        "Net Sales": [total_net],
        "Sales %": [1.0],  # Overall total represents 100% of sales
        "Return Rate %": [total_return_rate]
    })
    
    grouped = pd.concat([grouped, total_row], ignore_index=True)
    # ---------------------------------------------------

    return grouped[[dimension_col] + AGG_COLUMNS]


def style_aggregated_df(df: pd.DataFrame):
    """Formats percentages as 0.00% and flags Return Rate % > 15% in red."""

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

    # pandas >= 2.1 renamed applymap -> map; support both.
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
    """Renders Store / Color / Size aggregation tables and returns the
    underlying DataFrames so they can be reused for the Excel export."""
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
    """
    Builds a formatted .xlsx workbook in memory: one sheet per year plus
    an Overall sheet, each containing the Store / Color / Size aggregation
    tables with bold headers, number formatting, and red-flagged high
    return rates.
    """
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

        # Order sheets: years ascending, Overall last.
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

                current_row += 2  # spacing between tables

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
        "❌ Could not read the uploaded file. It may be corrupted or not a valid "
        f"Excel file.\n\nDetails: {e}"
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
stock_file = st.sidebar.file_uploader("Upload Stock Inventory", type=["xlsx"], key="stock_data")

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
tab_labels = [str(y) for y in years] + ["Overall", "🔍 Item Search", "🔨 Custom Visualizations", "📦 Stock vs Sales"]
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
# --------------------------------------------------------------------------
# Item Search & Deep Dive Tab (Master Search Engine)
# --------------------------------------------------------------------------
with tabs[len(years) + 1]:
    st.markdown("### 🔍 Master Search & Ledger")
    st.caption("Filter by Date, Store, or Item to dynamically view sales history and current stock.")

    search_df = df_processed.copy()
    search_df["EANCode"] = search_df["EANCode"].astype(str).str.strip()
    
    if "Date" in search_df.columns:
        search_df["Date"] = pd.to_datetime(search_df["Date"], errors='coerce')
        min_dt = search_df["Date"].min()
        max_dt = search_df["Date"].max()
    else:
        min_dt, max_dt = None, None

    # --- TOP ROW: GLOBAL DATE & STORE FILTERS ---
    f_row1_1, f_row1_2 = st.columns(2)
    with f_row1_1:
        if pd.notnull(min_dt) and pd.notnull(max_dt):
            sel_dates = st.date_input(
                "📅 Global Date Filter", 
                value=[], # Empty by default to show all-time data
                min_value=min_dt.date(), 
                max_value=max_dt.date(), 
                key="top_date"
            )
            if len(sel_dates) == 2:
                search_df = search_df[(search_df["Date"].dt.date >= sel_dates[0]) & (search_df["Date"].dt.date <= sel_dates[1])]
            elif len(sel_dates) == 1:
                search_df = search_df[search_df["Date"].dt.date == sel_dates[0]]
        else:
            st.info("No valid dates found in dataset.")
            
    with f_row1_2:
        store_list = sorted(search_df["Store Name"].dropna().unique().tolist())
        s_store = st.multiselect("🏬 Global Store Filter", store_list, key="top_store")
        if s_store:
            search_df = search_df[search_df["Store Name"].isin(s_store)]

    # --- BOTTOM ROW: ITEM-SPECIFIC FILTERS ---
    f_row2_1, f_row2_2, f_row2_3, f_row2_4 = st.columns(4)
    with f_row2_1:
        ean_list = ["All"] + sorted(search_df["EANCode"].dropna().unique().tolist())
        s_ean = st.selectbox("Barcode / EAN", ean_list, key="top_ean")
        if s_ean != "All": search_df = search_df[search_df["EANCode"] == s_ean]
        
    with f_row2_2:
        prod_list = ["All"] + sorted(search_df["Product"].dropna().unique().tolist())
        s_prod = st.selectbox("Product Name", prod_list, key="top_prod")
        if s_prod != "All": search_df = search_df[search_df["Product"] == s_prod]
        
    with f_row2_3:
        color_list = ["All"] + sorted(search_df["Color"].dropna().unique().tolist())
        s_col = st.selectbox("Color", color_list, key="top_col")
        if s_col != "All": search_df = search_df[search_df["Color"] == s_col]
        
    with f_row2_4:
        size_list = ["All"] + sorted(search_df["Size"].dropna().unique().tolist())
        s_size = st.selectbox("Size", size_list, key="top_size")
        if s_size != "All": search_df = search_df[search_df["Size"] == s_size]

    st.divider()
    
    st.markdown("#### 📊 Sales Snapshot for Current Filters")
    render_kpis(search_df)
    
    c_left, c_right = st.columns(2)
    with c_left:
        st.markdown("#### 🏬 Products Sold (Store-Wise)")
        if not search_df.empty:
            # Added EAN and Product to this table so you know exactly what sold
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
                import io
                file_bytes = stock_file.getvalue()
                
                if stock_file.name.endswith('.csv'):
                    df_stock_search = pd.read_csv(io.BytesIO(file_bytes), skiprows=6, low_memory=False)
                else:
                    df_stock_search = pd.read_excel(io.BytesIO(file_bytes), skiprows=6)
                
                if all(col in df_stock_search.columns for col in ["EANCode", "StoreName", "StockInHand"]):
                    df_stock_search["EANCode"] = df_stock_search["EANCode"].astype(str).str.strip()
                    search_eans = search_df["EANCode"].unique().tolist()
                    stock_filtered = df_stock_search[df_stock_search["EANCode"].isin(search_eans)]
                    
                    if stock_filtered.empty:
                        st.warning("No inventory found for the filtered items in the uploaded Stock file.")
                    else:
                        # Automatically sync the stock display with the Top Store Filter
                        if s_store:
                            stock_filtered = stock_filtered[stock_filtered["StoreName"].isin(s_store)]
                            
                        stock_agg = stock_filtered.groupby(["StoreName", "ProductName"], as_index=False)["StockInHand"].sum()
                        stock_agg = stock_agg.rename(columns={"StoreName": "Store", "ProductName": "Product", "StockInHand": "Units Available"})
                        st.dataframe(stock_agg.sort_values("Units Available", ascending=False).reset_index(drop=True), use_container_width=True)
                else:
                    st.warning("⚠️ Stock file is missing EANCode, StoreName, or StockInHand columns.")
            except Exception as e:
                st.error(f"❌ Error loading stock file: {e}")
        else:
            st.info("Upload the Stock file in the sidebar to see live inventory locations.")
    
    # --- DYNAMIC TRANSACTION LEDGER ---
    st.divider()
    st.markdown("#### 📅 Detailed Transaction Ledger")
    
    if not search_df.empty and "Date" in search_df.columns:
        # Extracted Product and EAN into the ledger view
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

   # --- GLOBAL CASCADING FILTERS ---
    st.markdown("#### 🔍 Global Filters")
    
    # Ensure Year column is extracted for filtering
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

    # --- DUAL CHART RENDERER ---
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
                    # Drop TOTAL row and calculate
                    clean_viz_df = viz_df[viz_df[x_axis].astype(str) != "TOTAL"]
                    plot_df = clean_viz_df.groupby(x_axis, as_index=False, dropna=False)[y_axis].sum()
                    
                    # Strict Top N logic (No 'Other' category)
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

  # First Row: Charts 1 & 2
    render_chart(chart_col1, "1")
    render_chart(chart_col2, "2")
    
    st.divider()
    
    # Second Row: Chart 3
    chart_col3, chart_col4 = st.columns(2)
    render_chart(chart_col3, "3")
    # chart_col4 is left empty to maintain width scaling




# ---------------------------------------------------------
# Stock vs Sales Analysis Tab
# ---------------------------------------------------------
with tabs[-1]:
    st.markdown("### ⚖️ Comprehensive Stock vs Sales Analysis")
    
    if stock_file is None:
        st.info("👈 Please upload the Stock Data (CSV/Excel) in the sidebar to unlock this analysis.")
    else:
        try:
            # 1. Load & Clean Stock Data
            if stock_file.name.endswith('.csv'):
                df_stock_raw = pd.read_csv(stock_file, skiprows=6, low_memory=False)
            else:
                df_stock_raw = pd.read_excel(stock_file, skiprows=6)
            
            stock_base_cols = ["EANCode", "StoreName", "StockInHand"]
            missing = [c for c in stock_base_cols if c not in df_stock_raw.columns]
            
            if missing:
                st.error(f"⚠️ Missing critical columns in Stock file: {missing}")
            else:
                # 2. MEMORY OPTIMIZATION & Column Mapping
                keep_cols = stock_base_cols + [c for c in ["ProductName", "ColorName", "SizeName"] if c in df_stock_raw.columns]
                df_stock = df_stock_raw[keep_cols].copy()
                
                if "ColorName" in df_stock.columns:
                    df_stock.rename(columns={"ColorName": "Color"}, inplace=True)
                if "SizeName" in df_stock.columns:
                    df_stock.rename(columns={"SizeName": "Size"}, inplace=True)
                if "StoreName" in df_stock.columns:
                    df_stock.rename(columns={"StoreName": "Store Name"}, inplace=True)
                if "ProductName" in df_stock.columns:
                    df_stock.rename(columns={"ProductName": "Product"}, inplace=True)
                
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
            
                # ---------------------------------------------------------
                # UI Filters (Year, Product, Stores)
                # ---------------------------------------------------------
                st.markdown("##### 🔍 Apply Filters")
                
                col1, col2 = st.columns(2)
                df_processed["Date"] = pd.to_datetime(df_processed["Date"], errors='coerce')
                df_processed["Year"] = df_processed["Date"].dt.year
                available_years = ["Overall"] + sorted(df_processed["Year"].dropna().astype(int).unique().tolist())
                
                with col1:
                    selected_year = st.selectbox("📅 Select Year", available_years)
                
                available_products = sorted(df_processed["Product"].dropna().unique().tolist())
                with col2:
                    selected_product = st.selectbox("👕 Select Product", available_products)
                
                col3, col4 = st.columns([3, 1])
                with col3:
                    available_stores = sorted(df_processed["Store Name"].dropna().unique().tolist())
                    target_defaults = [
                        "VIVIANA THANE STORE", "SEAWOOD GRAND CENTRAL MALL", 
                        "PANVEL ORION MALL", "AHMEDABAD ONE", 
                        "RM BORIVALI", "THE CAPITAL MALL VASAI", 
                        "RM VADODARA", "RM UDHANA"
                    ]
                    default_stores = [s for s in target_defaults if s in available_stores]
                    selected_stores = st.multiselect("🏬 Select Stores to Analyze", available_stores, default=default_stores)
                
                with col4:
                    st.markdown("<br>", unsafe_allow_html=True)
                    show_send_qty = st.checkbox("Show Send Qty & %", value=False)
                
                st.divider()

                # 3. Apply Filters to Sales & Stock
                all_product_sales = df_processed[df_processed["Product"] == selected_product].copy()
                filtered_sales = all_product_sales.copy()
                
                if selected_year != "Overall":
                    filtered_sales = filtered_sales[filtered_sales["Year"] == selected_year]
                if selected_stores:
                    filtered_sales = filtered_sales[filtered_sales["Store Name"].isin(selected_stores)]
                
                target_eans = all_product_sales["EANCode"].unique().tolist()
                filtered_stock = df_stock[df_stock["EANCode"].isin(target_eans)].copy()
                
                if selected_stores:
                    filtered_stock = filtered_stock[filtered_stock["Store Name"].isin(selected_stores)]
                
                # Standardize Color and Size strings to prevent split rows (e.g. 'Red' vs 'RED')
                for col in ["Color", "Size"]:
                    if col in filtered_sales.columns:
                        filtered_sales[col] = filtered_sales[col].astype(str).str.upper().str.strip()
                    if col in filtered_stock.columns:
                        filtered_stock[col] = filtered_stock[col].astype(str).str.upper().str.strip()

                # ---------------------------------------------------------
                # Master Aggregation Logic
                # ---------------------------------------------------------
                def generate_table(sales_df, stock_df, group_col):
                    if group_col not in sales_df.columns or group_col not in stock_df.columns:
                        return None
                    
                    s_agg = sales_df.groupby([group_col], as_index=False)["Quantity"].sum()
                    s_agg.rename(columns={"Quantity": "Sales Qty"}, inplace=True)
                    
                    k_agg = stock_df.groupby([group_col], as_index=False)["StockInHand"].sum()
                    
                    merged = pd.merge(s_agg, k_agg, on=[group_col], how="outer").fillna(0)
                    merged["Sales Qty"] = merged["Sales Qty"].astype(int)
                    merged["Current StoreStock"] = merged["StockInHand"].astype(int)
                    merged["Send Qty"] = merged["Sales Qty"] + merged["Current StoreStock"]
                    
                    merged["Percentage"] = (merged["Sales Qty"] / merged["Send Qty"].replace(0, 1)) * 100
                    merged["Percentage"] = merged["Percentage"].round(1).astype(str) + "%"
                    
                    display_cols = [group_col, "Sales Qty", "Send Qty", "Current StoreStock", "Percentage"]
                    final = merged[display_cols].sort_values(by="Sales Qty", ascending=False).reset_index(drop=True)
                    
                    if not show_send_qty:
                        final = final.drop(columns=["Send Qty", "Percentage"])
                    return final

                st.success(f"✅ Analytics for **{selected_product}** | Year: **{selected_year}** | Stores: **{len(selected_stores)} Selected**")

                # 4. Render 3 Separate Tables automatically
                st.markdown("#### 🏢 Store-wise Summary")
                df_store = generate_table(filtered_sales, filtered_stock, "Store Name")
                st.dataframe(df_store, use_container_width=True)
                
                col_left, col_right = st.columns(2)
                
                with col_left:
                    st.markdown("#### 🎨 Color-wise Summary")
                    df_color = generate_table(filtered_sales, filtered_stock, "Color")
                    if df_color is not None:
                        st.dataframe(df_color, use_container_width=True)
                    else:
                        st.warning("Color data missing.")
                        
                with col_right:
                    st.markdown("#### 📏 Size-wise Summary")
                    df_size = generate_table(filtered_sales, filtered_stock, "Size")
                    if df_size is not None:
                        st.dataframe(df_size, use_container_width=True)
                    else:
                        st.warning("Size data missing.")
                
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
