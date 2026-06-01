import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Model Report Workspace", layout="wide")
st.title("📊 Model Report Workspace")

# --- Workflow selection at the top ---
analysis_mode = st.radio(
    "Choose Analysis Mode:",
    options=["Single Report Mode", "Two-Report Comparison Mode"],
    horizontal=True,
    help="Switch between examining a single file or running a side-by-side comparison between two months."
)

# --- CACHED FUNCTIONS FOR LIGHTNING PERFORMANCE ---
@st.cache_data(show_spinner="Reading Excel structure...")
def get_sheets_single(file):
    xl = pd.ExcelFile(file)
    return sorted(xl.sheet_names)

@st.cache_data(show_spinner="Reading Excel structures...")
def get_common_sheets(file_old, file_new):
    xl_old = pd.ExcelFile(file_old)
    xl_new = pd.ExcelFile(file_new)
    return sorted(list(set(xl_old.sheet_names) & set(xl_new.sheet_names)))

@st.cache_data(show_spinner="Loading data into memory...")
def load_sheet_single(file, sheet_name):
    return pd.read_excel(file, sheet_name=sheet_name)

@st.cache_data(show_spinner="Loading data into memory...")
def load_sheet_comparison(file_old, file_new, sheet_name):
    df_old = pd.read_excel(file_old, sheet_name=sheet_name)
    df_new = pd.read_excel(file_new, sheet_name=sheet_name)
    return df_old, df_new

# --- VIBRANT COLOR MAPPER FOR PLOTS ---
def get_metric_colors(metric_name):
    m_lower = metric_name.lower()
    if "actual" in m_lower:
        return "#0072B2", "#56B4E9"  # Deep Vibrant Blue / Sky Blue
    elif "lasso" in m_lower:
        return "#D55E00", "#FFB591"  # Sharp Electric Orange / Soft Peach
    elif "sp" in m_lower or "scenario_qty_sp" in m_lower:
        return "#009E73", "#89E5C9"  # Emerald Green / Mint Green
    elif "wmape" in m_lower:
        return "#CC79A7", "#F0D0E0"  # Hot Magenta / Light Pink
    else:
        return "#F0E442", "#FFF79A"  # Neon Yellow / Soft Yellow


# ==========================================
#  FLOW A: SINGLE REPORT MODE
# ==========================================
if analysis_mode == "Single Report Mode":
    uploaded_file = st.file_uploader("📁 Upload Excel Report File", type=["xlsx", "xls"])
    
    if uploaded_file:
        sheet_list = get_sheets_single(uploaded_file)
        selected_sheet = st.sidebar.selectbox("🎯 Select Sheet to Analyze", sheet_list)
        st.subheader(f"Examining Sheet: `{selected_sheet}`")
        
        df_raw = load_sheet_single(uploaded_file, selected_sheet)
        df = df_raw.copy()
        all_cols = list(df.columns)
        
        # --- CASE A1: MODEL SHEETS (Time Series) ---
        if "_model" in selected_sheet.lower():
            date_col = next((c for c in all_cols if 'dt' in c.lower() or 'date' in c.lower() or pd.api.types.is_datetime64_any_dtype(df[c])), all_cols[0])
            numeric_cols = [c for c in all_cols if pd.api.types.is_numeric_dtype(df[c]) and c != date_col]
            
            st.sidebar.markdown("### 🛠️ Add Custom Chart Filters")
            filter_targets = st.sidebar.multiselect("Select columns to filter by:", options=all_cols, key="s_mod_targ", placeholder="Choose columns")
            
            filtered_df = df.copy()
            for col in filter_targets:
                combined_vals = sorted(list(df[col].dropna().astype(str).unique()))
                selected_vals = st.sidebar.multiselect(f"Values for [{col}]", options=combined_vals, key=f"s_mod_val_{col}", placeholder=f"All {col} values active")
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]
                    
            selected_metrics = st.multiselect("📊 Select Metrics to View on Plot", numeric_cols, default=numeric_cols[:3] if len(numeric_cols) >= 3 else numeric_cols)
            
            if not filtered_df.empty and selected_metrics:
                filtered_df[date_col] = filtered_df[date_col].astype(str)
                agg_df = filtered_df.groupby(date_col)[selected_metrics].sum().reset_index().sort_values(by=date_col)
                
                fig = go.Figure()
                for metric in selected_metrics:
                    curr_color, _ = get_metric_colors(metric)
                    fig.add_trace(go.Scatter(x=agg_df[date_col], y=agg_df[metric], mode='lines+markers', name=metric, line=dict(color=curr_color, width=3), marker=dict(size=6)))
                
                fig.update_layout(title=f"Trend Analysis Over Time ({date_col})", xaxis_title=date_col, yaxis_title="Sum of Quantities", hovermode="x unified", plot_bgcolor="white", xaxis=dict(showgrid=True, gridcolor="#EAEAEA"), yaxis=dict(showgrid=True, gridcolor="#EAEAEA"))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("⚠️ No data available to draw timeline plot based on active selections.")
                
        # --- CASE A2: METRIC SHEETS ---
        elif "_metric" in selected_sheet.lower():
            metric_view_layout = st.radio(
                "Select Metric View Layout:",
                options=["Pivot Table Matrix Explorer", "Raw Table Cell Viewer (e.g., WMAPEs)"],
                horizontal=True,
                key="s_metric_layout"
            )
            
            if metric_view_layout == "Pivot Table Matrix Explorer":
                categorical_cols = [c for c in all_cols if df[c].dtype == 'object' or isinstance(df[c].dtype, pd.CategoricalDtype)]
                if not categorical_cols: categorical_cols = all_cols
                
                st.sidebar.markdown("### 🛠️ Add Data Filters for Pivot")
                filter_targets = st.sidebar.multiselect("Select columns to filter data by:", options=all_cols, key="s_met_targ", placeholder="Choose columns")
                
                filtered_df = df.copy()
                for col in filter_targets:
                    combined_vals = sorted(list(df[col].dropna().astype(str).unique()))
                    selected_vals = st.sidebar.multiselect(f"Values for [{col}]", options=combined_vals, key=f"s_met_val_{col}", placeholder=f"All {col} active")
                    if selected_vals:
                        filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]
                
                st.markdown("### 🎛️ Setup Pivot Configurations")
                p_col1, p_col2, p_col3, p_col4 = st.columns(4)
                with p_col1:
                    row_sel = st.selectbox("Rows (Categorical Only)", categorical_cols, index=0, key="s_row")
                with p_col2:
                    remaining_cat = [c for c in categorical_cols if c != row_sel] or categorical_cols
                    bin_idx = next((i for i, c in enumerate(remaining_cat) if 'bin' in c.lower() or 'label' in c.lower() or 'scenario' in c.lower()), 0)
                    col_sel = st.selectbox("Columns (Bins / Labels Only)", remaining_cat, index=bin_idx, key="s_col")
                with p_col3:
                    val_idx = next((i for i, c in enumerate(all_cols) if 'itm' in c.lower() or 'id' in c.lower()), len(all_cols)-1)
                    val_sel = st.selectbox("Values to Aggregate", all_cols, index=val_idx, key="s_val")
                with p_col4:
                    agg_func = st.selectbox("Aggregation Function", ["count", "sum", "mean"], index=0, key="s_func")
                    
                try:
                    if filtered_df.empty:
                        st.warning("⚠️ No data matches current filters.")
                    else:
                        pivot_matrix = filtered_df.pivot_table(index=row_sel, columns=col_sel, values=val_sel, aggfunc=agg_func, fill_value=0)
                        pivot_matrix['Grand Total'] = pivot_matrix.sum(axis=1)
                        pivot_matrix.loc['Grand Total'] = pivot_matrix.sum(axis=0)
                        
                        st.markdown("### 🗺️ Distribution Breakdown Matrix")
                        st.dataframe(pivot_matrix.style.background_gradient(cmap="Blues", axis=None), use_container_width=True)
                except Exception as e:
                    st.error(f"Could not calculate pivot grid: {e}")
            
            else:
                default_keys = [c for c in all_cols if any(x in c.lower() for x in ['rfm', 'itm', 'id', 'menu']) and not any(m in c.lower() for m in ['wmape', 'actual', 'qty'])]
                
                id_cols = st.multiselect(
                    "Select Unique Row Key Column(s):", 
                    options=all_cols, 
                    default=default_keys if default_keys else [all_cols[0]] if all_cols else None,
                    key="s_raw_id"
                )
                
                if not id_cols:
                    st.warning("⚠️ Please choose at least one key column to structure rows.")
                else:
                    df_display = df.set_index(id_cols).sort_index()
                    if 'Report_Version' in df_display.columns:
                        df_display = df_display.drop(columns=['Report_Version'])
                    
                    format_rules = {}
                    for col in df_display.columns:
                        if 'wmape' in col.lower() or 'error' in col.lower():
                            format_rules[col] = lambda x: f"{x*100:.2f}%" if pd.notnull(x) else "-"
                        else:
                            format_rules[col] = lambda x: f"{x:,.2f}" if pd.notnull(x) else "-"
                            
                    st.markdown("### 📋 Raw Calculated Metric Table")
                    st.dataframe(df_display.style.format(format_rules), use_container_width=True)


# ==========================================
#  FLOW B: TWO-REPORT COMPARISON MODE
# ==========================================
elif analysis_mode == "Two-Report Comparison Mode":
    col1, col2 = st.columns(2)
    with col1:
        uploaded_old = st.file_uploader("📁 Upload PREVIOUS Month Excel (Old)", type=["xlsx", "xls"])
    with col2:
        uploaded_new = st.file_uploader("📁 Upload CURRENT Month Excel (New)", type=["xlsx", "xls"])

    if uploaded_old and uploaded_new:
        common_sheets = get_common_sheets(uploaded_old, uploaded_new)
        
        if not common_sheets:
            st.error("No common sheet names found between the two files!")
        else:
            selected_sheet = st.sidebar.selectbox("🎯 Select Sheet to Analyze", common_sheets)
            st.subheader(f"Comparing Sheet: `{selected_sheet}`")
            
            df_old_raw, df_new_raw = load_sheet_comparison(uploaded_old, uploaded_new, selected_sheet)
            df_old = df_old_raw.copy()
            df_new = df_new_raw.copy()
            
            df_old['Report_Version'] = 'Previous Month'
            df_new['Report_Version'] = 'Current Month'
            all_cols = [c for c in df_new.columns if c != 'Report_Version']
            
            # --- CASE B1: MODEL SHEETS ---
            if "_model" in selected_sheet.lower():
                date_col = next((c for c in all_cols if 'dt' in c.lower() or 'date' in c.lower() or pd.api.types.is_datetime64_any_dtype(df_new[c])), all_cols[0])
                numeric_cols = [c for c in all_cols if pd.api.types.is_numeric_dtype(df_new[c]) and c != date_col]
                
                st.sidebar.markdown("### 🛠️ Add Custom Chart Filters")
                filter_targets = st.sidebar.multiselect("Select columns to filter by:", options=all_cols, key="c_mod_targ", placeholder="Choose columns")
                
                filtered_df_old = df_old.copy()
                filtered_df_new = df_new.copy()
                
                for col in filter_targets:
                    combined_vals = sorted(list(set(df_old[col].dropna().astype(str).unique()) | set(df_new[col].dropna().astype(str).unique())))
                    selected_vals = st.sidebar.multiselect(f"Values for [{col}]", options=combined_vals, key=f"c_mod_val_{col}", placeholder=f"All {col} active")
                    if selected_vals:
                        filtered_df_old = filtered_df_old[filtered_df_old[col].astype(str).isin(selected_vals)]
                        filtered_df_new = filtered_df_new[filtered_df_new[col].astype(str).isin(selected_vals)]
                
                selected_metrics = st.multiselect("📊 Select Metrics to Compare on Plot", numeric_cols, default=numeric_cols[:3] if len(numeric_cols) >= 3 else numeric_cols)
                
                combined_df = pd.concat([filtered_df_old, filtered_df_new], ignore_index=True)
                
                if not combined_df.empty and selected_metrics:
                    combined_df[date_col] = combined_df[date_col].astype(str)
                    groupby_cols = [date_col, 'Report_Version']
                    agg_df = combined_df.groupby(groupby_cols)[selected_metrics].sum().reset_index().sort_values(by=date_col)
                    
                    fig = go.Figure()
                    for metric in selected_metrics:
                        curr_color, prev_color = get_metric_colors(metric)
                        
                        # Previous Month (Dotted tint line)
                        old_data = agg_df[agg_df['Report_Version'] == 'Previous Month']
                        fig.add_trace(go.Scatter(x=old_data[date_col], y=old_data[metric], mode='lines+markers', name=f"⏮️ {metric} (Prev Month)", line=dict(color=prev_color, dash='dot', width=2), marker=dict(size=5)))
                        
                        # Current Month (Solid bold line)
                        new_data = agg_df[agg_df['Report_Version'] == 'Current Month']
                        fig.add_trace(go.Scatter(x=new_data[date_col], y=new_data[metric], mode='lines+markers', name=f"🆕 {metric} (Curr Month)", line=dict(color=curr_color, width=3.5), marker=dict(size=7)))
                    
                    fig.update_layout(title=f"Comparison over Time ({date_col})", xaxis_title=date_col, yaxis_title="Sum of Quantities", hovermode="x unified", plot_bgcolor="white", xaxis=dict(showgrid=True, gridcolor="#EAEAEA"), yaxis=dict(showgrid=True, gridcolor="#EAEAEA"))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("⚠️ No data matches filters.")
                    
            # --- CASE B2: METRIC SHEETS ---
            elif "_metric" in selected_sheet.lower():
                metric_view_layout = st.radio(
                    "Select Metric View Layout:",
                    options=["Pivot Table Matrix Explorer", "Unified Interleaved Table Cell Comparison"],
                    horizontal=True,
                    key="c_metric_layout"
                )
                
                # --- SUB-ROUTE 1: PIVOT SETUP ---
                if metric_view_layout == "Pivot Table Matrix Explorer":
                    categorical_cols = [c for c in all_cols if df_new[c].dtype == 'object' or isinstance(df_new[c].dtype, pd.CategoricalDtype)]
                    if not categorical_cols: categorical_cols = all_cols
                    
                    st.sidebar.markdown("### 🛠️ Add Data Filters for Pivots")
                    filter_targets = st.sidebar.multiselect("Select columns to filter matrices by:", options=all_cols, key="c_met_targ", placeholder="Choose columns")
                    
                    filtered_df_old = df_old.copy()
                    filtered_df_new = df_new.copy()
                    
                    for col in filter_targets:
                        combined_vals = sorted(list(set(df_old[col].dropna().astype(str).unique()) | set(df_new[col].dropna().astype(str).unique())))
                        selected_vals = st.sidebar.multiselect(f"Values for [{col}]", options=combined_vals, key=f"c_met_val_{col}", placeholder=f"All {col} active")
                        if selected_vals:
                            filtered_df_old = filtered_df_old[filtered_df_old[col].astype(str).isin(selected_vals)]
                            filtered_df_new = filtered_df_new[filtered_df_new[col].astype(str).isin(selected_vals)]
                    
                    st.markdown("### 🎛️ Setup Pivot Configurations")
                    p_col1, p_col2, p_col3, p_col4 = st.columns(4)
                    with p_col1:
                        row_sel = st.selectbox("Rows (Categorical Only)", categorical_cols, index=0, key="c_row")
                    with p_col2:
                        remaining_cat_cols = [c for c in categorical_cols if c != row_sel] or categorical_cols
                        bin_idx = next((i for i, c in enumerate(remaining_cat_cols) if 'bin' in c.lower() or 'label' in c.lower() or 'scenario' in c.lower()), 0)
                        col_sel = st.selectbox("Columns (Bins / Labels Only)", remaining_cat_cols, index=bin_idx, key="c_col")
                    with p_col3:
                        val_idx = next((i for i, c in enumerate(all_cols) if 'itm' in c.lower() or 'id' in c.lower()), len(all_cols)-1)
                        val_sel = st.selectbox("Values to Aggregate", all_cols, index=val_idx, key="c_val")
                    with p_col4:
                        agg_func = st.selectbox("Aggregation Function", ["count", "sum", "mean"], index=0, key="c_func")
                        
                    try:
                        if filtered_df_old.empty or filtered_df_new.empty:
                            st.warning("⚠️ Data filters left one or both months empty.")
                        else:
                            def make_pivot(df):
                                pivot = df.pivot_table(index=row_sel, columns=col_sel, values=val_sel, aggfunc=agg_func, fill_value=0)
                                pivot['Grand Total'] = pivot.sum(axis=1)
                                pivot.loc['Grand Total'] = pivot.sum(axis=0)
                                return pivot

                            pivot_old = make_pivot(filtered_df_old)
                            pivot_new = make_pivot(filtered_df_new)
                            
                            pivot_old_aligned, pivot_new_aligned = pivot_old.align(pivot_new, fill_value=0)
                            delta_pivot = pivot_new_aligned - pivot_old_aligned
                            
                            st.markdown("---")
                            st.markdown("### 🗺️ Side-by-Side Monthly Distribution Matrix")
                            m_col1, m_col2 = st.columns(2)
                            with m_col1:
                                st.subheader("⏮️ Previous Month Pivot")
                                st.dataframe(pivot_old.style.background_gradient(cmap="Blues", axis=None), use_container_width=True)
                            with m_col2:
                                st.subheader("🆕 Current Month Pivot")
                                st.dataframe(pivot_new.style.background_gradient(cmap="Greens", axis=None), use_container_width=True)
                            
                            st.markdown("---")
                            st.markdown("### 🔄 Bin Shifting & Delta Analysis (Current - Previous)")
                            st.dataframe(delta_pivot.style.format("{:+}").background_gradient(cmap="RdYlGn", axis=None), use_container_width=True)
                            
                    except Exception as e:
                        st.error(f"Could not compute pivot matrix summaries: {e}")
                
                # --- FIXED ROUTE: UNIFIED INTERLEAVED COMPONENT ---
                else:
                    default_keys = [c for c in all_cols if any(x in c.lower() for x in ['rfm', 'itm', 'id', 'menu']) and not any(m in c.lower() for m in ['wmape', 'actual', 'qty'])]
                    
                    id_cols = st.multiselect(
                        "Select Grain/Key columns to merge data on:", 
                        options=all_cols, 
                        default=default_keys if default_keys else [all_cols[0]] if all_cols else None,
                        key="c_raw_id"
                    )
                    
                    if not id_cols:
                        st.warning("⚠️ Please select grain rows to cross-align tables.")
                    else:
                        st.sidebar.markdown("### 🗺️ Slice Metrics Columns")
                        numeric_cols = [c for c in all_cols if pd.api.types.is_numeric_dtype(df_new[c]) and c not in id_cols]
                        selected_cols = st.sidebar.multiselect("Visible Metric Columns:", numeric_cols, default=numeric_cols[:4] if len(numeric_cols) >= 4 else numeric_cols)
                        
                        if not selected_cols:
                            st.info("Select metrics on the sidebar selector to build the comparative table layout.")
                        else:
                            # Isolate indices
                            raw_old_indexed = df_old.set_index(id_cols)
                            raw_new_indexed = df_new.set_index(id_cols)
                            
                            # Construct unified structure
                            combined_index = raw_old_indexed.index.union(raw_new_indexed.index)
                            unified_df = pd.DataFrame(index=combined_index)
                            
                            interleaved_columns_list = []
                            formatting_rules_map = {}
                            delta_wmape_cols = []
                            
                            for metric in selected_cols:
                                c_old = f"{metric} (Old)"
                                c_new = f"{metric} (New)"
                                c_delta = f"{metric} (Delta)"
                                
                                unified_df[c_old] = pd.to_numeric(raw_old_indexed[metric] if metric in raw_old_indexed.columns else pd.NA, errors='coerce')
                                unified_df[c_new] = pd.to_numeric(raw_new_indexed[metric] if metric in raw_new_indexed.columns else pd.NA, errors='coerce')
                                unified_df[c_delta] = unified_df[c_new] - unified_df[c_old]
                                
                                interleaved_columns_list.extend([c_old, c_new, c_delta])
                                
                                # Apply formatting logic
                                if 'wmape' in metric.lower() or 'error' in metric.lower():
                                    formatting_rules_map[c_old] = lambda x: f"{x*100:.2f}%" if pd.notnull(x) else "-"
                                    formatting_rules_map[c_new] = lambda x: f"{x*100:.2f}%" if pd.notnull(x) else "-"
                                    formatting_rules_map[c_delta] = lambda x: f"{x*100:+.2f}%" if pd.notnull(x) else "0.00%"
                                    delta_wmape_cols.append(c_delta)
                                else:
                                    formatting_rules_map[c_old] = lambda x: f"{x:,.2f}" if pd.notnull(x) else "-"
                                    formatting_rules_map[c_new] = lambda x: f"{x:,.2f}" if pd.notnull(x) else "-"
                                    formatting_rules_map[c_delta] = lambda x: f"{x:+,.2f}" if pd.notnull(x) else "0.00"
                            
                            final_render_df = unified_df[interleaved_columns_list].sort_index()
                            
                            st.markdown("### 📋 Unified Interleaved Grain Comparison View")
                            st.write("Each chosen metric is evaluated side-by-side. For **WMAPE Delta** columns, **Negative values (Green)** indicate accuracy improvements.")
                            
                            # FIXED: Changed columns= to subset= inside background_gradient
                            styled_matrix = final_render_df.style.format(formatting_rules_map).background_gradient(
                                cmap="RdYlGn_r", 
                                subset=delta_wmape_cols, 
                                axis=0
                            )
                            
                            st.dataframe(styled_matrix, use_container_width=True)
else:
    st.info("Please upload both files at the top to initialize the visualization suite.")
