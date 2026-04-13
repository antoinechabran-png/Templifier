import streamlit as st
import pandas as pd
import io
import re
import openpyxl
from openpyxl.utils import get_column_letter

# Set page config
st.set_page_config(page_title="Market Research Templifier", layout="wide")

# Constants
GENERAL_METRICS = ["Mean", "Top Box", "Top 2 Boxes", "Top 3 Boxes", "Bottom Box", "Bottom 2 Boxes", "Bottom 3 Boxes"]
PASTELS = ['#E3F2FD', '#F3E5F5', '#E8F5E9', '#FFF3E0', '#FCE4EC'] 
SOFT_BORDER = '#B0BEC5'

def get_question_root(q_text):
    match = re.match(r"^(Q-\d+|S-\d+)", str(q_text))
    return match.group(1) if match else str(q_text)

st.title("📊 Market Research Templifier")

uploaded_file = st.file_uploader("Upload Raw Results (Excel)", type=["xlsx"])

if uploaded_file is not None:
    # --- 1. SETTINGS & LOADING ---
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames
    selected_sheet = st.selectbox("Select the main sheet to process", sheet_names)
    ws = wb[selected_sheet]
    
    # Delta Integration Settings
    st.sidebar.header("Integration Settings")
    use_deltas = st.sidebar.toggle("Enable Delta Integration", value=False)
    df_delta = None
    delta_metadata = {}
    
    if use_deltas:
        delta_sheet_name = st.sidebar.selectbox("Select the Delta sheet", sheet_names)
        ws_delta = wb[delta_sheet_name]
        df_delta = pd.DataFrame(ws_delta.values)
        # Identify which cells in the delta sheet are formatted as percentages
        for r in range(1, ws_delta.max_row + 1):
            for c in range(1, ws_delta.max_column + 1):
                cell = ws_delta.cell(row=r, column=c)
                if cell.number_format and '%' in cell.number_format:
                    delta_metadata[(r-1, c-1)] = True

    # Sidebar Data Structure Settings
    num_benchmarks = st.sidebar.slider("Number of Benchmarks in Raw File", 1, 5, 2)
    bench_start_col = 2
    product_start_col = bench_start_col + (num_benchmarks * 2)
    
    # BASED ON IMAGE: A product block has: 1 (Value) + 1 (Sig Gap) + X (Deltas)
    # Looking at your provided CSV snippet: cols_per_product seems to be 2 + num_benchmarks
    cols_per_product = 2 + num_benchmarks

    # Load main data into DataFrame
    df = pd.DataFrame(ws.values)
    df[0] = df[0].astype(str).str.strip()
    df[1] = df[1].astype(str).str.strip()

    # Capture metadata (Color and Percentages)
    cell_metadata = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            meta = {"color": None, "is_percent": False}
            if cell.fill and cell.fill.start_color.index != "00000000" and cell.fill.start_color.rgb != "FFFFFFFF":
                color_hex = cell.fill.start_color.rgb
                if isinstance(color_hex, str) and len(color_hex) == 8:
                    meta["color"] = f"#{color_hex[2:]}"
            if cell.number_format and '%' in cell.number_format:
                meta["is_percent"] = True
            if meta["color"] or meta["is_percent"]:
                cell_metadata[(r-1, c-1)] = meta

    # --- 2. PARSING PRODUCTS ---
    product_triplets = {} 
    for col_idx in range(product_start_col, len(df.columns) - (cols_per_product - 1), cols_per_product):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name) or str(p_name).strip().lower() in ["none", "nan", ""]: 
            p_name = f"Product at Col {get_column_letter(col_idx+1)}"
        product_triplets[str(p_name).strip()] = list(range(col_idx, col_idx + cols_per_product))

    all_p_names = list(product_triplets.keys())
    selected_products = st.sidebar.multiselect("Select Products", all_p_names, default=all_p_names)
    show_sig = st.sidebar.checkbox("Show Sig/Delta Columns", value=True)

    # UI Question Selection
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    ui_q_map = {}
    for q_full in raw_data_area[0].unique():
        if q_full in ["nan", "None", ""]: continue
        display_name = get_question_root(q_full)
        metrics = df[df[0] == q_full][1].unique().tolist()
        if display_name not in ui_q_map:
            ui_q_map[display_name] = {"originals": [q_full], "metrics": set(metrics)}
        else:
            ui_q_map[display_name]["originals"].append(q_full)
            ui_q_map[display_name]["metrics"].update(metrics)

    selected_q_metrics = {}
    for display_q, data in ui_q_map.items():
        with st.expander(f"❓ {display_q}"):
            metrics_list = sorted([str(m) for m in data["metrics"] if pd.notna(m)])
            sel_m = st.multiselect("Pick Metrics", metrics_list, default=metrics_list, key=f"m_{display_q}")
            for orig in data["originals"]:
                selected_q_metrics[orig] = set(sel_m)

    # --- 3. EXPORT GENERATION ---
    if st.button("🚀 Generate Beautiful Excel"):
        data_rows_indices = []
        for idx, row in df.iloc[5:].iterrows():
            q_name, m_name = str(row[0]).strip(), str(row[1]).strip()
            if q_name in selected_q_metrics and m_name in selected_q_metrics[q_name]:
                data_rows_indices.append(idx)

        if data_rows_indices:
            cols_to_keep = [0, 1]
            benchmark_cols = list(range(bench_start_col, product_start_col))
            cols_to_keep.extend(benchmark_cols)
            
            final_product_cols = []
            for p in selected_products:
                indices = product_triplets[p]
                kept = indices if show_sig else indices[:1]
                final_product_cols.append({'name': p, 'indices': kept})
                cols_to_keep.extend(kept)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Build the rows
                header_rows = [df.iloc[i].copy() for i in range(5)]
                body_rows = [df.iloc[i].copy() for i in data_rows_indices]
                final_rows = header_rows + body_rows
                
                # Apply Delta Logic (Transform % to whole number)
                for r_idx, row in enumerate(final_rows):
                    orig_r = r_idx if r_idx < 5 else data_rows_indices[r_idx - 5]
                    
                    for p_info in final_product_cols:
                        # Delta cols start at index 2 of the product block
                        delta_indices = p_info['indices'][2:]
                        for c_idx in delta_indices:
                            # Pull from Delta sheet if requested, otherwise use main
                            source_df = df_delta if (use_deltas and df_delta is not None) else df
                            val = source_df.iloc[orig_r, c_idx]
                            
                            # If it's a numeric percentage, multiply by 100
                            if isinstance(val, (int, float)):
                                # If it looks like a decimal percentage or is marked as percent in metadata
                                if delta_metadata.get((orig_r, c_idx), False) or (abs(val) <= 1.0 and val != 0):
                                    row.iloc[c_idx] = val * 100
                                else:
                                    row.iloc[c_idx] = val

                final_df = pd.DataFrame(final_rows)[cols_to_keep].reset_index(drop=True)
                final_df.to_excel(writer, index=False, header=False, sheet_name='Report')
                
                workbook = writer.book
                worksheet = writer.sheets['Report']
                
                # Formatting
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#455A64', 'font_color': 'white', 'border': 1, 'align': 'center'})
                sub_header_fmt = workbook.add_format({'bold': True, 'bg_color': '#CFD8DC', 'border': 1, 'align': 'center'})
                
                # Header Merging
                curr_c = product_start_col
                for p_info in final_product_cols:
                    n = len(p_info['indices'])
                    if n > 1:
                        worksheet.merge_range(2, curr_c, 2, curr_c + n - 1, p_info['name'], header_fmt)
                    else:
                        worksheet.write(2, curr_c, p_info['name'], header_fmt)
                    curr_c += n

                # Body Styling
                for r in range(5, len(final_df)):
                    orig_r = data_rows_indices[r-5]
                    for target_c, orig_c in enumerate(cols_to_keep):
                        if target_c < 2: continue
                        
                        val = final_df.iloc[r, target_c]
                        meta = cell_metadata.get((orig_r, orig_c), {"color": None, "is_percent": False})
                        
                        style = {'border': 1, 'border_color': SOFT_BORDER, 'align': 'center'}
                        if meta["color"]: style['bg_color'] = meta["color"]
                        
                        # Determine if this target_c is a Delta column
                        # Logic: Find which product block it belongs to and check its relative position
                        is_delta = False
                        relative_ptr = product_start_col
                        for p_info in final_product_cols:
                            block_len = len(p_info['indices'])
                            if relative_ptr <= target_c < relative_ptr + block_len:
                                if (target_c - relative_ptr) >= 2: is_delta = True
                                break
                            relative_ptr += block_len

                        if is_delta:
                            style['num_format'] = '0.00'
                        elif meta["is_percent"]:
                            style['num_format'] = '0%'
                        
                        worksheet.write(r, target_c, val if pd.notna(val) else "", workbook.add_format(style))

                worksheet.set_column(0, 0, 40)
                worksheet.freeze_panes(5, 2)

            st.success("✅ Delta integration complete.")
            st.download_button("📥 Download Report", output.getvalue(), "Market_Report.xlsx")
