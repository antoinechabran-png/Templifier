import streamlit as st
import pandas as pd
import io
import re
import openpyxl
from openpyxl.utils import get_column_letter

# Set page config
st.set_page_config(page_title="Market Research Templifier Pro", layout="wide")

# Constants
GENERAL_METRICS = ["Mean", "Top Box", "Top 2 Boxes", "Top 3 Boxes", "Bottom Box", "Bottom 2 Boxes", "Bottom 3 Boxes"]
PASTELS = ['#E3F2FD', '#F3E5F5', '#E8F5E9', '#FFF3E0', '#FCE4EC'] # Blue, Purple, Green, Orange, Pink
BORDER_COLOR = '#B0BEC5' # Soft blue-grey

def get_question_root(q_text):
    match = re.match(r"^(Q-\d+|S-\d+)", str(q_text))
    return match.group(1) if match else str(q_text)

st.title("🎨 Market Research Templifier - Designer Edition")

uploaded_file = st.file_uploader("Upload Raw Results (Excel)", type=["xlsx"])

if uploaded_file is not None:
    # --- 1. SETTINGS & LOADING ---
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames
    selected_sheet = st.selectbox("Select the sheet to process", sheet_names)
    ws = wb[selected_sheet]
    
    num_benchmarks = st.sidebar.slider("Number of Benchmarks in Raw File", 1, 5, 2)
    bench_start_col = 2
    product_start_col = bench_start_col + (num_benchmarks * 2)
    cols_per_product = 2 + num_benchmarks

    df = pd.DataFrame(ws.values)
    df[0] = df[0].astype(str).str.strip()
    df[1] = df[1].astype(str).str.strip()

    # Metadata capture
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
            p_name = f"Product {get_column_letter(col_idx+1)}"
        product_triplets[str(p_name).strip()] = list(range(col_idx, col_idx + cols_per_product))

    selected_products = st.sidebar.multiselect("Select Products", list(product_triplets.keys()), default=list(product_triplets.keys()))
    show_sig = st.sidebar.checkbox("Show Significance Columns", value=True)

    # --- UI FILTERING (Simplified for brevity) ---
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    selected_q_metrics = {q: set(df[df[0] == q][1].unique()) for q in raw_data_area[0].unique() if q not in ["nan", "None", ""]}

    # --- 3. EXPORT WITH DESIGNER STYLING ---
    if st.button("🚀 Generate Beautiful Excel"):
        data_rows = []
        for idx, row in df.iloc[5:].iterrows():
            q_name, m_name = str(row[0]).strip(), str(row[1]).strip()
            if q_name in selected_q_metrics and m_name in selected_q_metrics[q_name]:
                data_rows.append((idx, row))

        if data_rows:
            cols_to_keep = [0, 1]
            benchmark_cols = list(range(bench_start_col, product_start_col))
            cols_to_keep.extend(benchmark_cols)
            
            final_product_cols = []
            for p in selected_products:
                indices = product_triplets[p]
                kept = indices if show_sig else indices[:2]
                final_product_cols.append({'name': p, 'indices': kept})
                cols_to_keep.extend(kept)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_rows = [df.iloc[i] for i in range(5)] + [r[1] for r in data_rows]
                final_df = pd.DataFrame(final_rows)[cols_to_keep].reset_index(drop=True)
                final_df.to_excel(writer, index=False, header=False, sheet_name='Report')
                
                workbook = writer.book
                worksheet = writer.sheets['Report']
                worksheet.freeze_panes(5, 2) # Freeze headers and first 2 columns

                # FORMATS
                header_base = {'bold': True, 'bg_color': '#455A64', 'font_color': 'white', 'border': 1, 'border_color': BORDER_COLOR, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Segoe UI'}
                header_fmt = workbook.add_format(header_base)
                
                # Dynamic Pastel Formats for Column A
                pastel_formats = [workbook.add_format({'bg_color': c, 'border': 1, 'border_color': BORDER_COLOR, 'valign': 'vcenter', 'text_wrap': True, 'font_name': 'Segoe UI', 'font_size': 9}) for c in PASTELS]

                # --- A. HEADERS ---
                for b_idx in range(num_benchmarks):
                    b_col = 2 + (b_idx * 2)
                    worksheet.merge_range(2, b_col, 2, b_col + 1, str(df.iloc[2, b_col]), header_fmt)

                curr_col = product_start_col
                for p_info in final_product_cols:
                    n_cols = len(p_info['indices'])
                    # Apply a specific format for product headers to give them a distinct outline
                    p_header_fmt = workbook.add_format({**header_base, 'left': 2, 'right': 2})
                    if n_cols > 1:
                        worksheet.merge_range(2, curr_col, 2, curr_col + n_cols - 1, p_info['name'], p_header_fmt)
                    else:
                        worksheet.write(2, curr_col, p_info['name'], p_header_fmt)
                    curr_col += n_cols

                # --- B. DATA & MERGING ---
                start_r = 5
                pastel_idx = 0
                for r in range(5, len(final_df)):
                    q_val = final_df.iloc[r, 0]
                    is_last_in_group = (r == len(final_df) - 1 or final_df.iloc[r+1, 0] != q_val)
                    
                    # 1. Pastel Merging for Column A
                    if is_last_in_group:
                        fmt = pastel_formats[pastel_idx % len(PASTELS)]
                        if r > start_r: worksheet.merge_range(start_r, 0, r, 0, q_val, fmt)
                        else: worksheet.write(start_r, 0, q_val, fmt)
                        pastel_idx += 1
                        start_r = r + 1

                    # 2. Row styling
                    orig_row_idx = data_rows[r - 5][0]
                    for target_c in range(1, len(final_df.columns)):
                        orig_col_idx = cols_to_keep[target_c]
                        val = final_df.iloc[r, target_c]
                        meta = cell_metadata.get((orig_row_idx, orig_col_idx), {"color": None, "is_percent": False})
                        
                        # Border logic: Dotted separator if not the last row of a group
                        border_style = 1 if is_last_in_group else 3 # 3 is 'dash-dot' or similar in xlsxwriter
                        
                        cell_style = {
                            'border': 1, 
                            'border_color': BORDER_COLOR,
                            'bottom': 1 if is_last_in_group else 4, # 4 is dotted
                            'font_name': 'Segoe UI',
                            'font_size': 9
                        }
                        
                        # Product block outlining (thick borders on left/right)
                        is_product_col = target_c >= product_start_col
                        if is_product_col:
                            # Find if this is start or end of a product block
                            block_rel_pos = (target_c - product_start_col) % (len(final_product_cols[0]['indices']))
                            if block_rel_pos == 0: cell_style['left'] = 2
                            if block_rel_pos == len(final_product_cols[0]['indices']) - 1: cell_style['right'] = 2

                        if meta["color"]: cell_style['bg_color'] = meta["color"]
                        if meta["is_percent"] or (str(final_df.iloc[r, 1]) not in GENERAL_METRICS and isinstance(val, (int, float))):
                            cell_style['num_format'] = '0%'
                        
                        worksheet.write(r, target_c, val if pd.notna(val) else "", workbook.add_format(cell_style))

                # Auto-adjust column widths
                worksheet.set_column(0, 0, 30)
                worksheet.set_column(1, 1, 20)
                worksheet.set_column(2, len(final_df.columns)-1, 10)

            st.success("✅ Designer Report Generated!")
            st.download_button("📥 Download Styled Excel", output.getvalue(), "Designer_Report.xlsx")
