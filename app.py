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
    
    st.sidebar.header("Sheet Selection")
    main_sheet_name = st.sidebar.selectbox("Select the MAIN sheet (with colors)", sheet_names)
    delta_sheet_name = st.sidebar.selectbox("Select the DELTA sheet (with numbers)", sheet_names)
    
    num_benchmarks = st.sidebar.slider("Number of Benchmarks", 1, 5, 2)
    
    ws_main = wb[main_sheet_name]
    ws_delta = wb[delta_sheet_name]
    
    df_main = pd.DataFrame(ws_main.values)
    df_delta = pd.DataFrame(ws_delta.values)

    # Metadata capture from Main Sheet
    main_metadata = {}
    for r in range(1, ws_main.max_row + 1):
        for c in range(1, ws_main.max_column + 1):
            cell = ws_main.cell(row=r, column=c)
            meta = {"color": None, "is_percent": False}
            if cell.fill and cell.fill.start_color.rgb and cell.fill.start_color.rgb != "00000000" and cell.fill.start_color.rgb != "FFFFFFFF":
                color_hex = cell.fill.start_color.rgb
                meta["color"] = f"#{color_hex[2:]}" if len(color_hex) == 8 else f"#{color_hex}"
            if cell.number_format and '%' in cell.number_format:
                meta["is_percent"] = True
            if meta["color"] or meta["is_percent"]:
                main_metadata[(r-1, c-1)] = meta

    # --- 2. COLUMN MAPPING ---
    # Benchmark columns are always 2 per benchmark (Val, Sig)
    bench_cols_end = 2 + (num_benchmarks * 2)
    # Product columns in Delta sheet: 2 (Val, Sig) + num_benchmarks (Deltas)
    delta_cols_per_product = 2 + num_benchmarks
    # Product columns in Main sheet: 2 (Val, Sig)
    main_cols_per_product = 2

    products = []
    for col_idx in range(bench_cols_end, df_delta.shape[1], delta_cols_per_product):
        p_name = df_delta.iloc[2, col_idx]
        if pd.isna(p_name) or str(p_name).strip() == "": 
            p_name = f"Product {len(products)+1}"
        
        p_idx = len(products)
        products.append({
            "name": str(p_name).strip(),
            "delta_indices": list(range(col_idx, col_idx + delta_cols_per_product)),
            "main_indices": [bench_cols_end + (p_idx * 2), bench_cols_end + (p_idx * 2) + 1]
        })

    # --- 3. UI FILTERING (METRIC SELECTION) ---
    st.sidebar.header("Filters")
    regroup_mode = st.sidebar.toggle("Regroup by Question ID", value=True)
    
    selected_product_names = st.sidebar.multiselect("Select Products", [p['name'] for p in products], default=[p['name'] for p in products])
    final_products = [p for p in products if p['name'] in selected_product_names]

    st.header("🎯 Metric Selection")
    # Use df_delta as row reference
    raw_rows = df_delta.iloc[5:, [0, 1]].dropna(subset=[0])
    ui_q_map = {}
    for q_full in raw_rows[0].unique():
        display_name = get_question_root(q_full) if regroup_mode else q_full
        metrics = df_delta[df_delta[0] == q_full][1].unique().tolist()
        if display_name not in ui_q_map:
            ui_q_map[display_name] = {"originals": [q_full], "metrics": set(metrics)}
        else:
            ui_q_map[display_name]["originals"].append(q_full)
            ui_q_map[display_name]["metrics"].update(metrics)

    selected_q_metrics = {}
    for display_q, data in ui_q_map.items():
        with st.expander(f"❓ {display_q}"):
            metrics_list = sorted([str(m) for m in data["metrics"] if pd.notna(m)])
            gen_group = [m for m in metrics_list if m in GENERAL_METRICS]
            mod_group = [m for m in metrics_list if m not in GENERAL_METRICS]
            
            c1, c2, c3 = st.columns([1, 2, 2])
            is_active = c1.checkbox("Keep", value=True, key=f"act_{display_q}")
            
            if is_active:
                m_gen_key, m_mod_key = f"mg_{display_q}", f"mm_{display_q}"
                if m_gen_key not in st.session_state: st.session_state[m_gen_key] = gen_group
                if m_mod_key not in st.session_state: st.session_state[m_mod_key] = mod_group
                
                with c2:
                    st.write("General")
                    ca, cb = st.columns(2)
                    if ca.button("All", key=f"allg_{display_q}"): st.session_state[m_gen_key] = gen_group
                    if cb.button("None", key=f"clrg_{display_q}"): st.session_state[m_gen_key] = []
                    sel_gen = st.multiselect("Select", gen_group, key=m_gen_key)
                with c3:
                    st.write("Modalities")
                    ca, cb = st.columns(2)
                    if ca.button("All", key=f"allm_{display_q}"): st.session_state[m_mod_key] = mod_group
                    if cb.button("None", key=f"clrm_{display_q}"): st.session_state[m_mod_key] = []
                    sel_mod = st.multiselect("Select", mod_group, key=m_mod_key)
                
                for orig in data["originals"]:
                    selected_q_metrics[orig] = set(sel_gen + sel_mod)

    # --- 4. EXPORT ---
    if st.button("🚀 Generate Report"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Report')
            worksheet.freeze_panes(5, 2)
            
            # Formats
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#455A64', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            sub_header_fmt = workbook.add_format({'bold': True, 'bg_color': '#CFD8DC', 'border': 1, 'align': 'center'})
            pastel_fmts = [workbook.add_format({'bg_color': c, 'border': 1, 'valign': 'vcenter', 'text_wrap': True}) for c in PASTELS]

            # Build Header
            worksheet.write(2, 0, "Variable", header_fmt)
            worksheet.write(2, 1, "Metric", header_fmt)
            
            curr_col = 2
            # Benchmarks
            for i in range(num_benchmarks):
                label = df_delta.iloc[2, 2 + (i*2)]
                worksheet.merge_range(2, curr_col, 2, curr_col + 1, str(label), header_fmt)
                worksheet.write(4, curr_col, "Value", sub_header_fmt)
                worksheet.write(4, curr_col + 1, "Sig", sub_header_fmt)
                curr_col += 2
            
            # Products
            for p in final_products:
                n_cols = len(p['delta_indices'])
                worksheet.merge_range(2, curr_col, 2, curr_col + n_cols - 1, p['name'], header_fmt)
                for i, d_idx in enumerate(p['delta_indices']):
                    sub_label = df_delta.iloc[4, d_idx]
                    worksheet.write(4, curr_col + i, str(sub_label) if pd.notna(sub_label) else "", sub_header_fmt)
                curr_col += n_cols

            # Data Body
            out_row = 5
            last_root = None
            pastel_idx = -1
            
            for r_idx in range(5, len(df_delta)):
                q_val, m_val = str(df_delta.iloc[r_idx, 0]).strip(), str(df_delta.iloc[r_idx, 1]).strip()
                if q_val in selected_q_metrics and m_val in selected_q_metrics[q_val]:
                    # Question grouping & Pastel logic
                    root = get_question_root(q_val)
                    if root != last_root:
                        pastel_idx += 1
                        last_root = root
                    
                    row_fmt = pastel_fmts[pastel_idx % len(PASTELS)]
                    worksheet.write(out_row, 0, q_val, row_fmt)
                    worksheet.write(out_row, 1, m_val, workbook.add_format({'border': 1}))
                    
                    # Fill Benchmarks
                    for b_idx in range(num_benchmarks * 2):
                        c_idx = 2 + b_idx
                        val = df_delta.iloc[r_idx, c_idx]
                        meta = main_metadata.get((r_idx, c_idx), {"color": None, "is_percent": False})
                        fmt_dict = {'border': 1, 'align': 'center'}
                        if meta["color"]: fmt_dict['bg_color'] = meta["color"]
                        if meta["is_percent"]: fmt_dict['num_format'] = '0%'
                        worksheet.write(out_row, c_idx, val if pd.notna(val) else "", workbook.add_format(fmt_dict))
                    
                    # Fill Products
                    p_start_col = 2 + (num_benchmarks * 2)
                    for p in final_products:
                        # 1. Value Column
                        v_idx = p['delta_indices'][0]
                        v_main_idx = p['main_indices'][0]
                        v_meta = main_metadata.get((r_idx, v_main_idx), {"color": None, "is_percent": False})
                        v_fmt = {'border': 1, 'align': 'center', 'left': 2}
                        if v_meta["color"]: v_fmt['bg_color'] = v_meta["color"]
                        if v_meta["is_percent"]: v_fmt['num_format'] = '0%'
                        worksheet.write(out_row, p_start_col, df_delta.iloc[r_idx, v_idx] if pd.notna(df_delta.iloc[r_idx, v_idx]) else "", workbook.add_format(v_fmt))
                        
                        # 2. Sig Column (Colored, No Text)
                        s_main_idx = p['main_indices'][1]
                        s_meta = main_metadata.get((r_idx, s_main_idx), {"color": None, "is_percent": False})
                        s_fmt = {'border': 1}
                        if s_meta["color"]: s_fmt['bg_color'] = s_meta["color"]
                        worksheet.write(out_row, p_start_col + 1, "", workbook.add_format(s_fmt))
                        
                        # 3. Delta Columns
                        for i, d_idx in enumerate(p['delta_indices'][2:]):
                            d_val = df_delta.iloc[r_idx, d_idx]
                            d_fmt = {'border': 1, 'align': 'center', 'num_format': '0.00'}
                            if i == len(p['delta_indices'][2:]) - 1: d_fmt['right'] = 2
                            worksheet.write(out_row, p_start_col + 2 + i, d_val if pd.notna(d_val) else "", workbook.add_format(d_fmt))
                        
                        p_start_col += len(p['delta_indices'])
                    
                    out_row += 1

            worksheet.set_column(0, 0, 40)
            worksheet.set_column(1, 1, 20)

        st.success("✅ Done!")
        st.download_button("📥 Download Excel", output.getvalue(), "Market_Research_Report.xlsx")
