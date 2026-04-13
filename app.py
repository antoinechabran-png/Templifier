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

def get_question_root(q_text):
    match = re.match(r"^(Q-\d+|S-\d+)", str(q_text))
    return match.group(1) if match else str(q_text)

st.title("📊 Market Research Templifier")

uploaded_file = st.file_uploader("Upload Raw Results (Excel)", type=["xlsx"])

if uploaded_file is not None:
    # --- 1. DATA & COLOR LOADING ---
    # We use openpyxl to get colors, and pandas for the data
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames
    selected_sheet = st.selectbox("Select the sheet to process", sheet_names)
    ws = wb[selected_sheet]
    
    # Load data into pandas for easier manipulation
    df = pd.DataFrame(ws.values)
    df[0] = df[0].astype(str).str.strip()
    df[1] = df[1].astype(str).str.strip()

    # Create a color map for Significance columns (C, G, J, M...)
    # We store by (row_index, col_index)
    color_map = {}
    for r in range(6, ws.max_row + 1):
        for c in range(3, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            # Check if cell has a fill color (ignoring default white/none)
            if cell.fill and cell.fill.start_color.index != "00000000" and cell.fill.start_color.rgb != "FFFFFFFF":
                color_hex = cell.fill.start_color.rgb
                if isinstance(color_hex, str) and len(color_hex) == 8: # ARGB
                    color_map[(r-1, c-1)] = f"#{color_hex[2:]}" # Convert to #RRGGBB

    # --- 2. PARSING PRODUCTS ---
    product_triplets = {}
    for col_idx in range(4, len(df.columns) - 2, 3):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name): p_name = f"Product at Column {col_idx+1}"
        product_triplets[p_name] = [col_idx, col_idx + 1, col_idx + 2]

    # --- SIDEBAR ---
    regroup_mode = st.sidebar.toggle("Enable Question Regrouping", value=True)
    all_p_names = list(product_triplets.keys())
    selected_products = st.sidebar.multiselect("Select Products", all_p_names, default=all_p_names)
    show_sig = st.sidebar.checkbox("Show Significance Columns", value=True)

    # --- UI FILTERING LOGIC ---
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    ui_q_map = {}
    for q_full in raw_data_area[0].unique():
        if q_full in ["nan", "None", ""]: continue
        display_name = get_question_root(q_full) if regroup_mode else q_full
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
            gen_group = [m for m in metrics_list if m in GENERAL_METRICS]
            mod_group = [m for m in metrics_list if m not in GENERAL_METRICS]
            
            c1, c2, c3 = st.columns([1, 2, 2])
            is_active = c1.checkbox("Keep", value=True, key=f"act_{display_q}")
            
            if is_active:
                m_gen_k, m_mod_k = f"mg_{display_q}", f"mm_{display_q}"
                if m_gen_k not in st.session_state: st.session_state[m_gen_k] = gen_group
                if m_mod_k not in st.session_state: st.session_state[m_mod_k] = mod_group

                with c2:
                    st.write("General")
                    if st.button("Select All", key=f"allg_{display_q}"): st.session_state[m_gen_k] = gen_group
                    if st.button("Clear", key=f"clrg_{display_q}"): st.session_state[m_gen_k] = []
                    sel_gen = st.multiselect("Pick", gen_group, key=m_gen_k)
                with c3:
                    st.write("Modalities")
                    if st.button("Select All", key=f"allm_{display_q}"): st.session_state[m_mod_k] = mod_group
                    if st.button("Clear", key=f"clrm_{display_q}"): st.session_state[m_mod_k] = []
                    sel_mod = st.multiselect("Pick", mod_group, key=m_mod_k)
                
                for orig in data["originals"]:
                    selected_q_metrics[orig] = set(sel_gen + sel_mod)

    # --- 3. EXPORT ---
    if st.button("🚀 Generate Templated Excel"):
        data_rows = []
        # Keep track of the original row index to fetch the color later
        for idx, row in df.iloc[5:].iterrows():
            q_name, m_name = str(row[0]).strip(), str(row[1]).strip()
            if q_name in selected_q_metrics and m_name in selected_q_metrics[q_name]:
                data_rows.append((idx, row))

        if not data_rows:
            st.error("No data selected.")
        else:
            # Columns to keep
            cols_to_keep = [0, 1, 2]
            if show_sig: cols_to_keep.append(3)
            
            # Map products
            final_product_cols = []
            for p in selected_products:
                indices = product_triplets[p]
                if show_sig:
                    final_product_cols.append({'name': p, 'indices': indices})
                    cols_to_keep.extend(indices)
                else:
                    final_product_cols.append({'name': p, 'indices': [indices[0], indices[2]]})
                    cols_to_keep.extend([indices[0], indices[2]])

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Pre-process final dataframe
                final_rows = [df.iloc[i] for i in range(5)] + [r[1] for r in data_rows]
                final_df = pd.DataFrame(final_rows)[cols_to_keep].reset_index(drop=True)
                final_df.to_excel(writer, index=False, header=False, sheet_name='Report')
                
                workbook = writer.book
                worksheet = writer.sheets['Report']

                # Format definitions
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
                percent_fmt = workbook.add_format({'num_format': '0%', 'border': 1})
                std_fmt = workbook.add_format({'border': 1})
                merge_fmt = workbook.add_format({'valign': 'vcenter', 'align': 'left', 'border': 1, 'text_wrap': True})

                # A. Merge Product Names in Row 3
                current_col = 4 if show_sig else 3
                for p_info in final_product_cols:
                    num_cols = len(p_info['indices'])
                    if num_cols > 1:
                        worksheet.merge_range(2, current_col, 2, current_col + num_cols - 1, p_info['name'], header_fmt)
                    else:
                        worksheet.write(2, current_col, p_info['name'], header_fmt)
                    current_col += num_cols

                # B. Merge Question Names (Column A)
                start_r = 5
                for r in range(6, len(final_df)):
                    if final_df.iloc[r, 0] != final_df.iloc[start_r, 0]:
                        if r - 1 > start_r: worksheet.merge_range(start_r, 0, r - 1, 0, final_df.iloc[start_r, 0], merge_fmt)
                        else: worksheet.write(start_r, 0, final_df.iloc[start_r, 0], merge_fmt)
                        start_r = r
                worksheet.merge_range(start_r, 0, len(final_df)-1, 0, final_df.iloc[start_r, 0], merge_fmt)

                # C. Apply Data, Colors, and Percentages
                for target_r in range(5, len(final_df)):
                    orig_row_idx = data_rows[target_r - 5][0]
                    m_type = str(final_df.iloc[target_r, 1])
                    
                    for target_c in range(1, len(final_df.columns)):
                        orig_col_idx = cols_to_keep[target_c]
                        val = final_df.iloc[target_r, target_c]
                        
                        # Determine base format
                        fmt = percent_fmt if (m_type not in GENERAL_METRICS and isinstance(val, (int, float))) else std_fmt
                        
                        # Apply Color if exists in original color_map
                        if (orig_row_idx, orig_col_idx) in color_map:
                            hex_c = color_map[(orig_row_idx, orig_col_idx)]
                            # Create a unique format for this specific color
                            colored_fmt = workbook.add_format({'border': 1, 'bg_color': hex_c})
                            if fmt == percent_fmt: colored_fmt.set_num_format('0%')
                            worksheet.write(target_r, target_c, val if pd.notna(val) else "", colored_fmt)
                        else:
                            worksheet.write(target_r, target_c, val if pd.notna(val) else "", fmt)

            st.success("✅ Excel Generated with Colors and Merged Headers!")
            st.download_button("📥 Download", output.getvalue(), f"Formatted_{selected_sheet}.xlsx")
