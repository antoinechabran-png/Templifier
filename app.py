import streamlit as st
import pandas as pd
import io
import re

# Set page config
st.set_page_config(page_title="Market Research Templifier", layout="wide")

# Constants for Metric Categorization
GENERAL_METRICS = [
    "Mean", "Top Box", "Top 2 Boxes", "Top 3 Boxes", 
    "Bottom Box", "Bottom 2 Boxes", "Bottom 3 Boxes"
]

def get_question_root(q_text):
    """Extracts the root (e.g., Q-04) from a string like Q-04-1-Premium."""
    match = re.match(r"^(Q-\d+|S-\d+)", str(q_text))
    return match.group(1) if match else str(q_text)

st.title("📊 Market Research Templifier")
st.write("Clean data, regroup question attributes, and preserve formatting.")

# 1. FILE UPLOADER
uploaded_file = st.file_uploader("Upload Raw Results (Excel or CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    # --- DATA LOADING ---
    sheet_name = "Results"
    if uploaded_file.name.endswith('.xlsx'):
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select the sheet to process", xl.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    else:
        df = pd.read_csv(uploaded_file, header=None)

    # Pre-clean Column names and metrics for consistency (remove leading/trailing spaces)
    df[0] = df[0].astype(str).str.strip()
    df[1] = df[1].astype(str).str.strip()

    # --- PARSING PRODUCTS ---
    product_triplets = {}
    for col_idx in range(4, len(df.columns) - 2, 3):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name):
            p_name = f"Product at Column {col_idx+1}"
        product_triplets[p_name] = [col_idx, col_idx + 1, col_idx + 2]

    # --- SIDEBAR SETTINGS ---
    st.sidebar.header("Global Settings")
    regroup_mode = st.sidebar.toggle("Enable Question Regrouping (e.g., group all Q-04-x)", value=True)
    
    all_p_names = list(product_triplets.keys())
    selected_products = st.sidebar.multiselect(
        "Select Products to Keep", 
        all_p_names, 
        default=all_p_names
    )
    show_sig = st.sidebar.checkbox("Show Significance/Delta Columns", value=True)

    # --- MAIN UI: QUESTION & CATEGORIZED METRICS ---
    st.subheader("Step 1: Filter Questions and Metric Groups")
    
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    
    # Logic to build the Question Map
    ui_q_map = {}
    for q_full in raw_data_area[0].unique():
        if q_full in ["nan", "None", ""]: continue
        display_name = get_question_root(q_full) if regroup_mode else q_full
        metrics_for_this_q = df[df[0] == q_full][1].unique().tolist()
        
        if display_name not in ui_q_map:
            ui_q_map[display_name] = {"originals": [q_full], "metrics": set(metrics_for_this_q)}
        else:
            ui_q_map[display_name]["originals"].append(q_full)
            ui_q_map[display_name]["metrics"].update(metrics_for_this_q)

    # Dict to store the FINAL selection for processing
    selected_q_metrics = {} 

    for display_q, data in ui_q_map.items():
        with st.expander(f"❓ {display_q}", expanded=False):
            # Sort metrics and separate them
            metrics_list = sorted([str(m) for m in data["metrics"] if pd.notna(m)])
            gen_group = [m for m in metrics_list if m in GENERAL_METRICS]
            mod_group = [m for m in metrics_list if m not in GENERAL_METRICS]

            col_q, col_gen, col_mod = st.columns([1, 2, 2])
            with col_q:
                is_q_active = st.checkbox("Keep Group", value=True, key=f"active_{display_q}")

            if is_q_active:
                # Keys for multiselect persistence
                m_gen_key = f"ms_gen_{display_q}"
                m_mod_key = f"ms_mod_{display_q}"
                
                # Initialize session state if first time
                if m_gen_key not in st.session_state: st.session_state[m_gen_key] = gen_group
                if m_mod_key not in st.session_state: st.session_state[m_mod_key] = mod_group

                # GENERAL METRICS COLUMN
                with col_gen:
                    st.markdown("**General Metrics**")
                    b1, b2 = st.columns(2)
                    if b1.button("Select All", key=f"all_gen_btn_{display_q}"):
                        st.session_state[m_gen_key] = gen_group
                        st.rerun()
                    if b2.button("Unselect All", key=f"clr_gen_btn_{display_q}"):
                        st.session_state[m_gen_key] = []
                        st.rerun()
                    
                    sel_gen = st.multiselect("Pick General", gen_group, key=m_gen_key)

                # MODALITIES COLUMN
                with col_mod:
                    st.markdown("**Answer Modalities**")
                    b1, b2 = st.columns(2)
                    if b1.button("Select All", key=f"all_mod_btn_{display_q}"):
                        st.session_state[m_mod_key] = mod_group
                        st.rerun()
                    if b2.button("Unselect All", key=f"clr_mod_btn_{display_q}"):
                        st.session_state[m_mod_key] = []
                        st.rerun()
                        
                    sel_mod = st.multiselect("Pick Modalities", mod_group, key=m_mod_key)
                
                # Assign the combined set to every original attribute name in the group
                combined_selection = set(sel_gen + sel_mod)
                for orig in data["originals"]:
                    selected_q_metrics[orig] = combined_selection

    # --- PROCESSING ---
    if st.button("🚀 Generate Templated Excel", type="primary"):
        header_rows = df.iloc[:5]
        data_rows = df.iloc[5:]
        
        valid_rows_list = []
        for _, row in data_rows.iterrows():
            q_name = str(row[0]).strip()
            m_name = str(row[1]).strip()
            
            # THE FILTER: Only keep if the question is active AND the specific metric was kept in the multiselect
            if q_name in selected_q_metrics:
                if m_name in selected_q_metrics[q_name]:
                    valid_rows_list.append(row)
        
        if not valid_rows_list:
            st.error("No data found for selection. Ensure you haven't unselected everything.")
        else:
            filtered_data = pd.DataFrame(valid_rows_list)
            
            # Define columns to keep
            cols_to_keep = [0, 1, 2]
            if show_sig: cols_to_keep.append(3)
            for p_name, indices in product_triplets.items():
                if p_name in selected_products:
                    if show_sig:
                        cols_to_keep.extend(indices)
                    else:
                        cols_to_keep.append(indices[0]) # Value
                        cols_to_keep.append(indices[2]) # Base

            final_df = pd.concat([header_rows, filtered_data]).reset_index(drop=True)
            final_df = final_df[cols_to_keep]

            # --- EXPORT TO EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, header=False, sheet_name='Final_Report')
                workbook = writer.book
                worksheet = writer.sheets['Final_Report']
                
                # FORMATS
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                merge_fmt = workbook.add_format({'valign': 'vcenter', 'align': 'left', 'border': 1, 'text_wrap': True})
                percent_fmt = workbook.add_format({'num_format': '0%', 'border': 1})
                standard_fmt = workbook.add_format({'border': 1})

                # 1. Headers
                for col_num in range(len(final_df.columns)):
                    val = final_df.iloc[2, col_num]
                    if pd.notna(val): worksheet.write(2, col_num, val, header_fmt)

                # 2. Merge Column A
                start_row = 5
                row_count = len(final_df)
                if row_count > 5:
                    current_q = final_df.iloc[start_row, 0]
                    for r in range(start_row + 1, row_count):
                        new_q = final_df.iloc[r, 0]
                        if new_q != current_q:
                            if r - 1 > start_row: worksheet.merge_range(start_row, 0, r - 1, 0, current_q, merge_fmt)
                            else: worksheet.write(start_row, 0, current_q, merge_fmt)
                            start_row, current_q = r, new_q
                    if row_count - 1 > start_row: worksheet.merge_range(start_row, 0, row_count - 1, 0, current_q, merge_fmt)
                    else: worksheet.write(start_row, 0, current_q, merge_fmt)

                # 3. Numeric Formatting
                for r in range(5, row_count):
                    m_type = str(final_df.iloc[r, 1])
                    for c in range(1, len(final_df.columns)):
                        val = final_df.iloc[r, c]
                        fmt = percent_fmt if (m_type not in GENERAL_METRICS and isinstance(val, (int, float))) else standard_fmt
                        worksheet.write(r, c, val if pd.notna(val) else "", fmt)

            st.success("✅ Success! Filtered Excel generated.")
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name=f"Templated_{sheet_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
