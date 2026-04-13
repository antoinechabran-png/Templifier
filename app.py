import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="Market Research Templifier", layout="wide")

# Constants for Metric Categorization
GENERAL_METRICS = [
    "Mean", "Top Box", "Top 2 Boxes", "Top 3 Boxes", 
    "Bottom Box", "Bottom 2 Boxes", "Bottom 3 Boxes"
]

st.title("📊 Market Research Templifier")
st.write("Clean data, merge duplicate question labels, and preserve formatting.")

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

    # --- PARSING PRODUCTS (Starting Column E / Index 4) ---
    product_triplets = {}
    for col_idx in range(4, len(df.columns) - 2, 3):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name):
            p_name = f"Product at Column {col_idx+1}"
        product_triplets[p_name] = [col_idx, col_idx + 1, col_idx + 2]

    # --- SIDEBAR SETTINGS ---
    st.sidebar.header("Product & Sig Settings")
    all_p_names = list(product_triplets.keys())
    selected_products = st.sidebar.multiselect(
        "Select Products to Keep", 
        all_p_names, 
        default=all_p_names
    )
    show_sig = st.sidebar.checkbox("Show Significance/Delta Columns (D, F, I, L...)", value=True)

    # --- MAIN UI: QUESTION & CATEGORIZED METRICS ---
    st.subheader("Step 1: Filter Questions and Metric Groups")
    
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    q_m_map = {}
    for q in raw_data_area[0].unique():
        metrics_for_q = df[df[0] == q][1].unique().tolist()
        q_m_map[q] = metrics_for_q

    selected_q_metrics = {}

    for q, metrics in q_m_map.items():
        with st.expander(f"❓ {q}", expanded=False):
            gen_group = [m for m in metrics if m in GENERAL_METRICS]
            mod_group = [m for m in metrics if m not in GENERAL_METRICS]

            col_q, col_gen, col_mod = st.columns([1, 2, 2])
            with col_q:
                is_q_active = st.checkbox("Keep Question", value=True, key=f"q_active_{q}")

            if is_q_active:
                with col_gen:
                    st.markdown("**General Metrics**")
                    all_gen = st.checkbox("Select All General", value=True, key=f"all_gen_{q}")
                    sel_gen = st.multiselect("Pick General", gen_group, 
                                            default=gen_group if all_gen else [], key=f"sel_gen_{q}")
                with col_mod:
                    st.markdown("**Answer Modalities**")
                    all_mod = st.checkbox("Select All Modalities", value=True, key=f"all_mod_{q}")
                    sel_mod = st.multiselect("Pick Modalities", mod_group, 
                                            default=mod_group if all_mod else [], key=f"sel_mod_{q}")
                selected_q_metrics[q] = sel_gen + sel_mod

    # --- PROCESSING ---
    if st.button("🚀 Generate Templated Excel", type="primary"):
        header_rows = df.iloc[:5]
        data_rows = df.iloc[5:]
        
        valid_rows_list = []
        for _, row in data_rows.iterrows():
            q_name = row[0]
            m_name = row[1]
            if q_name in selected_q_metrics and m_name in selected_q_metrics[q_name]:
                valid_rows_list.append(row)
        
        if not valid_rows_list:
            st.error("No data found for selection.")
        else:
            filtered_data = pd.DataFrame(valid_rows_list)
            final_df = pd.concat([header_rows, filtered_data]).reset_index(drop=True)

            cols_to_keep = [0, 1, 2]
            if show_sig:
                cols_to_keep.append(3)

            for p_name, indices in product_triplets.items():
                if p_name in selected_products:
                    if show_sig:
                        cols_to_keep.extend(indices)
                    else:
                        cols_to_keep.append(indices[0]) # Value
                        cols_to_keep.append(indices[2]) # Base/Delta

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

                # 1. Format Product Headers (Row 3)
                for col_num in range(len(final_df.columns)):
                    val = final_df.iloc[2, col_num]
                    if pd.notna(val):
                        worksheet.write(2, col_num, val, header_fmt)

                # 2. Merge Column A logic (Starting row index 5)
                start_row = 5
                row_count = len(final_df)
                if row_count > 5:
                    current_q = final_df.iloc[start_row, 0]
                    for r in range(start_row + 1, row_count):
                        new_q = final_df.iloc[r, 0]
                        if new_q != current_q:
                            if r - 1 > start_row:
                                worksheet.merge_range(start_row, 0, r - 1, 0, current_q, merge_fmt)
                            else:
                                worksheet.write(start_row, 0, current_q, merge_fmt)
                            start_row = r
                            current_q = new_q
                    # Handle last group
                    if row_count - 1 > start_row:
                        worksheet.merge_range(start_row, 0, row_count - 1, 0, current_q, merge_fmt)
                    else:
                        worksheet.write(start_row, 0, current_q, merge_fmt)

                # 3. Handle Data Formatting (% and Borders)
                for r in range(5, row_count):
                    for c in range(1, len(final_df.columns)):
                        val = final_df.iloc[r, c]
                        # Apply percentage format if Column B says it's a modality or if value is a small float
                        # (Adjusting logic: If the original metric is not "Mean", it's usually a %)
                        metric_type = final_df.iloc[r, 1]
                        
                        target_fmt = standard_fmt
                        if metric_type not in ["Mean"] and isinstance(val, (int, float)):
                            target_fmt = percent_fmt
                        
                        if pd.notna(val):
                            worksheet.write(r, c, val, target_fmt)
                        else:
                            worksheet.write(r, c, "", standard_fmt)

            st.success("✅ Success! Questions merged and percentages preserved.")
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name=f"Templated_{sheet_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
