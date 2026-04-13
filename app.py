import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Market Research Templifier", layout="wide")

st.title("📊 Market Research Templifier")

# Define the standard "General" metrics as per your list
GENERAL_METRICS = [
    "Mean", "Top Box", "Top 2 Boxes", "Top 3 Boxes", 
    "Bottom Box", "Bottom 2 Boxes", "Bottom 3 Boxes"
]

uploaded_file = st.file_uploader("Upload Raw Results (Excel or CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    # --- DATA LOADING ---
    if uploaded_file.name.endswith('.xlsx'):
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select the sheet to process", xl.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    else:
        df = pd.read_csv(uploaded_file, header=None)

    # --- PARSING PRODUCTS (Sidebar) ---
    product_triplets = {}
    for col_idx in range(4, len(df.columns) - 2, 3):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name):
            p_name = f"Product @ Col {col_idx+1}"
        product_triplets[p_name] = [col_idx, col_idx + 1, col_idx + 2]

    st.sidebar.header("Product & Sig Settings")
    selected_products = st.sidebar.multiselect(
        "Products to Include", 
        list(product_triplets.keys()), 
        default=list(product_triplets.keys())
    )
    show_sig = st.sidebar.checkbox("Show Significance/Delta Columns", value=True)

    # --- MAIN UI: QUESTION & CATEGORIZED METRICS ---
    st.subheader("Step 1: Filter Questions and Metric Groups")
    
    # Identify questions (Col A) and their metrics (Col B)
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    q_m_map = {}
    for q in raw_data_area[0].unique():
        metrics_for_q = df[df[0] == q][1].unique().tolist()
        q_m_map[q] = metrics_for_q

    selected_q_metrics = {}

    for q, metrics in q_m_map.items():
        with st.expander(f"❓ {q}", expanded=False):
            # Split metrics into two groups
            gen_group = [m for m in metrics if m in GENERAL_METRICS]
            mod_group = [m for m in metrics if m not in GENERAL_METRICS]

            col_q, col_gen, col_mod = st.columns([1, 2, 2])
            
            with col_q:
                is_q_active = st.checkbox("Keep Question", value=True, key=f"q_{q}")

            if is_q_active:
                with col_gen:
                    st.markdown("**General Metrics**")
                    all_gen = st.checkbox("Select All General", value=True, key=f"all_gen_{q}")
                    sel_gen = st.multiselect(
                        "Pick General", gen_group, 
                        default=gen_group if all_gen else [], 
                        key=f"ms_gen_{q}"
                    )

                with col_mod:
                    st.markdown("**Answer Modalities**")
                    all_mod = st.checkbox("Select All Modalities", value=True, key=f"all_mod_{q}")
                    sel_mod = st.multiselect(
                        "Pick Modalities", mod_group, 
                        default=mod_group if all_mod else [], 
                        key=f"ms_mod_{q}"
                    )
                
                selected_q_metrics[q] = sel_gen + sel_mod

    # --- PROCESSING & EXPORT ---
    if st.button("🚀 Generate Templated Excel", type="primary"):
        header_rows = df.iloc[:5]
        data_rows = df.iloc[5:]
        
        # Row Filtering
        valid_rows = [row for _, row in data_rows.iterrows() 
                      if row[0] in selected_q_metrics and row[1] in selected_q_metrics[row[0]]]
        
        if not valid_rows:
            st.error("Please select at least one metric.")
        else:
            final_df_rows = pd.concat([header_rows, pd.DataFrame(valid_rows)])

            # Column Filtering (A, B, C constant; D optional; triplets optional)
            cols_to_keep = [0, 1, 2]
            if show_sig: cols_to_keep.append(3)

            for p_name in selected_products:
                indices = product_triplets[p_name]
                if show_sig:
                    cols_to_keep.extend(indices)
                else:
                    cols_to_keep.append(indices[0]) # Value
                    cols_to_keep.append(indices[2]) # Base

            final_df = final_df_rows[cols_to_keep]

            # Write to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, header=False, sheet_name='Final_Report')
                
                # Auto-formatting
                workbook = writer.book
                worksheet = writer.sheets['Final_Report']
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
                
                for col_num in range(len(final_df.columns)):
                    val = final_df.iloc[2, col_num]
                    if pd.notna(val):
                        worksheet.write(2, col_num, val, header_fmt)

            st.success("✅ Success! Your file is ready.")
            st.download_button("📥 Download Excel", output.getvalue(), file_name="Templifier_Result.xlsx")
            )
