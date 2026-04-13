import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Market Research Templifier", layout="wide")

st.title("📊 Market Research Templifier")
st.write("Upload your Excel results. Customize questions, metrics per question, and products.")

# 1. FILE UPLOADER
uploaded_file = st.file_uploader("Choose a file", type=["xlsx", "csv"])

if uploaded_file is not None:
    # 2. SHEET SELECTION
    sheet_name = None
    if uploaded_file.name.endswith('.xlsx'):
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select the sheet to process", xl.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    else:
        df = pd.read_csv(uploaded_file, header=None)

    # 3. DATA PARSING - QUESTIONS & METRICS
    # We map every unique question in Col A to its available metrics in Col B
    # Data starts at row 6 (index 5)
    raw_data_area = df.iloc[5:, [0, 1]].dropna(subset=[0])
    q_m_map = {}
    for q in raw_data_area[0].unique():
        # Get all metrics associated with this specific question
        metrics_for_q = df[df[0] == q][1].unique().tolist()
        q_m_map[q] = metrics_for_q

    # 4. PRODUCT PARSING (Starting Column E / index 4)
    # Triplet logic: Name is in Row 3 (index 2)
    product_triplets = {}
    for col_idx in range(4, len(df.columns) - 2, 3):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name):
            p_name = f"Product @ Col {col_idx+1}"
        product_triplets[p_name] = [col_idx, col_idx + 1, col_idx + 2]

    # --- SIDEBAR SETTINGS ---
    st.sidebar.header("Global Settings")
    
    # Product Multi-select (Default: All)
    all_p_names = list(product_triplets.keys())
    selected_products = st.sidebar.multiselect(
        "Products to Include", 
        all_p_names, 
        default=all_p_names
    )

    # Sig Column Toggle
    show_sig = st.sidebar.checkbox("Show Significance/Delta Columns", value=True)
    
    st.sidebar.divider()
    st.sidebar.info("Use the main area to select specific Metrics per Question.")

    # --- MAIN AREA: QUESTION & METRIC SELECTION ---
    st.subheader("Step 1: Select Questions & Specific Metrics")
    st.write("Toggle the checkbox to keep a question, then select which metrics for that question stay.")
    
    selected_q_metrics = {} # Format: {Question: [List of selected metrics]}

    for q, metrics in q_m_map.items():
        with st.expander(f"❓ {q}", expanded=False):
            cols = st.columns([1, 4])
            is_active = cols[0].checkbox("Keep Question", value=True, key=f"check_{q}")
            if is_active:
                sel_metrics = cols[1].multiselect(
                    f"Metrics for {q}", 
                    options=metrics, 
                    default=metrics, 
                    key=f"ms_{q}"
                )
                selected_q_metrics[q] = sel_metrics

    # --- PROCESSING ---
    if st.button("🚀 Generate Final Template", type="primary"):
        # ROW FILTERING
        # We only keep rows where (Question in keys) AND (Metric in selected list)
        header_rows = df.iloc[:5]
        data_rows = df.iloc[5:]
        
        valid_rows = []
        for idx, row in data_rows.iterrows():
            q_val = row[0]
            m_val = row[1]
            if q_val in selected_q_metrics and m_val in selected_q_metrics[q_val]:
                valid_rows.append(row)
        
        if not valid_rows:
            st.error("No data matches your selection. Please check your Question/Metric filters.")
        else:
            filtered_data = pd.DataFrame(valid_rows)
            final_df_rows = pd.concat([header_rows, filtered_data])

            # COLUMN FILTERING
            # Always keep A (0), B (1), C (2)
            cols_to_keep = [0, 1, 2]
            
            # Global Sig Column D (3)
            if show_sig:
                cols_to_keep.append(3)

            # Products (E onwards)
            for p_name in selected_products:
                indices = product_triplets[p_name]
                if show_sig:
                    cols_to_keep.extend(indices) # Keep Value, Sig, and Base
                else:
                    # Hide the "Sig" column (usually middle of the triplet)
                    cols_to_keep.append(indices[0]) # Value
                    cols_to_keep.append(indices[2]) # Base/Other

            final_df = final_df_rows[cols_to_keep]

            # --- EXPORT TO EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, header=False, sheet_name='Templated_Results')
                
                # Professional Formatting
                workbook = writer.book
                worksheet = writer.sheets['Templated_Results']
                
                # Style for Product Header
                header_fmt = workbook.add_format({
                    'bold': True, 
                    'bg_color': '#2E75B6', 
                    'font_color': 'white', 
                    'border': 1,
                    'align': 'center'
                })
                
                # Apply formatting to Row 3 (Product Names)
                for col_num in range(len(final_df.columns)):
                    val = final_df.iloc[2, col_num]
                    if pd.notna(val):
                        worksheet.write(2, col_num, val, header_fmt)

            st.success("✅ Template Created Successfully!")
            st.download_button(
                label="📥 Download Cleaned Excel",
                data=output.getvalue(),
                file_name=f"Templated_{sheet_name if sheet_name else 'Results'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
