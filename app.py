import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Market Research Templifier", layout="wide")

st.title("📊 Market Research Templifier")
st.write("Upload your raw results Excel file to filter and format.")

# Allow both Excel and CSV, but Sheet selection only applies to Excel
uploaded_file = st.file_uploader("Choose a file", type=["xlsx", "csv"])

if uploaded_file is not None:
    # 1. SHEET SELECTOR
    sheet_name = None
    if uploaded_file.name.endswith('.xlsx'):
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select the sheet to process", xl.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    else:
        df = pd.read_csv(uploaded_file, header=None)
    
    st.sidebar.header("Filter Settings")

    # 2. QUESTION SELECTION (Column A, from Row 5 index)
    # Filter rows starting from row 6 (index 5)
    questions_list = df.iloc[5:, 0].dropna().unique().tolist()
    selected_questions = st.sidebar.multiselect(
        "Select Questions to Keep", questions_list, default=questions_list
    )

    # 3. METRIC SELECTION (Usually Column B/C, from Row 5 index)
    # Using Column index 1 as the Metric column
    metrics_list = df.iloc[5:, 1].dropna().unique().tolist()
    selected_metrics = st.sidebar.multiselect(
        "Select Metrics to Keep", metrics_list, default=metrics_list
    )

    # 4. PRODUCT SELECTION (Starting Column E)
    # Triplet logic: EFG (4,5,6), HIJ (7,8,9), etc.
    # Name is in Row 3 (index 2)
    product_triplets = {}
    # Iterate through columns starting from E (index 4) in steps of 3
    for col_idx in range(4, len(df.columns) - 2, 3):
        p_name = df.iloc[2, col_idx]
        if pd.isna(p_name): # Fallback to Row 5 label if Row 3 is empty
            p_name = f"Product at Col {col_idx+1}"
        
        product_triplets[p_name] = [col_idx, col_idx + 1, col_idx + 2]

    selected_products = st.sidebar.multiselect(
        "Select Products to Keep (Columns E onwards)", 
        list(product_triplets.keys()), 
        default=list(product_triplets.keys())
    )

    # 5. SIG COLUMN TOGGLE
    # Specific columns mentioned: D (3), F (5), I (8), L (11), O (14)...
    # Pattern: Column 3, and then index 5 + 3n
    show_sig = st.sidebar.checkbox("Show Significance/Delta Columns (D, F, I, L...)", value=True)

    if st.button("Generate Template"):
        # --- ROW FILTERING ---
        # Keep header rows (0-4) and rows matching selections
        header_rows = df.iloc[:5]
        data_rows = df.iloc[5:]
        
        filtered_data = data_rows[
            (data_rows[0].isin(selected_questions)) & 
            (data_rows[1].isin(selected_metrics))
        ]
        
        final_df_rows = pd.concat([header_rows, filtered_data])

        # --- COLUMN FILTERING ---
        # Always keep Col A (0), B (1), C (2)
        cols_to_keep = [0, 1, 2]
        
        # Column D (3) logic
        if show_sig:
            cols_to_keep.append(3)

        # Products (E onwards) logic
        for p_name, indices in product_triplets.items():
            if p_name in selected_products:
                if show_sig:
                    # Keep all 3 columns (e.g., E, F, G)
                    cols_to_keep.extend(indices)
                else:
                    # Hide the middle one (F, I, L...)
                    # Keep first and third (e.g., E and G)
                    cols_to_keep.append(indices[0])
                    cols_to_keep.append(indices[2])

        final_df = final_df_rows[cols_to_keep]

        # --- EXPORT ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, header=False, sheet_name='Templated_Results')
            
            # Basic formatting to make it look professional
            workbook = writer.book
            worksheet = writer.sheets['Templated_Results']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
            
            # Apply header format to row 3 (Product names)
            for col_num, value in enumerate(final_df.columns):
                worksheet.write(2, col_num, final_df.iloc[2, col_num], header_format)
        
        st.success("Template Processed!")
        st.download_button(
            label="Download Cleaned Excel",
            data=output.getvalue(),
            file_name=f"Templated_{sheet_name if sheet_name else 'Results'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
