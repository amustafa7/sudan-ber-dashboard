import streamlit as st
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
import pandas as pd
import google.generativeai as genai
import tempfile
import json
import re

st.set_page_config(page_title="Sudan BER Model Manager", layout="wide")

st.title("🌿 Sudan Biodiversity Expenditure Review (BER) — Automated Update & Dashboard")

# 1. API Setup
api_key = st.sidebar.text_input("Google AI Studio API Key", type="password")
if api_key:
    genai.configure(api_key=api_key)

# 2. Upload Base Workbook
uploaded_file = st.sidebar.file_uploader("Upload 'Sudan BER Model 2015 2025.xlsx'", type=["xlsx"])

def date_to_int(year, month):
    return year * 12 + month

def int_to_year_month(val):
    val = int(val)
    return f"{val // 12}-{val % 12:02d}"

if uploaded_file:
    # Load workbook
    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    ws_c = wb["C Spending"]
    
    # Locate C10 table bounds
    start_row = 135
    end_row = start_row
    while ws_c.cell(row=end_row, column=1).value is not None:
        end_row += 1
    current_projects_count = end_row - start_row

    st.sidebar.success(f"Workbook loaded! Active projects in C10: {current_projects_count}")

    # Tabs for Workflow
    tab_dash, tab_entry, tab_export = st.tabs(["📊 Interactive Dashboard", "➕ Project Ingestion (Manual & AI)", "💾 Export Workbook"])

    # --- TAB 1: DASHBOARD ---
    with tab_dash:
        # Read values from C10 for dynamic inspection
        proj_data = []
        for r in range(start_row, end_row):
            proj_data.append({
                "Project": ws_c.cell(row=r, column=1).value,
                "Agency": ws_c.cell(row=r, column=2).value,
                "Lifetime US$ m": ws_c.cell(row=r, column=14).value,
                "Category": ws_c.cell(row=r, column=25).value,
                "Sudan Share": ws_c.cell(row=r, column=17).value,
                "Rate Central": ws_c.cell(row=r, column=20).value,
            })
        df_proj = pd.DataFrame(proj_data)
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Active Projects", len(df_proj))
        col2.metric("Total Lifetime Budget (US$ m)", f"${df_proj['Lifetime US$ m'].sum():,.2f} M")
        col3.metric("Key Contributing Agencies", df_proj['Agency'].nunique())

        st.subheader("International Biodiversity Support by Agency")
        agency_summary = df_proj.groupby("Agency")["Lifetime US$ m"].sum().reset_index()
        st.bar_chart(agency_summary.set_index("Agency"))

        st.subheader("Project Database View (C10)")
        st.dataframe(df_proj, use_container_width=True)

    # --- TAB 2: DATA ENTRY (AI & MANUAL) ---
    with tab_entry:
        st.markdown("### Add New Project to Table C10")
        
        # Sub-section: Document Extraction
        extracted_data = {}
        with st.expander("📄 Optional: Extract Details from Project PDF / Contract using Gemini"):
            pdf_file = st.file_uploader("Upload Project Document (PDF)", type=["pdf"])
            if pdf_file and api_key:
                if st.button("Extract with Gemini"):
                    with st.spinner("Analyzing document via Gemini 1.5 Pro..."):
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(pdf_file.read())
                            tmp_path = tmp.name
                        
                        uploaded_gemini_file = genai.upload_file(tmp_path)
                        model = genai.GenerativeModel("gemini-1.5-pro")
                        prompt = """
                        Extract the project details for the Sudan Biodiversity Expenditure Review into JSON:
                        - project_title (string)
                        - agency (e.g., FAO, UNDP, UNEP, IFAD, JICA, etc.)
                        - lifetime_budget_usd_m (float in millions)
                        - start_year (int), start_month (int 1-12)
                        - end_year (int), end_month (int 1-12)
                        - sudan_share (float 0.0 to 1.0)
                        - rate_central (float 0.0 to 1.0 attribution to biodiversity)
                        - category_group (integer 1-9 for BIOFIN category)
                        - basis_grade (string)
                        - source_note (string)
                        Return ONLY valid JSON.
                        """
                        response = model.generate_content([uploaded_gemini_file, prompt])
                        clean_json = re.search(r'\{.*\}', response.text, re.DOTALL)
                        if clean_json:
                            extracted_data = json.loads(clean_json.group())
                            st.success("Extraction complete! Values populated below.")

        # Manual / Pre-filled Form
        with st.form("project_input_form"):
            col_a, col_b = st.columns(2)
            with col_a:
                f_title = st.text_input("Project Title (Col A)", value=extracted_data.get("project_title", ""))
                f_agency = st.text_input("Agency (Col B)", value=extracted_data.get("agency", "UNDP"))
                f_budget = st.number_input("Lifetime Budget (US$ m) (Col N)", value=float(extracted_data.get("lifetime_budget_usd_m", 1.0)), format="%.4f")
                f_s_yr = st.number_input("Start Year", min_value=2010, max_value=2030, value=int(extracted_data.get("start_year", 2021)))
                f_s_mo = st.number_input("Start Month", min_value=1, max_value=12, value=int(extracted_data.get("start_month", 1)))
                f_e_yr = st.number_input("End Year", min_value=2010, max_value=2035, value=int(extracted_data.get("end_year", 2025)))
                f_e_mo = st.number_input("End Month", min_value=1, max_value=12, value=int(extracted_data.get("end_month", 12)))

            with col_b:
                f_sudan_share = st.number_input("Sudan Share (Col Q)", value=float(extracted_data.get("sudan_share", 1.0)), min_value=0.0, max_value=1.0)
                f_delivery = st.number_input("Delivery Factor (Col R)", value=1.0, min_value=0.0, max_value=1.0)
                f_rate_cons = st.number_input("Rate Conservative (Col S)", value=0.05, min_value=0.0, max_value=1.0)
                f_rate_cent = st.number_input("Rate Central (Col T)", value=float(extracted_data.get("rate_central", 0.05)), min_value=0.0, max_value=1.0)
                f_rate_upp = st.number_input("Rate Upper (Col U)", value=0.05, min_value=0.0, max_value=1.0)
                f_cat = st.number_input("BIOFIN Category (1–9) (Col Y)", min_value=1, max_value=9, value=int(extracted_data.get("category_group", 7)))
                f_grade = st.text_input("Basis · Grade (Col Z)", value=extracted_data.get("basis_grade", "Actual delivery · B"))
                f_source = st.text_input("Source / Note (Col AA)", value=extracted_data.get("source_note", "Added via AI Ingestion"))

            submitted = st.form_submit_button("Append Project & Update All Sheets")

            if submitted:
                new_row = end_row
                ws_c.cell(row=new_row, column=1, value=f_title)
                ws_c.cell(row=new_row, column=2, value=f_agency)
                ws_c.cell(row=new_row, column=14, value=f_budget)
                
                start_val = date_to_int(f_s_yr, f_s_mo)
                end_val = date_to_int(f_e_yr, f_e_mo)
                ws_c.cell(row=new_row, column=15, value=start_val)
                ws_c.cell(row=new_row, column=16, value=end_val)
                ws_c.cell(row=new_row, column=17, value=f_sudan_share)
                ws_c.cell(row=new_row, column=18, value=f_delivery)
                ws_c.cell(row=new_row, column=19, value=f_rate_cons)
                ws_c.cell(row=new_row, column=20, value=f_rate_cent)
                ws_c.cell(row=new_row, column=21, value=f_rate_upp)
                ws_c.cell(row=new_row, column=22, value=1)
                ws_c.cell(row=new_row, column=23, value=1)
                ws_c.cell(row=new_row, column=24, value=1)
                ws_c.cell(row=new_row, column=25, value=f_cat)
                ws_c.cell(row=new_row, column=26, value=f_grade)
                ws_c.cell(row=new_row, column=27, value=f_source)

                # Replicate monthly pro-rata formula across years 2015-2025 (cols C to M)
                for col_idx, yr in enumerate(range(2015, 2026), start=3):
                    yr_start = date_to_int(yr, 1)
                    yr_end = date_to_int(yr, 12)
                    formula = (
                        f"=IF(MIN($P{new_row},{date_to_int(2025, 11)})<MAX($O{new_row},{date_to_int(2015, 0)}),0,"
                        f"$N{new_row}*MAX(0,MIN(MIN($P{new_row},{yr_end}),{yr_end})-MAX(MAX($O{new_row},{yr_start}),{yr_start})+1)/($P{new_row}-$O{new_row}+1))"
                    )
                    ws_c.cell(row=new_row, column=col_idx, value=formula)

                # Dynamically expand all referencing formulas across C Spending
                old_ref = f"$205"
                new_ref = f"${new_row}"
                for r in range(1, 130):
                    for c in range(1, 20):
                        cell = ws_c.cell(row=r, column=c)
                        if cell.value and isinstance(cell.value, str) and old_ref in cell.value:
                            cell.value = cell.value.replace(old_ref, new_ref)
                        elif isinstance(cell.value, ArrayFormula) and old_ref in cell.value.text:
                            cell.value = ArrayFormula(cell.value.ref, cell.value.text.replace(old_ref, new_ref))

                # Save updated workbook
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out_tmp:
                    wb.save(out_tmp.name)
                    st.session_state["updated_wb_path"] = out_tmp.name
                
                st.success(f"Project successfully appended at Row {new_row}! All rollups and 'Summary of the results' updated.")

    # --- TAB 3: EXPORT ---
    with tab_export:
        st.markdown("### Download Updated Model")
        if "updated_wb_path" in st.session_state:
            with open(st.session_state["updated_wb_path"], "rb") as f:
                st.download_button(
                    label="📥 Download Updated Sudan BER Model (Excel)",
                    data=f.read(),
                    file_name="Sudan_BER_Model_2015_2025_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("No updates made yet. Add a project in Tab 2 to generate an updated workbook.")
