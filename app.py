import streamlit as st
import pandas as pd
import pymupdf as fitz
import json
import os
import time
import io
from google import genai
from google.genai import types

RULES_FILE = "rules.json"
DEFAULT_CATEGORIES = ["General / Universal", "Oven", "Paint Booth", "Conveyors"]

def load_rules():
    """Loads rules as a dictionary: {'Oven': [...], 'Paint Booth': [...]}. Handles legacy list conversion."""
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, "r") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return {"General / Universal": data, "Oven": [], "Paint Booth": [], "Conveyors": []}
                elif isinstance(data, dict):
                    for cat in DEFAULT_CATEGORIES:
                        if cat not in data:
                            data[cat] = []
                    return data
        except Exception:
            pass
    return {cat: [] for cat in DEFAULT_CATEGORIES}

def save_rules(rules_dict):
    with open(RULES_FILE, "w") as f:
        json.dump(rules_dict, f, indent=4)

def generate_excel_report(audit_result, disc_df, category, model_used, total_files, total_pages, stock_length_mm):
    """Generates an in-memory Excel file with Summary and Discrepancies tabs."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Audit Summary & Metadata
        summary_data = {
            "Parameter": [
                "Equipment Category",
                "Standard Stock Profile Length",
                "Total Drawing Files Scanned",
                "Total Drawing Pages Scanned",
                "Audit Status",
                "AI Model Used",
                "Total Discrepancies Found",
                "General Audit Notes"
            ],
            "Details": [
                category,
                f"{stock_length_mm} mm",
                total_files,
                total_pages,
                audit_result.get("audit_status", "N/A"),
                model_used,
                len(disc_df) if disc_df is not None else 0,
                audit_result.get("general_notes", "")
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Audit Summary", index=False)
        
        # Sheet 2: Discrepancies Table
        if disc_df is not None and not disc_df.empty:
            disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)
        else:
            no_disc_df = pd.DataFrame([{"Status": "No discrepancies found. All normalized quantities and drawing specs match ERP BOM."}])
            no_disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)
            
    output.seek(0)
    return output

st.set_page_config(page_title="Categorized ERP BOM & Drawing Audit Agent", layout="wide")

st.title("🛠️ Engineering BOM & Drawing Audit Agent")
st.markdown("Multi-file, multi-page agent for cross-checking multiple PDF drawings against a single ERP CSV BOM with strict Column Disambiguation.")

# --- LOAD MEMORY DATA ---
all_rules = load_rules()

# --- SIDEBAR: CONFIG & CATEGORY MEMORY MANAGEMENT ---
st.sidebar.header("1. Agent Configuration")
api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

selected_model = st.sidebar.selectbox(
    "Preferred Gemini Model",
    ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.header("2. Equipment & Category Memory")

category_list = list(all_rules.keys())
selected_category = st.sidebar.selectbox("Select Active Equipment Category", category_list)

with st.sidebar.expander("➕ Add New Equipment Category"):
    new_cat_name = st.text_input("Category Name (e.g. 'Air Handling Unit')")
    if st.button("Create Category"):
        if new_cat_name and new_cat_name not in all_rules:
            all_rules[new_cat_name] = []
            save_rules(all_rules)
            st.success(f"Category '{new_cat_name}' created!")
            st.rerun()

with st.sidebar.expander(f"🧠 Manage Rules for [{selected_category}]"):
    cat_rules = all_rules.get(selected_category, [])
    
    new_rule = st.text_area(f"Teach a rule for {selected_category}:", placeholder="e.g. 'Oven panels must specify 100mm Rockwool density'")
    if st.button("Save Rule to Category"):
        if new_rule:
            all_rules[selected_category].append(new_rule)
            save_rules(all_rules)
            st.success("Rule saved permanently to category memory!")
            st.rerun()
            
    if cat_rules:
        st.write(f"**Saved Rules ({selected_category}):**")
        for i, r in enumerate(cat_rules, 1):
            st.text(f"{i}. {r}")
        if st.button(f"Clear All Rules for {selected_category}"):
            all_rules[selected_category] = []
            save_rules(all_rules)
            st.rerun()
    else:
        st.caption("No specific rules learned for this category yet.")

st.sidebar.markdown("---")
st.sidebar.header("3. Normalization Settings")
stock_length_mm = st.sidebar.number_input(
    "Standard Stock Profile Length (mm)",
    min_value=500,
    max_value=12000,
    value=3000,
    step=500,
    help="Default raw material length (e.g., 3000mm, 6000mm) used to convert total cut length into ERP piece counts."
)

# --- MAIN APP INTERFACE ---
if not api_key:
    st.warning("Please enter your Gemini API Key in the sidebar to proceed.")
else:
    st.info(f"📌 **Active Audit Mode:** `Category: {selected_category}` | Multi-PDF File Upload Enabled | Stock Length: `{stock_length_mm} mm` | Active Normalizations: Column Guard (ITEM vs QTY) + Variant Suffixes (-A/-B) & Cut-Lengths (-1000mm)")

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📁 Upload ERP BOM (.csv)")
        csv_file = st.file_uploader("Upload single ERP BOM CSV file", type=["csv"])
        
    with col2:
        st.subheader("📐 Upload Drawing PDFs (.pdf)")
        pdf_files = st.file_uploader(
            "Upload one or multiple Engineering Drawing PDFs", 
            type=["pdf"], 
            accept_multiple_files=True
        )

    if csv_file and pdf_files:
        erp_df = pd.read_csv(csv_file)
        
        st.write(f"**Uploaded ERP BOM Data ({len(erp_df)} total rows):**")
        st.dataframe(erp_df, use_container_width=True)
        st.write(f"**Selected PDF Drawing Files:** {len(pdf_files)} file(s) ready for processing.")

        if st.button("Run Full Multi-PDF Cross-Check Audit", type="primary"):
            with st.spinner("Processing all PDF drawing pages across all files, compiling BOM data, and auditing against ERP CSV..."):
                try:
                    # 1. Convert ALL PDF Pages from ALL Uploaded PDF Files into Image Parts
                    image_parts = []
                    total_pages = 0
                    total_files = len(pdf_files)

                    for pdf_file in pdf_files:
                        pdf_bytes = pdf_file.read()
                        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                        total_pages += len(doc)
                        
                        for page_num in range(len(doc)):
                            page = doc[page_num]
                            pix = page.get_pixmap(dpi=300)
                            img_bytes = pix.tobytes("png")
                            image_parts.append(
                                types.Part.from_bytes(data=img_bytes, mime_type="image/png")
                            )

                    st.info(f"📄 Successfully loaded {total_pages} page(s) across {total_files} PDF drawing file(s) for audit.")

                    # 2. Compile Active Rules Context
                    universal_rules = all_rules.get("General / Universal", [])
                    category_rules = all_rules.get(selected_category, [])
                    
                    combined_rules = []
                    if universal_rules:
                        combined_rules.append("General/Universal Engineering Rules:")
                        combined_rules.extend([f" - {r}" for r in universal_rules])
                    
                    if category_rules and selected_category != "General / Universal":
                        combined_rules.append(f"\nSpecific Rules for [{selected_category}]:")
                        combined_rules.extend([f" - {r}" for r in category_rules])

                    rules_context_str = "\n".join(combined_rules) if combined_rules else "None specified."

                    # 3. Prepare API client & prompt for Multi-PDF Cross-Check
                    client = genai.Client(api_key=api_key)
                    erp_csv_text = erp_df.to_csv(index=False)

                    prompt = f"""
                    You are an expert mechanical engineering AI auditor specializing in {selected_category} manufacturing systems.
                    You are provided with all drawing pages from {total_files} engineering PDF drawing files and a single ERP BOM CSV dataset.
                    
                    Equipment Category Being Audited: {selected_category}
                    Configured Standard Stock Length: {stock_length_mm} mm
                    
                    Mandatory Learned Memory Rules:
                    {rules_context_str}
                    
                    CRITICAL MANDATORY PART CODE NORMALIZATION & MULTI-DRAWING AGGREGATION RULES:
                    
                    0. STRICT TABLE COLUMN GUARD & DISAMBIGUATION (ITEM NO. vs QTY):
                       - Engineering drawing tables have distinct columns: [ITEM / ITEM NO.] | [QTY / QUANTITY] | [PART NUMBER] | [DESCRIPTION].
                       - 'ITEM' / 'ITEM NO.' is strictly the row index / serial position (e.g., Row 1, Row 2, ..., Row 16).
                       - 'QTY' / 'QUANTITY' is the piece count required for that line item.
                       - CRITICAL SAFETY MANDATE: NEVER use the 'ITEM' number as a quantity multiplier under any circumstances.
                       - SANITY CHECK EXAMPLE:
                         * Row line 16 has: ITEM = 16, QTY = 1, PART NUMBER = 'SF0000000236-1553'.
                         * Quantity multiplier MUST be 1 (from QTY column).
                         * Length calculation = 1553 mm x 1 = 1553 mm.
                         * ABSOLUTELY BANNED: Multiplying 1553 mm x 16 (using ITEM 16).
                       - Double check headers for every table column before parsing values.
                    
                    1. AGGREGATE ACROSS ALL FILES & PAGES:
                       - Collect and consolidate all BOM table entries and balloon callouts from ALL provided drawing pages across ALL uploaded PDF files into a single master drawing inventory.
                    
                    2. RULE A: VARIANT SUFFIX NORMALIZATION (Letter Suffixes like -A, -B)
                       - Suffixes containing letters (e.g., '-A', '-B', '-REV1') indicate variants or revisions.
                       - Strip the variant suffix to find the Parent Part Code (e.g., 'SF0000000001-A' -> 'SF0000000001').
                       - Sum quantities directly across ALL drawing pages/files (e.g., Base Qty 5 + Variant-A Qty 1 + Variant-B Qty 2 = Total 7 pcs).
                    
                    3. RULE B: CUT-LENGTH SUFFIX NORMALIZATION & {stock_length_mm}mm STOCK CONVERSION (Numeric Suffixes like -1000, -1500)
                       - Suffixes containing numbers (e.g., '-1000', '-1500', '-2500') represent cut lengths in millimeters (mm).
                       - Strip the length suffix to identify the Parent Part Code (e.g., 'SF0000000020-1000' -> Parent Code 'SF0000000020').
                       - For each cut length entry across all drawings, calculate length contribution: (Cut Length in mm) x (Quantity from QTY column ONLY).
                       - Sum ALL millimeter contributions for the same Parent Part Code across ALL files/pages to get TOTAL MILLIMETERS REQUIRED.
                       - CONVERT TO ERP STOCK PIECES:
                         * ERP BOM lists raw profiles in standard {stock_length_mm} mm stock lengths (Pieces).
                         * Calculate required stock pieces = ceil(TOTAL MILLIMETERS REQUIRED / {stock_length_mm}). Always round UP to the next whole integer.
                       - WORKED CALCULATION EXAMPLE (using {stock_length_mm} mm stock):
                         * Item 16: 'SF0000000236-1553', Qty: 1 -> 1553 mm x 1 = 1553 mm (NOT 1553 x 16!)
                         * Item 17: 'SF0000000236-1750', Qty: 1 -> 1750 mm x 1 = 1750 mm
                         * Total required length for 'SF0000000236' = 1553 + 1750 = 3303 mm
                         * ERP Stock Pieces Required = ceil(3303 / {stock_length_mm}) = ceil(3303 / 3000) = 2 pieces.
                    
                    4. RULE C: CONSOLIDATED BOM BLOCK / SUMMARY TABLE LOOKUP
                       - Check if any drawing page across the files contains an explicitly written 'Consolidated BOM', 'Summary Table', or 'Assembly BOM'.
                       - Cross-verify your calculated parent quantities against any summary table present on the drawings.
                    
                    ERP AUDIT & ALARM TRIGGER:
                       - Compare the consolidated calculated drawing quantity for each Parent Part Code against the corresponding Parent Part Code Qty in the single uploaded ERP BOM CSV.
                       - IF ERP BOM Qty matches the consolidated required quantity, mark as OK.
                       - IF THERE IS ANY VARIATION, IMMEDIATELY trigger an ALARM / DISCREPANCY.
                       - In the discrepancy report, show the full multi-file math breakdown (e.g., "Drawings combined require 3303 mm [1553mm x1 + 1750mm x1] = ceil(3303/{stock_length_mm}) = 2 stock pcs vs ERP BOM Qty: X").
                    
                    ERP BOM Data:
                    {erp_csv_text}
                    
                    Instructions:
                    1. Extract and aggregate all BOM items across ALL pages of ALL uploaded drawing files.
                    2. STRICTLY ensure multipliers are taken ONLY from the 'QTY' column, NEVER the 'ITEM' column.
                    3. Apply Rule A or Rule B normalization depending on suffix type.
                    4. Compare consolidated parent totals against the single ERP BOM CSV.
                    5. Check for missing parts, extra parts, or rule violations based on learned memory rules.
                    6. Output a structured JSON response matching the required schema.
                    """

                    output_schema = {
                        "type": "OBJECT",
                        "properties": {
                            "audit_status": {"type": "STRING"},
                            "general_notes": {"type": "STRING"},
                            "discrepancies": {
                                "type": "ARRAY",
                                "items": {
                                    "type": "OBJECT",
                                    "properties": {
                                        "part_number": {"type": "STRING"},
                                        "issue_type": {"type": "STRING"},
                                        "drawing_details": {"type": "STRING"},
                                        "erp_details": {"type": "STRING"},
                                        "recommendation": {"type": "STRING"}
                                    },
                                    "required": ["part_number", "issue_type", "drawing_details", "erp_details", "recommendation"]
                                }
                            }
                        },
                        "required": ["audit_status", "discrepancies", "general_notes"]
                    }

                    request_contents = [prompt] + image_parts

                    # Fallback Execution Loop across models
                    candidate_models = [selected_model, "gemini-2.5-flash", "gemini-1.5-flash"]
                    candidate_models = list(dict.fromkeys(candidate_models))

                    response = None
                    successful_model = None
                    last_error = None

                    for model_name in candidate_models:
                        try:
                            response = client.models.generate_content(
                                model=model_name,
                                contents=request_contents,
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    response_schema=output_schema
                                )
                            )
                            successful_model = model_name
                            break
                        except Exception as err:
                            last_error = err
                            err_msg = str(err)
                            if "503" in err_msg or "UNAVAILABLE" in err_msg:
                                st.warning(f"Model `{model_name}` is experiencing high traffic (503). Retrying automatically with backup model...")
                                time.sleep(1)
                                continue
                            else:
                                raise err

                    if response is None:
                        raise last_error

                    audit_result = json.loads(response.text)
                    
                    st.success(f"Audit Complete across {total_files} drawing file(s) / {total_pages} page(s)! (Executed with `{successful_model}`)")
                    st.write(f"**Status:** {audit_result.get('audit_status')}")
                    st.write(f"**Notes:** {audit_result.get('general_notes')}")
                    
                    disc_list = audit_result.get("discrepancies", [])
                    disc_df = pd.DataFrame(disc_list) if disc_list else None

                    if disc_df is not None and not disc_df.empty:
                        st.warning(f"Found {len(disc_list)} Discrepancies/Checks:")
                        st.dataframe(disc_df, use_container_width=True)
                    else:
                        st.info("No discrepancies found. All aggregated normalized quantities, stock profile lengths, and drawing parts across all files match the ERP BOM perfectly.")

                    # --- EXCEL DOWNLOAD BUTTON ---
                    excel_data = generate_excel_report(audit_result, disc_df, selected_category, successful_model, total_files, total_pages, stock_length_mm)
                    filename_cat = selected_category.lower().replace(" ", "_").replace("/", "_")
                    
                    st.download_button(
                        label="📥 Download Consolidated Multi-Drawing Audit Report (.xlsx)",
                        data=excel_data,
                        file_name=f"audit_report_{filename_cat}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
