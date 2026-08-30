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

def generate_excel_report(audit_result, disc_df, category, model_used):
    """Generates an in-memory Excel file with Summary and Discrepancies tabs."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Audit Summary & Metadata
        summary_data = {
            "Parameter": [
                "Equipment Category",
                "Audit Status",
                "AI Model Used",
                "Total Discrepancies Found",
                "General Audit Notes"
            ],
            "Details": [
                category,
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
            no_disc_df = pd.DataFrame([{"Status": "No discrepancies found between ERP BOM and Drawing."}])
            no_disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)
            
    output.seek(0)
    return output

st.set_page_config(page_title="Categorized ERP BOM & Drawing Audit Agent", layout="wide")

st.title("🛠️ Engineering BOM & Drawing Audit Agent")
st.markdown("Multi-category agent for cross-checking ERP CSV BOMs against drawing PDFs with category-specific learned rules.")

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

# --- MAIN APP INTERFACE ---
if not api_key:
    st.warning("Please enter your Gemini API Key in the sidebar to proceed.")
else:
    st.info(f"📌 **Active Audit Mode:** `Equipment Category: {selected_category}` | Active Rules: `{len(all_rules.get(selected_category, []))}` category rules, `{len(all_rules.get('General / Universal', []))}` universal rules.")

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📁 Upload ERP BOM (.csv)")
        csv_file = st.file_uploader("Upload exported ERP BOM", type=["csv"])
        
    with col2:
        st.subheader("📐 Upload Drawing (.pdf)")
        pdf_file = st.file_uploader("Upload Engineering Drawing PDF", type=["pdf"])
        drawing_page_num = st.number_input("Target Drawing Page Number", min_value=1, value=2, step=1)

    if csv_file and pdf_file:
        erp_df = pd.read_csv(csv_file)
        
        st.write(f"**Uploaded ERP BOM Data ({len(erp_df)} total rows):**")
        st.dataframe(erp_df, use_container_width=True)

        if st.button("Run Category-Specific Cross-Check Audit", type="primary"):
            with st.spinner(f"Agent is analyzing {selected_category} drawing, parsing BOM, and applying active rules..."):
                try:
                    # 1. Process PDF Page to Image
                    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
                    page = doc[drawing_page_num - 1]
                    pix = page.get_pixmap(dpi=300)
                    image_path = "temp_audit_page.png"
                    pix.save(image_path)

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

                    # 3. Prepare API client & prompt
                    client = genai.Client(api_key=api_key)
                    erp_csv_text = erp_df.to_csv(index=False)

                    prompt = f"""
                    You are an expert mechanical engineering AI auditor specializing in {selected_category} manufacturing systems.
                    Your task is to cross-check the BOM extracted from the provided engineering drawing image against the provided ERP BOM CSV data.
                    
                    Equipment Category Being Audited: {selected_category}
                    
                    Mandatory Rules to Strictly Follow for this Category:
                    {rules_context_str}
                    
                    ERP BOM Data:
                    {erp_csv_text}
                    
                    Instructions:
                    1. Extract the BOM table items (Qty, Part Number, Description, Material/Spec) from the drawing image.
                    2. Compare them against the ERP BOM CSV data.
                    3. Identify any mismatches in quantities, missing part numbers, extra items, or specific violations of the category rules provided above.
                    4. Return a structured JSON response containing an array of discrepancies found, or a clear confirmation if everything matches perfectly.
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

                    # Fallback Execution Loop
                    candidate_models = [selected_model, "gemini-2.5-flash", "gemini-1.5-flash"]
                    candidate_models = list(dict.fromkeys(candidate_models))

                    response = None
                    successful_model = None
                    last_error = None

                    for model_name in candidate_models:
                        try:
                            response = client.models.generate_content(
                                model=model_name,
                                contents=[
                                    prompt,
                                    types.Part.from_bytes(data=open(image_path, "rb").read(), mime_type="image/png")
                                ],
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
                    
                    st.success(f"Audit Complete for [{selected_category}]! (Executed with `{successful_model}`)")
                    st.write(f"**Status:** {audit_result.get('audit_status')}")
                    st.write(f"**Notes:** {audit_result.get('general_notes')}")
                    
                    disc_list = audit_result.get("discrepancies", [])
                    disc_df = pd.DataFrame(disc_list) if disc_list else None

                    if disc_df is not None and not disc_df.empty:
                        st.warning(f"Found {len(disc_list)} Discrepancies/Checks:")
                        st.dataframe(disc_df, use_container_width=True)
                    else:
                        st.info("No discrepancies found between the drawing and the ERP BOM.")

                    # --- EXCEL DOWNLOAD BUTTON ---
                    excel_data = generate_excel_report(audit_result, disc_df, selected_category, successful_model)
                    filename_cat = selected_category.lower().replace(" ", "_").replace("/", "_")
                    
                    st.download_button(
                        label="📥 Download Audit Report (.xlsx)",
                        data=excel_data,
                        file_name=f"audit_report_{filename_cat}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
