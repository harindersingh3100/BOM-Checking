import json
import os
import io
import streamlit as st
import pandas as pd
import pymupdf as fitz
from google import genai
from google.genai import types

# --- PERSISTENT FILE STORAGE SETUP ---
DATA_DIR = os.path.expanduser("~/.audit_agent")
os.makedirs(DATA_DIR, exist_ok=True)
RULES_FILE = os.path.join(DATA_DIR, "rules.json")

DEFAULT_CATEGORIES = ["General / Universal", "Oven", "Paint Booth", "Conveyors"]

def load_rules():
    """Loads rules dictionary permanently from local persistent storage."""
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, "r") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    for cat in DEFAULT_CATEGORIES:
                        if cat not in data:
                            data[cat] = []
                    return data
        except Exception:
            pass
    return {cat: [] for cat in DEFAULT_CATEGORIES}

def save_rules(rules_dict):
    """Saves rules permanently to disk."""
    with open(RULES_FILE, "w") as f:
        json.dump(rules_dict, f, indent=4)

def generate_excel_report(audit_result, disc_df, master_df, category, model_used, total_files, total_pages, stock_length_mm):
    """Generates an in-memory Excel file with Summary, Discrepancies, and Full Master Reconciliation tabs."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # TAB 1: SUMMARY
        summary_data = {
            "Parameter": [
                "Equipment Category",
                "Standard Stock Profile Length",
                "Total Drawing Files Scanned",
                "Total Drawing Pages Scanned",
                "Audit Status",
                "AI Model Used",
                "Total Discrepancies Found",
                "Total Scanned Line Items",
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
                len(master_df) if master_df is not None else 0,
                audit_result.get("general_notes", "")
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Audit Summary", index=False)
        
        # TAB 2: DISCREPANCIES ONLY
        if disc_df is not None and not disc_df.empty:
            disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)
        else:
            no_disc_df = pd.DataFrame([{"Status": "No discrepancies found. All quantities, descriptions, and drawing specs match ERP BOM bi-directionally."}])
            no_disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)

        # TAB 3: FULL MASTER RECONCILIATION
        if master_df is not None and not master_df.empty:
            master_df.to_excel(writer, sheet_name="Full Master Reconciliation", index=False)
            
    output.seek(0)
    return output

st.set_page_config(
    page_title="Engineering ERP BOM & Drawing Audit Agent", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# CLEAN WINDOWS AQUA LIGHT THEME CUSTOM CSS
# ==========================================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #ebf5fb 0%, #d4e6f1 50%, #aed6f1 100%);
    color: #1c2833;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
}
.main .block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1280px;
}
h1 { color: #1b4f72 !important; font-weight: 700 !important; }
h2, h3, h4 { color: #21618c !important; font-weight: 600 !important; }
.stForm, div[data-testid="stExpander"] {
    background: rgba(255, 255, 255, 0.92) !important;
    border: 1px solid rgba(41, 128, 185, 0.25) !important;
    box-shadow: 0 4px 16px rgba(0, 51, 102, 0.08);
    backdrop-filter: blur(8px);
    border-radius: 10px !important;
    padding: 1.5rem;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 8px; background: rgba(255, 255, 255, 0.7);
    padding: 6px 10px; border-radius: 10px;
    border: 1px solid rgba(41, 128, 185, 0.2);
}
.stTabs [data-baseweb="tab"] {
    height: 42px; background-color: transparent;
    border-radius: 6px; color: #4a6572; font-weight: 600; padding: 0 20px;
}
.stTabs [aria-selected="true"] {
    background: #ffffff !important; color: #1b4f72 !important;
    border: 1px solid #2980b9 !important; box-shadow: 0 2px 6px rgba(41, 128, 185, 0.15);
}
.stButton > button[kind="primary"], div.stForm button[kind="primary"] {
    background: linear-gradient(180deg, #3498db 0%, #2980b9 100%) !important;
    color: #ffffff !important; font-weight: 600 !important;
    border: 1px solid #1f618d !important; border-radius: 6px !important;
    padding: 0.5rem 1.5rem !important;
}
section[data-testid="stSidebar"] { background-color: #f4f8fb !important; border-right: 1px solid #d4e6f1; }
.stTextInput input, .stSelectbox div[data-baseweb="select"], .stNumberInput input, .stTextArea textarea {
    background-color: #ffffff !important; color: #1c2833 !important; border: 1px solid #aed6f1 !important; border-radius: 6px !important;
}
[data-testid="stDataFrame"] { border: 1px solid #d4e6f1; border-radius: 8px; background-color: #ffffff; }
</style>
""", unsafe_allow_html=True)

st.title("🛠️ Engineering BOM & Drawing Audit Agent")

all_rules = load_rules()

if "active_category" not in st.session_state:
    st.session_state["active_category"] = list(all_rules.keys())[0]

# --- SIDEBAR CONFIGURATION ---
st.sidebar.markdown("### ⚙️ Engine Settings")
api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

selected_model = st.sidebar.selectbox(
    "Preferred Gemini Model",
    ["gemini-3.6-flash", "gemini-3-flash", "Custom Model Name"],
    index=0
)

if selected_model == "Custom Model Name":
    selected_model = st.sidebar.text_input("Enter Custom Model Identifier", value="gemini-3.6-flash")

stock_length_mm = st.sidebar.number_input(
    "Standard Stock Profile Length (mm)",
    min_value=500, max_value=12000, value=3000, step=500
)

st.sidebar.markdown("---")
st.sidebar.caption(f"💾 **Rules Database Storage:**\n`{RULES_FILE}`")

# --- MAIN NAVIGATION TABS ---
tab_audit, tab_categories, tab_rules = st.tabs([
    "🔍 Run Cross-Check Audit", 
    "📁 Category Manager", 
    "🧠 Check Rules (Table View)"
])

# ==========================================
# TAB 1: RUN BI-DIRECTIONAL CROSS-CHECK AUDIT
# ==========================================
with tab_audit:
    st.subheader("Run Bi-Directional Multi-PDF Drawing vs ERP BOM Cross-Check")
    
    category_list = list(all_rules.keys())
    selected_category = st.selectbox(
        "Select Equipment Category for Audit", 
        category_list, 
        index=category_list.index(st.session_state["active_category"]) if st.session_state["active_category"] in category_list else 0,
        key="audit_category_select"
    )

    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar to proceed.")
    else:
        st.info(f"📌 **Active Category:** `{selected_category}` | **Stock Profile:** `{stock_length_mm} mm` | **Active Rules:** `{len(all_rules.get(selected_category, []))} rule(s)`")

        col1, col2 = st.columns(2)
        with col1:
            bom_file = st.file_uploader("Upload ERP BOM File (Excel or CSV)", type=["xlsx", "xls", "csv"])
        with col2:
            pdf_files = st.file_uploader("Upload Drawing PDF Files", type=["pdf"], accept_multiple_files=True)

        if bom_file and pdf_files:
            try:
                file_name = bom_file.name.lower()
                if file_name.endswith(('.xlsx', '.xls')):
                    erp_df = pd.read_excel(bom_file)
                else:
                    try:
                        erp_df = pd.read_csv(bom_file, encoding='utf-8')
                    except UnicodeDecodeError:
                        bom_file.seek(0)
                        erp_df = pd.read_csv(bom_file, encoding='latin1')
            except Exception as e:
                st.error(f"Error reading ERP BOM file: {e}")
                st.stop()

            st.write(f"**ERP BOM Data Preview ({len(erp_df)} items loaded):**")
            st.dataframe(erp_df, use_container_width=True)

            if st.button("Run Full Bi-Directional Audit", type="primary"):
                with st.spinner("Processing drawing pages and performing full 2-way audit against ERP BOM..."):
                    try:
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
                                image_parts.append(
                                    types.Part.from_bytes(data=pix.tobytes("png"), mime_type="image/png")
                                )

                        universal_rules = all_rules.get("General / Universal", [])
                        category_rules = all_rules.get(selected_category, [])
                        combined_rules = []
                        if universal_rules:
                            combined_rules.append("General/Universal Rules:")
                            combined_rules.extend([f" - {r}" for r in universal_rules])
                        if category_rules and selected_category != "General / Universal":
                            combined_rules.append(f"\nCategory [{selected_category}] Rules:")
                            combined_rules.extend([f" - {r}" for r in category_rules])

                        rules_context_str = "\n".join(combined_rules) if combined_rules else "None specified."
                        client = genai.Client(api_key=api_key)

                        prompt = f"""
                        You are an expert mechanical engineering AI auditor specializing in {selected_category} manufacturing systems.
                        Perform a strict, deterministic BI-DIRECTIONAL (TWO-WAY) AUDIT comparing the provided PDF drawings against the ERP BOM data.
                        
                        Equipment Category: {selected_category}
                        Standard Stock Profile Length: {stock_length_mm} mm
                        Learned Rules:
                        {rules_context_str}
                        
                        DETERMINISTIC PROCESSING & SORTING INSTRUCTIONS:
                        1. SCAN ORDER: Scan drawing tables systematically from top-to-bottom and left-to-right.
                        2. OUTPUT SORTING: Sort BOTH the `master_parts_list` and `discrepancies` arrays in strictly ascending alphabetical order by `part_number`.
                        
                        CRITICAL AUDIT INSTRUCTIONS:
                        1. MASTER RECONCILIATION LIST (`master_parts_list`):
                           - Build a complete, itemized row-by-row reconciliation table for EVERY SINGLE unique part code present in either the drawing set or the ERP BOM.
                           - Include: `part_number`, `drawing_description`, `drawing_quantity`, `erp_description`, `erp_quantity`, `status`, `notes`.
                           - `status` MUST be one of: ["MATCH", "QUANTITY_MISMATCH", "DESCRIPTION_MISMATCH", "MISSING_IN_DRAWING", "MISSING_IN_ERP"].

                        2. DISCREPANCIES TABLE (`discrepancies`):
                           - Include all items where `status != "MATCH"`. Provide clear recommendations.

                        3. STRICT TABLE COLUMN GUARD:
                           - Engineering tables on drawings have distinct columns: [ITEM / ITEM NO.] | [QTY / QUANTITY] | [PART NUMBER] | [DESCRIPTION].
                           - 'ITEM' is purely the row index. 'QTY' is the quantity multiplier.
                           - NEVER use 'ITEM' index as a quantity multiplier!

                        4. PART DESCRIPTION COMPARISON:
                           - Compare text descriptions for matching part codes. Flag material grade, dimensions, or text mismatches as "DESCRIPTION_MISMATCH".

                        ERP BOM Data Table:
                        {erp_df.to_csv(index=False)}
                        """

                        output_schema = {
                            "type": "OBJECT",
                            "properties": {
                                "audit_status": {"type": "STRING"},
                                "general_notes": {"type": "STRING"},
                                "master_parts_list": {
                                    "type": "ARRAY",
                                    "items": {
                                        "type": "OBJECT",
                                        "properties": {
                                            "part_number": {"type": "STRING"},
                                            "drawing_description": {"type": "STRING"},
                                            "drawing_quantity": {"type": "STRING"},
                                            "erp_description": {"type": "STRING"},
                                            "erp_quantity": {"type": "STRING"},
                                            "status": {"type": "STRING"},
                                            "notes": {"type": "STRING"}
                                        },
                                        "required": [
                                            "part_number", "drawing_description", "drawing_quantity", 
                                            "erp_description", "erp_quantity", "status", "notes"
                                        ]
                                    }
                                },
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
                            "required": ["audit_status", "master_parts_list", "discrepancies", "general_notes"]
                        }

                        # Temperature set to 0.0 forces greedy deterministic output across repeat executions
                        response = client.models.generate_content(
                            model=selected_model,
                            contents=[prompt] + image_parts,
                            config=types.GenerateContentConfig(
                                temperature=0.0,
                                response_mime_type="application/json",
                                response_schema=output_schema
                            )
                        )

                        audit_result = json.loads(response.text)
                        st.success(f"Audit Complete! Executed with model `{selected_model}` (Deterministic Mode: Temp=0.0).")
                        st.write(f"**Status:** {audit_result.get('audit_status')}")
                        st.write(f"**Notes:** {audit_result.get('general_notes')}")
                        
                        disc_list = audit_result.get("discrepancies", [])
                        master_list = audit_result.get("master_parts_list", [])

                        disc_df = pd.DataFrame(disc_list) if disc_list else None
                        master_df = pd.DataFrame(master_list) if master_list else None

                        # Python-level secondary sorting to enforce exact identical UI ordering
                        if master_df is not None and not master_df.empty:
                            master_df = master_df.sort_values(by="part_number").reset_index(drop=True)

                        if disc_df is not None and not disc_df.empty:
                            disc_df = disc_df.sort_values(by="part_number").reset_index(drop=True)
                            st.warning(f"Found {len(disc_list)} Discrepancies:")
                            st.dataframe(disc_df, use_container_width=True)
                        else:
                            st.info("No discrepancies found. ERP BOM perfectly matches drawings bi-directionally.")

                        if master_df is not None and not master_df.empty:
                            with st.expander(f"📋 View Full Master Reconciliation Table ({len(master_df)} total items cross-checked)", expanded=True):
                                st.dataframe(master_df, use_container_width=True)

                        excel_data = generate_excel_report(
                            audit_result, disc_df, master_df, selected_category, 
                            selected_model, total_files, total_pages, stock_length_mm
                        )
                        st.download_button(
                            label="📥 Download Excel Audit Report (With Full Master Reconciliation)",
                            data=excel_data,
                            file_name=f"audit_{selected_category.lower().replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )

                    except Exception as e:
                        st.error(f"Error executing audit: {e}")

# ==========================================
# TAB 2: CATEGORY MANAGER
# ==========================================
with tab_categories:
    st.subheader("📁 Equipment Category Management")
    st.markdown("Add or delete equipment categories.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### ➕ Add New Equipment Category")
        with st.form("add_category_form", clear_on_submit=True):
            new_cat_input = st.text_input("New Category Name", placeholder="e.g. Air Handling Unit")
            submit_add_cat = st.form_submit_button("Create Category", type="primary")

            if submit_add_cat:
                clean_name = new_cat_input.strip()
                if clean_name:
                    if clean_name not in all_rules:
                        all_rules[clean_name] = []
                        save_rules(all_rules)
                        st.session_state["active_category"] = clean_name
                        st.success(f"Category **'{clean_name}'** successfully created!")
                        st.rerun()
                    else:
                        st.warning(f"Category '{clean_name}' already exists.")
                else:
                    st.error("Please enter a valid category name.")

    with col2:
        st.markdown("### 🗑️ Delete Equipment Category")
        deletable_categories = [c for c in all_rules.keys() if c != "General / Universal"]
        
        if deletable_categories:
            selected_del_cat = st.selectbox("Select Category to Delete", deletable_categories)
            if st.button(f"Delete Category '{selected_del_cat}'", type="secondary"):
                del all_rules[selected_del_cat]
                save_rules(all_rules)
                st.session_state["active_category"] = "General / Universal"
                st.success(f"Category '{selected_del_cat}' deleted!")
                st.rerun()
        else:
            st.info("No custom categories available to delete.")

# ==========================================
# TAB 3: CHECK RULES (DYNAMIC TABLE VIEW)
# ==========================================
with tab_rules:
    st.subheader("🧠 Category Rules Table & Editor")
    
    category_options = list(all_rules.keys())
    selected_rule_cat = st.selectbox(
        "🏷️ Select Category from Category Manager:", 
        category_options,
        index=category_options.index(st.session_state["active_category"]) if st.session_state["active_category"] in category_options else 0,
        key="rules_tab_category_select"
    )

    st.session_state["active_category"] = selected_rule_cat
    cat_rules = all_rules.get(selected_rule_cat, [])

    st.info(f"📋 **Viewing Category:** `{selected_rule_cat}` | **Total Check Rules Configured:** `{len(cat_rules)}`")

    if cat_rules:
        rules_table_df = pd.DataFrame({
            "Rule ID": [f"RULE-{i+1:02d}" for i in range(len(cat_rules))],
            "Equipment Category": selected_rule_cat,
            "Check Rule Description": cat_rules
        })
        st.dataframe(rules_table_df, use_container_width=True, hide_index=True)

        rules_to_delete = st.multiselect(
            "Select rule(s) to remove:",
            options=list(range(len(cat_rules))),
            format_func=lambda idx: f"[{rules_table_df.iloc[idx]['Rule ID']}] {cat_rules[idx]}"
        )

        if st.button("Delete Selected Rule(s)", type="secondary"):
            if rules_to_delete:
                for idx in sorted(rules_to_delete, reverse=True):
                    cat_rules.pop(idx)
                all_rules[selected_rule_cat] = cat_rules
                save_rules(all_rules)
                st.success("Selected rule(s) deleted!")
                st.rerun()
            else:
                st.warning("Please select at least one rule to delete.")
    else:
        st.caption(f"No check rules currently saved for **'{selected_rule_cat}'**.")

    st.markdown("---")
    with st.form("add_rule_form", clear_on_submit=True):
        new_rule_text = st.text_area(
            f"Enter new rule for [{selected_rule_cat}]:",
            placeholder=f"e.g. Check for missing sheet metal side panels in assembly drawings."
        )
        submit_rule = st.form_submit_button(f"Save Rule to '{selected_rule_cat}'", type="primary")

        if submit_rule:
            clean_rule = new_rule_text.strip()
            if clean_rule:
                all_rules[selected_rule_cat].append(clean_rule)
                save_rules(all_rules)
                st.success(f"New rule saved to **'{selected_rule_cat}'**!")
                st.rerun()
            else:
                st.error("Rule description cannot be empty.")
