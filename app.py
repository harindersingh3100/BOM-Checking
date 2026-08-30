import streamlit as st
import pandas as pd
import pymupdf as fitz
import json
import os
import time
import io
from google import genai
from google.genai import types

# --- PERSISTENT FILE STORAGE SETUP ---
# Stores rules outside the codebase so app/code updates never erase them
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

def generate_excel_report(audit_result, disc_df, category, model_used, total_files, total_pages, stock_length_mm):
    """Generates an in-memory Excel file with Summary and Discrepancies tabs."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
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
        
        if disc_df is not None and not disc_df.empty:
            disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)
        else:
            no_disc_df = pd.DataFrame([{"Status": "No discrepancies found. All normalized quantities and drawing specs match ERP BOM."}])
            no_disc_df.to_excel(writer, sheet_name="Discrepancies", index=False)
            
    output.seek(0)
    return output

st.set_page_config(page_title="Engineering ERP BOM & Drawing Audit Agent", layout="wide")

st.title("🛠️ Engineering BOM & Drawing Audit Agent")

# Load rules into memory
all_rules = load_rules()

# Ensure active session category state exists
if "active_category" not in st.session_state:
    st.session_state["active_category"] = list(all_rules.keys())[0]

# --- SIDEBAR CONFIGURATION ---
st.sidebar.header("⚙️ Agent Settings")
api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

selected_model = st.sidebar.selectbox(
    "Preferred Gemini Model",
    ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"],
    index=0
)

stock_length_mm = st.sidebar.number_input(
    "Standard Stock Profile Length (mm)",
    min_value=500,
    max_value=12000,
    value=3000,
    step=500,
    help="Default raw material length used to convert cut lengths into raw stock piece counts."
)

st.sidebar.markdown("---")
st.sidebar.caption(f"💾 **Rules Storage Path:** `{RULES_FILE}`")

# --- MAIN NAVIGATION TABS ---
tab_audit, tab_categories, tab_rules = st.tabs([
    "🔍 Run Cross-Check Audit", 
    "📁 Category Manager", 
    "🧠 Check Rules (Table View)"
])

# ==========================================
# TAB 1: RUN CROSS-CHECK AUDIT
# ==========================================
with tab_audit:
    st.subheader("Run Multi-PDF Drawing vs ERP BOM Cross-Check")
    
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
            csv_file = st.file_uploader("Upload ERP BOM CSV File", type=["csv"])
        with col2:
            pdf_files = st.file_uploader("Upload Drawing PDF Files", type=["pdf"], accept_multiple_files=True)

        if csv_file and pdf_files:
            erp_df = pd.read_csv(csv_file)
            st.write(f"**ERP BOM Data Preview ({len(erp_df)} items):**")
            st.dataframe(erp_df, use_container_width=True)

            if st.button("Run Full Multi-PDF Audit", type="primary"):
                with st.spinner("Processing drawing pages and auditing against ERP BOM..."):
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
                        Cross-check provided PDF drawings against the ERP BOM CSV data.
                        
                        Equipment Category: {selected_category}
                        Standard Stock Profile Length: {stock_length_mm} mm
                        Learned Rules:
                        {rules_context_str}
                        
                        CRITICAL AUDIT RULES:
                        0. STRICT TABLE COLUMN GUARD:
                           - Engineering tables have distinct columns: [ITEM / ITEM NO.] | [QTY / QUANTITY] | [PART NUMBER] | [DESCRIPTION].
                           - 'ITEM' is purely the row index. 'QTY' is the multiplier.
                           - NEVER use 'ITEM' index as a quantity multiplier!
                           - Example: Row 16 has ITEM=16, QTY=1. Multiplier is strictly 1 (NOT 16).
                        
                        1. AGGREGATE ACROSS ALL DRAWING PAGES:
                           - Consolidate all parts across all PDF drawing pages.
                        
                        2. VARIANT SUFFIXES (-A, -B): Strip suffix and sum quantities.
                        
                        3. CUT-LENGTH SUFFIXES (-1000, -1553):
                           - Cut Length (mm) x Quantity (from QTY column ONLY).
                           - Sum cut lengths for parent profile, calculate required stock pieces = ceil(Total Length / {stock_length_mm}).
                        
                        ERP BOM Data:
                        {erp_df.to_csv(index=False)}
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

                        response = client.models.generate_content(
                            model=selected_model,
                            contents=[prompt] + image_parts,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                response_schema=output_schema
                            )
                        )

                        audit_result = json.loads(response.text)
                        st.success(f"Audit Complete! Executed with model `{selected_model}`.")
                        st.write(f"**Status:** {audit_result.get('audit_status')}")
                        st.write(f"**Notes:** {audit_result.get('general_notes')}")
                        
                        disc_list = audit_result.get("discrepancies", [])
                        disc_df = pd.DataFrame(disc_list) if disc_list else None

                        if disc_df is not None and not disc_df.empty:
                            st.warning(f"Found {len(disc_list)} Discrepancies:")
                            st.dataframe(disc_df, use_container_width=True)
                        else:
                            st.info("No discrepancies found. ERP BOM perfectly matches drawings.")

                        excel_data = generate_excel_report(audit_result, disc_df, selected_category, selected_model, total_files, total_pages, stock_length_mm)
                        st.download_button(
                            label="📥 Download Excel Audit Report",
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
    st.markdown("Add new categories or delete existing categories. All added categories will instantly reflect in the **Rules Check Table**.")

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
                        st.success(f"Category **'{clean_name}'** successfully created! It is now available in the Rules Check table.")
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
            cat_rule_count = len(all_rules.get(selected_del_cat, []))
            st.caption(f"Contains **{cat_rule_count}** active rule(s).")

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
    st.markdown("Select any category from the Category Manager to view, add, or delete its check rules in a structured table view.")

    # DYNAMIC CATEGORY SELECTION LINKED DIRECTLY TO CATEGORY MANAGER
    category_options = list(all_rules.keys())
    
    selected_rule_cat = st.selectbox(
        "🏷️ Select Category from Category Manager:", 
        category_options,
        index=category_options.index(st.session_state["active_category"]) if st.session_state["active_category"] in category_options else 0,
        key="rules_tab_category_select"
    )

    # Keep active session state in sync
    st.session_state["active_category"] = selected_rule_cat

    cat_rules = all_rules.get(selected_rule_cat, [])

    # Display Category Status Badge
    st.info(f"📋 **Viewing Category:** `{selected_rule_cat}` | **Total Check Rules Configured:** `{len(cat_rules)}`")

    # Table View Rendering
    st.markdown(f"#### Active Rules Table for `{selected_rule_cat}`")
    
    if cat_rules:
        rules_table_df = pd.DataFrame({
            "Rule ID": [f"RULE-{i+1:02d}" for i in range(len(cat_rules))],
            "Equipment Category": selected_rule_cat,
            "Check Rule Description": cat_rules
        })
        st.dataframe(rules_table_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown(f"#### 🗑️ Delete Rules from `{selected_rule_cat}`")
        
        # Multi-select dropdown to remove specific rules
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
        st.caption(f"No check rules currently saved for **'{selected_rule_cat}'**. Add one below to get started.")

    st.markdown("---")
    st.markdown(f"#### ➕ Add New Check Rule to `{selected_rule_cat}`")
    
    with st.form("add_rule_form", clear_on_submit=True):
        new_rule_text = st.text_area(
            f"Enter new rule for [{selected_rule_cat}]:",
            placeholder=f"e.g. Ensure all {selected_rule_cat} structural components check for profile thickness and mounting brackets."
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

    # --- BACKUP & RESTORE SECTION ---
    st.markdown("---")
    st.markdown("### 💾 Rule Database Persistence & Backup")
    
    b_col1, b_col2 = st.columns(2)
    
    with b_col1:
        st.markdown("**Export Rules Backup**")
        json_bytes = json.dumps(all_rules, indent=4).encode('utf-8')
        st.download_button(
            label="📥 Download Rules Backup (.json)",
            data=json_bytes,
            file_name="audit_rules_backup.json",
            mime="application/json"
        )
        
    with b_col2:
        st.markdown("**Restore / Import Rules Backup**")
        uploaded_backup = st.file_uploader("Upload backup JSON file", type=["json"], key="rules_backup_uploader")
        if uploaded_backup:
            try:
                backup_data = json.load(uploaded_backup)
                if isinstance(backup_data, dict):
                    all_rules = backup_data
                    save_rules(all_rules)
                    st.success("Rules database successfully restored from backup!")
                    st.rerun()
            except Exception as e:
                st.error(f"Invalid backup file format: {e}")
