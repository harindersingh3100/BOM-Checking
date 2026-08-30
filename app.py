import streamlit as st
import pandas as pd
import pymupdf as fitz
import json
import os
from google import genai
from google.genai import types

# File path for persistent memory (rules)
RULES_FILE = "rules.json"

def load_rules():
    if os.path.exists(RULES_FILE):
        with open(RULES_FILE, "r") as f:
            return json.load(f)
    return []

def save_rules(rules):
    with open(RULES_FILE, "w") as f:
        json.dump(rules, f, indent=4)

st.set_page_config(page_title="ERP BOM vs Drawing Audit Agent", layout="wide")

st.title("🛠️ Engineering BOM & Drawing Audit Agent")
st.markdown("Interactive multi-device agent for cross-checking ERP CSV BOMs against drawing PDFs with persistent memory.")

# --- SIDEBAR: CONFIG & MEMORY ---
st.sidebar.header("1. Agent Configuration")
api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

st.sidebar.header("2. Agent Memory (Learned Rules)")
learned_rules = load_rules()

with st.sidebar.expander("View & Teach Rules"):
    new_rule = st.text_area("Teach the agent a new rule (e.g., 'Always verify Rockwool density is 64 kg/m^3 for all cassettes')")
    if st.button("Save Rule to Memory"):
        if new_rule:
            learned_rules.append(new_rule)
            save_rules(learned_rules)
            st.success("Rule saved permanently to agent memory!")
            st.rerun()
            
    if learned_rules:
        st.write("**Current Permanent Rules:**")
        for i, rule in enumerate(learned_rules, 1):
            st.text(f"{i}. {rule}")
        if st.button("Clear All Rules"):
            save_rules([])
            st.rerun()

# --- MAIN APP INTERFACE ---
if not api_key:
    st.warning("Please enter your Gemini API Key in the sidebar to proceed.")
else:
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
        st.write("**Preview of Uploaded ERP BOM:**")
        st.dataframe(erp_df.head())

        if st.button("Run Cross-Check Audit", type="primary"):
            with st.spinner("Agent is analyzing drawing, parsing BOM, and applying learned rules..."):
                try:
                    # 1. Process PDF Page to Image
                    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
                    page = doc[drawing_page_num - 1]
                    pix = page.get_pixmap(dpi=300)
                    image_path = "temp_audit_page.png"
                    pix.save(image_path)

                    # 2. Prepare API client
                    client = genai.Client(api_key=api_key)
                    
                    rules_context = "\n".join([f"- {r}" for r in learned_rules]) if learned_rules else "None"
                    erp_csv_text = erp_df.to_csv(index=False)

                    prompt = f"""
                    You are an expert mechanical engineering AI auditor.
                    Your task is to cross-check the BOM extracted from the provided engineering drawing image against the provided ERP BOM CSV data.
                    
                    Permanent Rules to Strictly Follow:
                    {rules_context}
                    
                    ERP BOM Data:
                    {erp_csv_text}
                    
                    Instructions:
                    1. Extract the BOM table items (Qty, Part Number, Description) from the drawing image.
                    2. Compare them against the ERP BOM CSV data.
                    3. Identify any mismatches in quantities, missing part numbers, extra items, or rule violations based on the permanent rules.
                    4. Return a structured JSON response containing an array of discrepancies found, or a statement if everything matches perfectly.
                    """

                    # Define the JSON Output Schema cleanly outside the call
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

                    # Execute API Request
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[
                            prompt,
                            types.Part.from_bytes(data=open(image_path, "rb").read(), mime_type="image/png")
                        ],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=output_schema
                        )
                    )

                    audit_result = json.loads(response.text)
                    
                    st.success("Audit Complete!")
                    st.write(f"**Status:** {audit_result.get('audit_status')}")
                    st.write(f"**Notes:** {audit_result.get('general_notes')}")
                    
                    disc_list = audit_result.get("discrepancies", [])
                    if disc_list:
                        st.warning(f"Found {len(disc_list)} Discrepancies/Checks:")
                        disc_df = pd.DataFrame(disc_list)
                        st.dataframe(disc_df)
                    else:
                        st.info("No discrepancies found between the drawing and the ERP BOM.")

                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
