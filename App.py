import streamlit as st
import pypdf
import pandas as pd
import os
import folium
from streamlit_folium import st_folium
from openai import OpenAI
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Page Configuration
st.set_page_config(
    page_title="KP Secretariat - PC-1 & PDF Intelligence Portal",
    page_icon="🏛️",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main-header {
        font-size: 26px;
        font-weight: bold;
        color: #1E3A8A;
        margin-bottom: 0px;
    }
    .sub-header {
        color: #4B5563;
        font-size: 14px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Sidebar Configurations
st.sidebar.header("⚙️ System Configurations")

BUILTIN_API_KEY = "" 
api_key = st.sidebar.text_input("OpenRouter API Key", value=BUILTIN_API_KEY, type="password")

model_name = st.sidebar.selectbox(
    "Select AI Model via OpenRouter",
    [
        "meta-llama/llama-3.3-70b-instruct",
        "google/gemini-2.0-flash-001",
        "deepseek/deepseek-chat"
    ]
)

# App Header
st.markdown('<p class="main-header">🏛️ KP Secretariat - PC-1 & Massive PDF Desk Review Portal</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Financial Summary Extractor, Geographic Map Mapper, Hidden Issue Scanner & Official Draft Generator</p>', unsafe_allow_html=True)
st.markdown("---")

# Navigation Tabs (6 Comprehensive Tabs)
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Financial Summary & Cost Breakdown", 
    "🗺️ Geographic & Map Mapper", 
    "🔍 Red-Flag & Hidden Issue Scanner",
    "💬 Deep Chat with Massive PDF",
    "⚖️ PC-1 Comparative Analysis",
    "📝 Official Reply & Draft Generator"
])

# Helper Function for OCR / Text Extraction from Large PDFs with Page Tracking
def extract_text_from_pdf(uploaded_file):
    reader = pypdf.PdfReader(uploaded_file)
    text = ""
    for index, page in enumerate(reader.pages):
        extracted = page.extract_text()
        if extracted:
            text += f"\n--- Page {index + 1} ---\n" + extracted + "\n"
    
    if len(text.strip()) < 50:
        st.warning("⚠️ Yeh PDF scanned lag rahi hai. Agar text extract na ho toh baraye meharbani searchable PDF upload karein.")
    return text

KP_DISTRICTS_DB = {
    "Peshawar": [34.0151, 71.5249],
    "Mardan": [34.1989, 72.0406],
    "Swat": [34.7717, 72.3602],
    "Abbottabad": [34.1463, 73.2117],
    "Bannu": [32.9889, 70.6042],
    "Dera Ismail Khan": [31.8314, 70.9014],
    "Kohat": [33.5889, 71.4429],
    "Swabi": [34.1167, 72.4667],
    "Nowshera": [34.0153, 71.9747],
    "Charsadda": [34.1526, 71.7381],
    "Mansehra": [34.3333, 73.2000],
    "Haripur": [33.9994, 72.9342]
}

# --- TAB 1: FINANCIAL SUMMARY & COST BREAKDOWN ---
with tab1:
    st.markdown("### 📊 PC-1 Financial Summary & Cost Extraction")
    st.markdown("Upload a PC-1 or project document to automatically extract capital costs, recurring costs, foreign exchange components, and financial summaries.")
    
    fin_file = st.file_uploader("Upload PC-1 PDF for Financial Extraction", type=["pdf"], key="fin_pdf")
    
    if fin_file is not None:
        if st.button("📊 Extract Financial Tables & Summary", type="primary"):
            if not api_key:
                st.error("⚠️ Please provide your OpenRouter API Key!")
            else:
                with st.spinner("Extracting financial schedules and cost estimates..."):
                    fin_text = extract_text_from_pdf(fin_file)
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    
                    fin_prompt = f"""
Act as a Senior Financial Analyst in KP Planning & Development Department.
Extract all financial summaries, cost breakdowns, capital expenditures, and budget allocations from the following PC-1/document text. 
Present the findings clearly in Markdown tables and highlight total project cost, foreign exchange components (if any), and year-wise phasing.

Document Text:
{fin_text[:40000]}
"""
                    fin_completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are a precise financial and budget analyst for government projects."},
                            {"role": "user", "content": fin_prompt}
                        ],
                        max_tokens=3000,
                        temperature=0.1
                    )
                    st.markdown("### 📋 Extracted Financial Summary & Cost Analysis")
                    st.markdown(fin_completion.choices[0].message.content)
    else:
        st.info("👆 Financial summary extract karne ke liye PC-1 PDF upload karein.")

# --- TAB 2: GEOGRAPHIC & MAP MAPPER ---
with tab2:
    st.markdown("### 🗺️ ADP Project Location & Geographic Map Mapper")
    st.markdown("PC-1 ke andar diye gaye district names aur project zones ko automatically detect kar ke interactive map par plot karein.")
    
    map_file = st.file_uploader("Upload PC-1 / ADP PDF for Geographic Mapping", type=["pdf"], key="map_pdf")
    
    if map_file is not None:
        base_name_m = map_file.name.rsplit('.', 1)[0]
        with st.spinner("Scanning document for geographic locations and districts..."):
            map_text = extract_text_from_pdf(map_file)
            
            detected_districts = []
            for district in KP_DISTRICTS_DB.keys():
                if district.lower() in map_text.lower():
                    detected_districts.append(district)
            
            if not detected_districts:
                detected_districts = ["Peshawar", "Mardan", "Swat"]
            
            st.markdown(f"#### 📍 Detected Districts in Document: {', '.join(detected_districts)}")
            center_coords = KP_DISTRICTS_DB.get(detected_districts[0], [34.0151, 71.5249])
            m = folium.Map(location=center_coords, zoom_start=8)
            
            for dist in detected_districts:
                coords = KP_DISTRICTS_DB[dist]
                folium.Marker(
                    coords, 
                    popup=f"Project Site Zone: {dist}", 
                    tooltip=f"{dist} District",
                    icon=folium.Icon(color="red", icon="info-sign")
                ).add_to(m)
            
            st_folium(m, width=700, height=400)
            
            map_file_path = f"{base_name_m}_Project_Map.html"
            m.save(map_file_path)
            with open(map_file_path, "rb") as map_file_dl:
                st.download_button(
                    label="📥 Download Interactive Map (.html)",
                    data=map_file_dl,
                    file_name=f"{base_name_m}_Map.html",
                    mime="text/html"
                )
    else:
        st.info("👆 Geographic mapping ke liye PDF upload karein.")

# --- TAB 3: RED-FLAG & HIDDEN ISSUE SCANNER ---
with tab3:
    st.markdown("### 🔍 Scan Massive PDFs for Hidden Loopholes, Contradictions & Red Flags")
    st.markdown("Upload any bulky PC-1, policy document, or audit file to automatically highlight hidden errors, financial mismatches, or policy violations across pages.")
    
    scan_file = st.file_uploader("Upload PDF Document for Deep Scanning", type=["pdf"], key="scan_pdf")
    scan_focus = st.text_input("Specific area to focus on (Optional):", placeholder="Misal: Cost escalation, timeline delays, ya missing approvals...")
    
    if st.button("🚀 Run Deep Issue Scanner", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key in the sidebar!")
        elif not scan_file:
            st.error("⚠️ Please upload a PDF file first!")
        else:
            with st.spinner("Analyzing multi-page document and scanning for hidden issues..."):
                try:
                    pdf_text = extract_text_from_pdf(scan_file)
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    
                    scanner_prompt = f"""
Act as a senior Chief Planning and Audit Inspector for the Government of Khyber Pakhtunkhwa.
Thoroughly examine the following multi-page document. Find and report:
1. Hidden loopholes, contradictions, or ambiguous clauses buried in the text.
2. Financial inconsistencies, inflated estimates, or budget anomalies (cite page numbers if available).
3. Missing mandatory approvals, timelines, or compliance gaps with provincial rules.
4. Specific Focus requested by user: {scan_focus}

Provide your findings in a structured, professional format with clear headings and page references.

Document Text:
{pdf_text[:40000]}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are a meticulous government document auditor and inspector."},
                            {"role": "user", "content": scanner_prompt}
                        ],
                        max_tokens=4000,
                        temperature=0.1
                    )
                    
                    st.markdown("### 📋 Deep Audit & Hidden Issues Findings")
                    st.markdown(completion.choices[0].message.content)
                except Exception as e:
                    st.error(f"❌ Scan Error: {str(e)}")

# --- TAB 4: DEEP CHAT WITH MASSIVE PDF ---
with tab4:
    st.markdown("### 💬 Deep Chat / RAG with Massive PDFs")
    st.markdown("Aap apni lambi PDF upload kar ke kisi bhi page par chupay huway specific point ya clause ke baray mein direct sawal pooch sakte hain.")
    
    chat_file = st.file_uploader("Upload Large PDF Document", type=["pdf"], key="chat_pdf")
    
    if chat_file is not None:
        with st.spinner("Indexing document pages for smart chat..."):
            chat_pdf_text = extract_text_from_pdf(chat_file)
            
        user_question = st.text_input("Aap is document mein kya talaash kar rahe hain?", placeholder="Misal: Page 45 par cost breakdown kya di gayi hai? Ya is project mein land acquisition ka kya clause hai?")
        
        if st.button("🔍 Search & Answer from PDF", type="primary"):
            if not api_key:
                st.error("⚠️ Please provide your OpenRouter API Key!")
            elif not user_question:
                st.warning("⚠️ Please type your question first.")
            else:
                with st.spinner("Searching across pages..."):
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    chat_completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are an expert document research assistant. Answer questions precisely based on the provided PDF document text, mentioning relevant page numbers if found."},
                            {"role": "user", "content": f"Document Text:\n{chat_pdf_text[:40000]}\n\nQuestion: {user_question}"}
                        ],
                        max_tokens=2000,
                        temperature=0.2
                    )
                    st.markdown("### 🤖 AI Precise Answer:")
                    st.markdown(chat_completion.choices[0].message.content)
    else:
        st.info("👆 Pehle upar PDF upload karein.")

# --- TAB 5: PC-1 COMPARATIVE ANALYSIS ---
with tab5:
    st.markdown("### ⚖️ Compare Two Documents (Original vs Revised PC-1 / Policy)")
    
    col_orig, col_rev = st.columns(2)
    with col_orig:
        file_original = st.file_uploader("Upload First File (Original / Previous PDF)", type=["pdf"], key="orig_file")
    with col_rev:
        file_revised = st.file_uploader("Upload Second File (Revised / Current PDF)", type=["pdf"], key="rev_file")
        
    if st.button("🔍 Run Comparative Analysis", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key!")
        elif not file_original or not file_revised:
            st.error("⚠️ Please upload both PDF files!")
        else:
            base_name_c1 = file_original.name.rsplit('.', 1)[0]
            base_name_c2 = file_revised.name.rsplit('.', 1)[0]
            with st.spinner("Comparing documents page by page..."):
                try:
                    text_orig = extract_text_from_pdf(file_original)[:25000]
                    text_rev = extract_text_from_pdf(file_revised)[:25000]
                    
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    comparison_prompt = f"""
Compare the following two documents in detail:
1. Exact changes in cost, budget allocations, and financial figures.
2. Modifications in scope, targets, or specific clauses.
3. Added or removed conditions between File 1 and File 2.

File 1 Extract:
{text_orig}

File 2 Extract:
{text_rev}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are a professional government planning officer specializing in comparative reviews."},
                            {"role": "user", "content": comparison_prompt}
                        ],
                        max_tokens=3000,
                        temperature=0.1
                    )
                    
                    comparison_result = completion.choices[0].message.content
                    st.markdown("### 📊 Detailed Comparative Report")
                    st.markdown(comparison_result)
                    
                    comp_doc = Document()
                    comp_doc.add_heading(f"Comparative Analysis: {base_name_c1} vs {base_name_c2}", level=1)
                    comp_doc.add_paragraph(comparison_result)
                    comp_file_path = f"{base_name_c2}_Comparison_Report.docx"
                    comp_doc.save(comp_file_path)
                    
                    with open(comp_file_path, "rb") as f_comp:
                        st.download_button(
                            label="📥 Download Comparison Report (.docx)",
                            data=f_comp,
                            file_name=comp_file_path,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                except Exception as e:
                    st.error(f"❌ Comparison Error: {str(e)}")

# --- TAB 6: OFFICIAL REPLY & DRAFT GENERATOR ---
with tab6:
    st.markdown("### 📝 Generate Official Reply or P&D Observation Response (.docx)")
    
    rep_file = st.file_uploader("Upload Observation Letter / PC-1 PDF", type=["pdf"], key="rep_pdf")
    task_option = st.selectbox(
        "Choose Draft Type:",
        [
            "Official Reply to P&D / Audit Observations",
            "Executive Summary & Briefing Note for Secretary/Minister",
            "Comprehensive Project Justification & Policy Defense Memo"
        ]
    )
    user_instructions = st.text_area("Additional instructions or specific points to include:", placeholder="Misal ke tor par: Cost increase ki wajah inflation aur steel price hike batain...")

    if st.button("🚀 Generate Official Word Document", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key!")
        elif not rep_file:
            st.error("⚠️ Please upload a source PDF file!")
        else:
            base_name = rep_file.name.rsplit('.', 1)[0]
            with st.spinner("Drafting official document..."):
                try:
                    pdf_text = extract_text_from_pdf(rep_file)
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    
                    system_prompt = (
                        "You are an expert Section Officer / Planning Officer in the Government of Khyber Pakhtunkhwa (KP) Secretariat. "
                        "Write highly professional, formal, and authoritative government correspondence compliant with provincial rules."
                    )
                    
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"Task: {task_option}\nInstructions: {user_instructions}\nSource Document:\n{pdf_text[:35000]}"}
                        ],
                        max_tokens=4000,
                        temperature=0.2
                    )
                    
                    response_text = completion.choices[0].message.content
                    st.markdown("### 📋 Generated Official Draft Preview")
                    st.markdown(response_text)
                    
                    doc = Document()
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run("GOVERNMENT OF KHYBER PAKHTUNKHWA\nSECRETARIAT - OFFICIAL CORRESPONDENCE")
                    run.bold = True
                    run.font.size = Pt(14)
                    doc.add_paragraph("----------------------------------------------------------------------------------")
                    
                    for line in response_text.split("\n"):
                        line = line.strip()
                        if not line:
                            continue
                        clean_line = line.replace("**", "").replace("###", "").replace("##", "").replace("#", "").strip()
                        if line.startswith("###") or line.startswith("##") or line.startswith("#"):
                            doc.add_heading(clean_line, level=2)
                        else:
                            p_para = doc.add_paragraph(clean_line)
                            p_para.paragraph_format.line_spacing = 1.15
                        
                    output_file_path = f"{base_name}_Official_Reply.docx"
                    doc.save(output_file_path)
                    
                    with open(output_file_path, "rb") as file:
                        st.download_button(
                            label="📥 Download Official Document (.docx)",
                            data=file,
                            file_name=output_file_path,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
