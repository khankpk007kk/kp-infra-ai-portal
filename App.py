import streamlit as st
import pypdf
import folium
from streamlit_folium import st_folium
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io # Added for in-memory file generation

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
    .stButton>button {
        width: 100%;
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

# Helper Function for OCR / Text Extraction from Large PDFs with Caching
@st.cache_data(show_spinner=False)
def extract_text_from_pdf(uploaded_file_bytes, file_name):
    reader = pypdf.PdfReader(io.BytesIO(uploaded_file_bytes))
    text = ""
    for index, page in enumerate(reader.pages):
        extracted = page.extract_text()
        if extracted:
            text += f"\n--- Page {index + 1} ---\n" + extracted + "\n"
    
    if len(text.strip()) < 50:
        st.warning("⚠️ Yeh PDF scanned lag rahi hai. Agar text extract na ho toh baraye meharbani searchable PDF upload karein.")
    return text

def truncate_text_safely(text, max_chars=40000):
    """Safely truncates text at word boundaries to avoid cutting words for AI."""
    if len(text) <= max_chars:
        return text
    truncated = text[:max_chars].rsplit(' ', 1)[0]
    return truncated + "\n\n[...Text truncated for AI context limits...]"

KP_DISTRICTS_DB = {
    "Peshawar": [34.0151, 71.5249], "Mardan": [34.1989, 72.0406], "Swat": [34.7717, 72.3602],
    "Abbottabad": [34.1463, 73.2117], "Bannu": [32.9889, 70.6042], "Dera Ismail Khan": [31.8314, 70.9014],
    "Kohat": [33.5889, 71.4429], "Swabi": [34.1167, 72.4667], "Nowshera": [34.0153, 71.9747],
    "Charsadda": [34.1526, 71.7381], "Mansehra": [34.3333, 73.2000], "Haripur": [33.9994, 72.9342]
}

# --- TAB 1: FINANCIAL SUMMARY & COST BREAKDOWN ---
with tab1:
    st.markdown("### 📊 PC-1 Financial Summary & Cost Extraction")
    fin_file = st.file_uploader("Upload PC-1 PDF for Financial Extraction", type=["pdf"], key="fin_pdf")
    
    if fin_file is not None:
        if st.button("📊 Extract Financial Tables & Summary", type="primary"):
            if not api_key: st.error("⚠️ Please provide your OpenRouter API Key!")
            else:
                with st.spinner("Extracting financial schedules..."):
                    fin_text = extract_text_from_pdf(fin_file.getvalue(), fin_file.name)
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    
                    fin_prompt = f"""
Act as a Senior Financial Analyst in KP Planning & Development Department.
Extract all financial summaries, cost breakdowns, capital expenditures, and budget allocations.
Present findings clearly in Markdown tables.

Document Text:
{truncate_text_safely(fin_text)}
"""
                    fin_completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "Precise financial analyst."},
                            {"role": "user", "content": fin_prompt}
                        ],
                        max_tokens=3000, temperature=0.1
                    )
                    st.markdown("### 📋 Extracted Financial Summary")
                    st.markdown(fin_completion.choices[0].message.content)
    else:
        st.info("👆 Financial summary extract karne ke liye PC-1 PDF upload karein.")

# --- TAB 2: GEOGRAPHIC & MAP MAPPER ---
with tab2:
    st.markdown("### 🗺️ ADP Project Location & Geographic Map Mapper")
    map_file = st.file_uploader("Upload PC-1 / ADP PDF for Geographic Mapping", type=["pdf"], key="map_pdf")
    
    if map_file is not None:
        base_name_m = map_file.name.rsplit('.', 1)[0]
        with st.spinner("Scanning document for geographic locations..."):
            map_text = extract_text_from_pdf(map_file.getvalue(), map_file.name)
            
            detected_districts = [d for d in KP_DISTRICTS_DB.keys() if d.lower() in map_text.lower()]
            if not detected_districts: detected_districts = ["Peshawar", "Mardan", "Swat"]
            
            st.markdown(f"#### 📍 Detected Districts: {', '.join(detected_districts)}")
            center_coords = KP_DISTRICTS_DB.get(detected_districts[0], [34.0151, 71.5249])
            m = folium.Map(location=center_coords, zoom_start=8)
            
            for dist in detected_districts:
                folium.Marker(KP_DISTRICTS_DB[dist], popup=f"Project Site: {dist}", tooltip=f"{dist}", icon=folium.Icon(color="red")).add_to(m)
            
            st_folium(m, width=700, height=400)
            
            # In-memory file generation for map
            map_html_str = m._repr_html_()
            st.download_button("📥 Download Interactive Map (.html)", data=map_html_str.encode('utf-8'), file_name=f"{base_name_m}_Map.html", mime="text/html")
    else:
        st.info("👆 Geographic mapping ke liye PDF upload karein.")

# --- TAB 3: RED-FLAG & HIDDEN ISSUE SCANNER ---
with tab3:
    st.markdown("### 🔍 Scan Massive PDFs for Hidden Loopholes")
    scan_file = st.file_uploader("Upload PDF Document for Deep Scanning", type=["pdf"], key="scan_pdf")
    scan_focus = st.text_input("Specific area to focus on (Optional):", placeholder="Cost escalation, timeline delays...")
    
    if st.button("🚀 Run Deep Issue Scanner", type="primary"):
        if not api_key: st.error("⚠️ Please provide API Key!")
        elif not scan_file: st.error("⚠️ Please upload a PDF!")
        else:
            with st.spinner("Analyzing multi-page document..."):
                try:
                    pdf_text = extract_text_from_pdf(scan_file.getvalue(), scan_file.name)
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    
                    scanner_prompt = f"""
Act as Chief Planning and Audit Inspector for Govt of KP.
Find hidden loopholes, financial inconsistencies, missing approvals.
Specific Focus: {scan_focus}

Document Text:
{truncate_text_safely(pdf_text)}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "Meticulous govt auditor."},
                            {"role": "user", "content": scanner_prompt}
                        ],
                        max_tokens=4000, temperature=0.1
                    )
                    st.markdown("### 📋 Deep Audit Findings")
                    st.markdown(completion.choices[0].message.content)
                except Exception as e:
                    st.error(f"❌ Scan Error: {str(e)}")

# --- TAB 4: DEEP CHAT WITH MASSIVE PDF ---
with tab4:
    st.markdown("### 💬 Deep Chat with Massive PDFs")
    chat_file = st.file_uploader("Upload Large PDF Document", type=["pdf"], key="chat_pdf")
    
    if chat_file is not None:
        chat_pdf_text = extract_text_from_pdf(chat_file.getvalue(), chat_file.name)
        user_question = st.text_input("Aap is document mein kya talaash kar rahe hain?")
        
        if st.button("🔍 Search & Answer from PDF", type="primary"):
            if not api_key: st.error("⚠️ API Key missing!")
            elif not user_question: st.warning("⚠️ Type your question.")
            else:
                with st.spinner("Searching across pages..."):
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    chat_completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "Expert document research assistant."},
                            {"role": "user", "content": f"Document Text:\n{truncate_text_safely(chat_pdf_text)}\n\nQuestion: {user_question}"}
                        ],
                        max_tokens=2000, temperature=0.2
                    )
                    st.markdown("### 🤖 AI Precise Answer:")
                    st.markdown(chat_completion.choices[0].message.content)
    else:
        st.info("👆 Pehle upar PDF upload karein.")

# --- TAB 5: PC-1 COMPARATIVE ANALYSIS ---
with tab5:
    st.markdown("### ⚖️ Compare Two Documents")
    col_orig, col_rev = st.columns(2)
    with col_orig: file_original = st.file_uploader("Upload First File", type=["pdf"], key="orig_file")
    with col_rev: file_revised = st.file_uploader("Upload Second File", type=["pdf"], key="rev_file")
        
    if st.button("🔍 Run Comparative Analysis", type="primary"):
        if not api_key: st.error("⚠️ API Key missing!")
        elif not file_original or not file_revised: st.error("⚠️ Upload both files!")
        else:
            base_name_c2 = file_revised.name.rsplit('.', 1)[0]
            with st.spinner("Comparing documents..."):
                try:
                    text_orig = extract_text_from_pdf(file_original.getvalue(), file_original.name)
                    text_rev = extract_text_from_pdf(file_revised.getvalue(), file_revised.name)
                    
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    comparison_prompt = f"""
Compare these two documents:
1. Exact changes in cost/budget.
2. Scope modifications.
3. Added/removed conditions.

File 1: {truncate_text_safely(text_orig, 20000)}
File 2: {truncate_text_safely(text_rev, 20000)}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "Planning officer."},
                            {"role": "user", "content": comparison_prompt}
                        ],
                        max_tokens=3000, temperature=0.1
                    )
                    comparison_result = completion.choices[0].message.content
                    st.markdown("### 📊 Detailed Comparative Report")
                    st.markdown(comparison_result)
                    
                    # In-memory docx generation
                    comp_doc = Document()
                    comp_doc.add_heading(f"Comparative Analysis", level=1)
                    comp_doc.add_paragraph(comparison_result)
                    
                    docx_bytes = io.BytesIO()
                    comp_doc.save(docx_bytes)
                    docx_bytes.seek(0)
                    
                    st.download_button("📥 Download Report (.docx)", data=docx_bytes, file_name=f"{base_name_c2}_Comparison.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# --- TAB 6: OFFICIAL REPLY & DRAFT GENERATOR ---
with tab6:
    st.markdown("### 📝 Generate Official Reply (.docx)")
    rep_file = st.file_uploader("Upload Observation Letter / PC-1 PDF", type=["pdf"], key="rep_pdf")
    task_option = st.selectbox("Choose Draft Type:", ["Official Reply", "Executive Summary", "Project Justification"])
    user_instructions = st.text_area("Additional instructions:")

    if st.button("🚀 Generate Official Word Document", type="primary"):
        if not api_key: st.error("⚠️ API Key missing!")
        elif not rep_file: st.error("⚠️ Upload source PDF!")
        else:
            base_name = rep_file.name.rsplit('.', 1)[0]
            with st.spinner("Drafting official document..."):
                try:
                    pdf_text = extract_text_from_pdf(rep_file.getvalue(), rep_file.name)
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    
                    system_prompt = "Expert Section Officer in KP Secretariat. Write formal govt correspondence."
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"Task: {task_option}\nInstructions: {user_instructions}\nSource Document:\n{truncate_text_safely(pdf_text)}"}
                        ],
                        max_tokens=4000, temperature=0.2
                    )
                    
                    response_text = completion.choices[0].message.content
                    st.markdown("### 📋 Generated Draft Preview")
                    st.markdown(response_text)
                    
                    doc = Document()
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run("GOVERNMENT OF KHYBER PAKHTUNKHWA\nSECRETARIAT - OFFICIAL CORRESPONDENCE")
                    run.bold = True
                    run.font.size = Pt(14)
                    
                    for line in response_text.split("\n"):
                        clean_line = line.strip().replace("**", "").replace("#", "")
                        if not clean_line: continue
                        if line.startswith("#"): doc.add_heading(clean_line, level=2)
                        else: doc.add_paragraph(clean_line)
                    
                    # In-memory docx generation
                    docx_bytes_reply = io.BytesIO()
                    doc.save(docx_bytes_reply)
                    docx_bytes_reply.seek(0)
                    
                    st.download_button("📥 Download Official Document (.docx)", data=docx_bytes_reply, file_name=f"{base_name}_Official_Reply.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
