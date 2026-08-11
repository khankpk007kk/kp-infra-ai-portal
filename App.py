import streamlit as st
import pypdf
import folium
from streamlit_folium import st_folium
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

# Page Configuration (Sidebar disabled)
st.set_page_config(
    page_title="KP Secretariat - PC-1 & PDF Intelligence Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============ CUSTOM STYLING ============
st.markdown("""
<style>
    /* --- OVERALL BOLD TEXT --- */
    body, p, span, label, div, .stMarkdown, caption {
        font-weight: 600 !important;
    }
    h1, h2, h3, h4, h5 { font-weight: 800 !important; }

    /* --- BIG BOLD TITLE --- */
    .big-title {
        font-size: 42px;
        font-weight: 900;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 4px;
    }
    .sub-header {
        color: #4B5563;
        font-size: 16px;
        font-weight: 700;
        text-align: center;
        margin-bottom: 18px;
    }
    .config-title {
        font-size: 18px;
        font-weight: 800;
        color: #1E3A8A;
        text-align: center;
        margin-top: 10px;
    }

    /* --- HIDE SIDEBAR COMPLETELY --- */
    section[data-testid="stSidebar"] { display: none; }

    /* --- SOLID NAVIGATION BAR --- */
    .stTabs [data-baseweb="tab-list"] {
        background: linear-gradient(90deg, #1E3A8A 0%, #2563EB 100%);
        border-radius: 14px;
        padding: 8px;
        gap: 6px;
        box-shadow: 0 6px 18px rgba(30,58,138,0.35);
    }
    .stTabs [data-baseweb="tab"] {
        color: #ffffff !important;
        font-weight: 700 !important;
        border-radius: 10px;
        padding: 10px 18px;
        border: none !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(255,255,255,0.18);
    }
    .stTabs [aria-selected="true"] {
        background: #ffffff !important;
        color: #1E3A8A !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.25);
    }

    /* --- BUTTONS --- */
    .stButton>button, .stDownloadButton>button {
        width: 100%;
        border-radius: 12px;
        font-weight: 700;
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white !important;
        border: none;
    }
</style>
""", unsafe_allow_html=True)

# ============ LOGO (TOP CENTER) ============
try:
    lc1, lc2, lc3 = st.columns([1, 1.2, 1])
    with lc2:
        st.image("kp_logo.png", use_container_width=True)
except:
    pass

# ============ BIG BOLD TITLE ============
st.markdown('<h1 class="big-title">🏛️ KP Secretariat – Smart PC-1 & PDF Review Portal</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Financial Summary Extractor • Geographic Map Mapper • Hidden Issue Scanner • Official Draft Generator</p>', unsafe_allow_html=True)

# ============ CENTERED SYSTEM CONFIGURATIONS (No Sidebar) ============
st.markdown('<p class="config-title">⚙️ System Configurations</p>', unsafe_allow_html=True)
cfg_l, cfg_c, cfg_r = st.columns([1, 2, 1])
with cfg_c:
    BUILTIN_API_KEY = ""
    api_key = st.text_input("🔑 OpenRouter API Key", value=BUILTIN_API_KEY, type="password", placeholder="Apni OpenRouter API key yahan enter karein...")
    model_name = st.selectbox(
        "🤖 Select AI Model via OpenRouter",
        [
            "meta-llama/llama-3.3-70b-instruct",
            "google/gemini-2.0-flash-001",
            "deepseek/deepseek-chat"
        ]
    )
st.markdown("---")

# Navigation Tabs (Solid Bar)
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Financial Summary & Cost Breakdown", 
    "🗺️ Geographic & Map Mapper", 
    "🔍 Red-Flag & Hidden Issue Scanner",
    "💬 Deep Chat with Massive PDF",
    "⚖️ PC-1 Comparative Analysis",
    "📝 Official Reply & Draft Generator"
])

# ============ HELPER FUNCTIONS ============
@st.cache_data(show_spinner=False)
def extract_text_from_pdf(uploaded_file_bytes, file_name):
    reader = pypdf.PdfReader(io.BytesIO(uploaded_file_bytes))
    text = ""
    for index, page in enumerate(reader.pages):
        extracted = page.extract_text()
        if extracted:
            text += f"\n--- Page {index + 1} ---\n" + extracted + "\n"
    if len(text.strip()) < 50:
        st.warning("⚠️ Yeh PDF scanned lag rahi hai. Baraye meharbani searchable PDF upload karein.")
    return text

def truncate_text_safely(text, max_chars=40000):
    if len(text) <= max_chars:
        return text
    truncated = text[:max_chars].rsplit(' ', 1)[0]
    return truncated + "\n\n[...Text truncated for AI context limits...]"

def clean_markdown_for_docx(text):
    text = re.sub(r'\*\*\*(.*?)\*\*\*', r'\1', text)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'_(.*?)_', r'\1', text)
    text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)
    return text

def add_logo_to_docx(doc):
    try:
        section = doc.sections[0]
        header = section.header
        header_para = header.paragraphs[0]
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header_para.add_run()
        run.add_picture("kp_logo.png", width=Inches(1.0))
    except:
        pass

KP_DISTRICTS_DB = {
    "Peshawar": [34.0151, 71.5249], "Mardan": [34.1989, 72.0406], "Swat": [34.7717, 72.3602],
    "Abbottabad": [34.1463, 73.2117], "Bannu": [32.9889, 70.6042], "Dera Ismail Khan": [31.8314, 70.9014],
    "Kohat": [33.5889, 71.4429], "Swabi": [34.1167, 72.4667], "Nowshera": [34.0153, 71.9747],
    "Charsadda": [34.1526, 71.7381], "Mansehra": [34.3333, 73.2000], "Haripur": [33.9994, 72.9342]
}

# --- TAB 1 ---
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
                        messages=[{"role": "system", "content": "Precise financial analyst."}, {"role": "user", "content": fin_prompt}],
                        max_tokens=3000, temperature=0.1
                    )
                    st.markdown("### 📋 Extracted Financial Summary")
                    st.markdown(fin_completion.choices[0].message.content)
    else:
        st.info("👆 Financial summary extract karne ke liye PC-1 PDF upload karein.")

# --- TAB 2 ---
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
            map_html_str = m._repr_html_()
            st.download_button("📥 Download Interactive Map (.html)", data=map_html_str.encode('utf-8'), file_name=f"{base_name_m}_Map.html", mime="text/html")
    else:
        st.info("👆 Geographic mapping ke liye PDF upload karein.")

# --- TAB 3 ---
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
                        messages=[{"role": "system", "content": "Meticulous govt auditor."}, {"role": "user", "content": scanner_prompt}],
                        max_tokens=4000, temperature=0.1
                    )
                    st.markdown("### 📋 Deep Audit Findings")
                    st.markdown(completion.choices[0].message.content)
                except Exception as e:
                    st.error(f"❌ Scan Error: {str(e)}")

# --- TAB 4 ---
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

# --- TAB 5 ---
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
                        messages=[{"role": "system", "content": "Planning officer."}, {"role": "user", "content": comparison_prompt}],
                        max_tokens=3000, temperature=0.1
                    )
                    comparison_result = completion.choices[0].message.content
                    st.markdown("### 📊 Detailed Comparative Report")
                    st.markdown(comparison_result)
                    
                    cleaned_result = clean_markdown_for_docx(comparison_result)
                    comp_doc = Document()
                    add_logo_to_docx(comp_doc)
                    comp_doc.add_heading("Comparative Analysis", level=1)
                    for para_text in cleaned_result.split('\n\n'):
                        if para_text.strip():
                            p = comp_doc.add_paragraph(para_text.strip())
                            p.paragraph_format.space_after = Pt(12)
                            p.paragraph_format.line_spacing = 1.5
                    docx_bytes = io.BytesIO()
                    comp_doc.save(docx_bytes)
                    docx_bytes.seek(0)
                    st.download_button("📥 Download Report (.docx)", data=docx_bytes, file_name=f"{base_name_c2}_Comparison.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# --- TAB 6 ---
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
                    
                    cleaned_response = clean_markdown_for_docx(response_text)
                    doc = Document()
                    add_logo_to_docx(doc)
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run("\nGOVERNMENT OF KHYBER PAKHTUNKHWA\nSECRETARIAT - OFFICIAL CORRESPONDENCE\n")
                    run.bold = True
                    run.font.size = Pt(14)
                    doc.add_paragraph("_" * 80)
                    for line in cleaned_response.split('\n'):
                        line = line.strip()
                        if not line:
                            doc.add_paragraph()
                            continue
                        p_para = doc.add_paragraph(line)
                        p_para.paragraph_format.space_after = Pt(12)
                        p_para.paragraph_format.line_spacing = 1.5
                    docx_bytes_reply = io.BytesIO()
                    doc.save(docx_bytes_reply)
                    docx_bytes_reply.seek(0)
                    st.download_button("📥 Download Official Document (.docx)", data=docx_bytes_reply, file_name=f"{base_name}_Official_Reply.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
