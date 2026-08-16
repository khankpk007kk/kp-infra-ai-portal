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
import logging

# ============ PAGE CONFIG ============
st.set_page_config(
    page_title="KP Secretariat - PC-1 & PDF Intelligence Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============ SESSION STATE INIT ============
if "extracted_texts" not in st.session_state:
    st.session_state.extracted_texts = {}  # cache for uploaded files
if "map_html" not in st.session_state:
    st.session_state.map_html = None
if "map_filename" not in st.session_state:
    st.session_state.map_filename = "map"

# ============ CUSTOM STYLING + ANIMATIONS ============
st.markdown("""
<style>
    /* --- ANIMATED BACKGROUND (Moving Light Effect) --- */
    .stApp {
        background: linear-gradient(-45deg, #0a192f, #1e3a8a, #0f172a, #1e3a8a) !important;
        background-size: 400% 400% !important;
        animation: gradientBG 18s ease infinite !important;
    }
    @keyframes gradientBG {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* --- MAIN CONTAINER TRANSPARENCY FOR READABILITY --- */
    .main .block-container {
        background: rgba(255, 255, 255, 0.92);
        backdrop-filter: blur(10px);
        border-radius: 30px;
        padding: 2rem 2rem 2rem 2rem;
        margin-top: 10px;
        margin-bottom: 20px;
        box-shadow: 0 20px 50px rgba(0,0,0,0.5);
    }

    /* --- LOGO WRAPPER WITH 12-PIECE ANIMATED RING --- */
    .logo-wrapper {
        position: relative;
        display: inline-block;
        padding: 12px;
        border-radius: 50%;
    }
    /* 12-piece spinning ring */
    .logo-wrapper::before {
        content: '';
        position: absolute;
        top: -8px; left: -8px; right: -8px; bottom: -8px;
        border-radius: 50%;
        background: repeating-conic-gradient(
            from 0deg,
            #fbbf24 0deg 15deg,
            #1e3a8a 15deg 30deg,
            #fbbf24 30deg 45deg,
            #1e3a8a 45deg 60deg,
            #fbbf24 60deg 75deg,
            #1e3a8a 75deg 90deg,
            #fbbf24 90deg 105deg,
            #1e3a8a 105deg 120deg,
            #fbbf24 120deg 135deg,
            #1e3a8a 135deg 150deg,
            #fbbf24 150deg 165deg,
            #1e3a8a 165deg 180deg,
            #fbbf24 180deg 195deg,
            #1e3a8a 195deg 210deg,
            #fbbf24 210deg 225deg,
            #1e3a8a 225deg 240deg,
            #fbbf24 240deg 255deg,
            #1e3a8a 255deg 270deg,
            #fbbf24 270deg 285deg,
            #1e3a8a 285deg 300deg,
            #fbbf24 300deg 315deg,
            #1e3a8a 315deg 330deg,
            #fbbf24 330deg 345deg,
            #1e3a8a 345deg 360deg
        );
        animation: spinRing 6s linear infinite;
        mask: radial-gradient(farthest-side, transparent calc(100% - 5px), #fff calc(100% - 3px));
        -webkit-mask: radial-gradient(farthest-side, transparent calc(100% - 5px), #fff calc(100% - 3px));
        z-index: 10;
        pointer-events: none;
    }
    /* Glowing pulse behind logo */
    .logo-wrapper::after {
        content: '';
        position: absolute;
        top: -20px; left: -20px; right: -20px; bottom: -20px;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(251,191,36,0.25) 0%, transparent 70%);
        animation: pulseGlow 2.5s ease-in-out infinite alternate;
        z-index: 5;
        pointer-events: none;
    }
    @keyframes spinRing {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    @keyframes pulseGlow {
        0% { transform: scale(0.95); opacity: 0.5; }
        100% { transform: scale(1.2); opacity: 1; }
    }

    /* --- OVERALL BOLD TEXT --- */
    body, p, span, label, div, .stMarkdown, caption {
        font-weight: 600 !important;
    }
    h1, h2, h3, h4, h5 { font-weight: 800 !important; }

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

    /* --- HIDE SIDEBAR --- */
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
    /* Override for map generate button specifically */
    .stButton>button[kind="secondary"] { 
        background: #1E3A8A; 
    }

    /* --- WARNING BOX --- */
    .stAlert {
        border-radius: 12px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ============ DATA PRIVACY WARNING ============
st.warning("⚠️ **Data Privacy Notice:** Your document text is sent to OpenRouter (third-party AI). Do NOT upload classified or highly sensitive government documents.")

# ============ LOGO (TOP CENTER) WITH 12-PIECE ANIMATION ============
lc1, lc2, lc3 = st.columns([1, 1.2, 1])
with lc2:
    # Wrapping the image in a div with the animated ring class
    st.markdown('<div class="logo-wrapper" style="display: flex; justify-content: center;">', unsafe_allow_html=True)
    try:
        st.image("kp_logo.png", use_container_width=True)
    except:
        st.markdown("🏛️ **KP Logo** (Place `kp_logo.png` here)")
    st.markdown('</div>', unsafe_allow_html=True)

# ============ BIG BOLD TITLE ============
st.markdown('<h1 class="big-title">🏛️ KP Secretariat – Smart PC-1 & PDF Review Portal</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Financial Summary Extractor • Geographic Map Mapper • Hidden Issue Scanner • Official Draft Generator</p>', unsafe_allow_html=True)

# ============ CENTERED SYSTEM CONFIGURATIONS ============
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

# ============ NAVIGATION TABS ============
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
def validate_and_extract_pdf(uploaded_bytes, file_name):
    """Validates PDF and extracts text with error handling."""
    try:
        reader = pypdf.PdfReader(io.BytesIO(uploaded_bytes))
        if len(reader.pages) == 0:
            raise ValueError("PDF has no pages.")
        text = ""
        for index, page in enumerate(reader.pages):
            extracted = page.extract_text()
            if extracted:
                text += f"\n--- Page {index + 1} ---\n" + extracted + "\n"
        if len(text.strip()) < 50:
            raise ValueError("Scanned PDF or no text found. Please upload a searchable PDF.")
        return text
    except Exception as e:
        raise ValueError(f"PDF Processing Error: {str(e)}")

def safe_truncate_text(text, max_chars=35000):
    """Truncate safely, ensuring we don't break words if possible."""
    if len(text) <= max_chars:
        return text
    # Try to cut at last space within limit
    truncated = text[:max_chars]
    last_space = truncated.rfind(' ')
    if last_space > 0:
        return truncated[:last_space] + "\n\n[...Truncated for AI context...]"
    return truncated + "\n\n[...Truncated for AI context...]"

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

# ============ ENHANCED DISTRICT DB WITH FUZZY ALIASES ============
KP_DISTRICTS_DB = {
    "Peshawar": [34.0151, 71.5249], "Mardan": [34.1989, 72.0406], "Swat": [34.7717, 72.3602],
    "Abbottabad": [34.1463, 73.2117], "Bannu": [32.9889, 70.6042], "Dera Ismail Khan": [31.8314, 70.9014],
    "Kohat": [33.5889, 71.4429], "Swabi": [34.1167, 72.4667], "Nowshera": [34.0153, 71.9747],
    "Charsadda": [34.1526, 71.7381], "Mansehra": [34.3333, 73.2000], "Haripur": [33.9994, 72.9342]
}
# Fuzzy aliases for better detection
FUZZY_ALIASES = {
    "di khan": "Dera Ismail Khan", "d.i. khan": "Dera Ismail Khan", "pesh": "Peshawar",
    "mard": "Mardan", "swabi": "Swabi", "nowshera": "Nowshera", "charsadda": "Charsadda",
    "mansehra": "Mansehra", "haripur": "Haripur", "kohat": "Kohat", "bannu": "Bannu",
    "abbottabad": "Abbottabad", "swat": "Swat"
}

# ============ TAB 1: FINANCIAL EXTRACTION ============
with tab1:
    st.markdown("### 📊 PC-1 Financial Summary & Cost Extraction")
    fin_file = st.file_uploader("Upload PC-1 PDF for Financial Extraction", type=["pdf"], key="fin_pdf")
    if fin_file is not None:
        if st.button("📊 Extract Financial Tables & Summary", type="primary"):
            if not api_key:
                st.error("⚠️ Please provide your OpenRouter API Key!")
            else:
                try:
                    with st.status("📄 Extracting financial schedules...", expanded=True) as status:
                        # Extract with validation
                        fin_text = validate_and_extract_pdf(fin_file.getvalue(), fin_file.name)
                        st.session_state.extracted_texts[fin_file.name] = fin_text
                        
                        client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                        status.update(label="🤖 Sending to AI for financial analysis...")
                        fin_prompt = f"""
Act as a Senior Financial Analyst in KP Planning & Development Department.
Extract all financial summaries, cost breakdowns, capital expenditures, and budget allocations.
Present findings clearly in Markdown tables.

Document Text:
{safe_truncate_text(fin_text)}
"""
                        fin_completion = client.chat.completions.create(
                            model=model_name,
                            messages=[{"role": "system", "content": "Precise financial analyst."}, {"role": "user", "content": fin_prompt}],
                            max_tokens=3000, temperature=0.1
                        )
                        status.update(label="✅ Extraction Complete!", state="complete")
                        st.markdown("### 📋 Extracted Financial Summary")
                        st.markdown(fin_completion.choices[0].message.content)
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    else:
        st.info("👆 Financial summary extract karne ke liye PC-1 PDF upload karein.")

# ============ TAB 2: GEOGRAPHIC MAPPER (WITH GENERATE BUTTON) ============
with tab2:
    st.markdown("### 🗺️ ADP Project Location & Geographic Map Mapper")
    map_file = st.file_uploader("Upload PC-1 / ADP PDF for Geographic Mapping", type=["pdf"], key="map_pdf")
    
    if map_file is not None:
        base_name_m = map_file.name.rsplit('.', 1)[0]
        
        # Detect districts on file upload
        try:
            map_text = validate_and_extract_pdf(map_file.getvalue(), map_file.name)
            st.session_state.extracted_texts[map_file.name] = map_text
            
            # Advanced detection with fuzzy matching
            detected_districts = []
            map_text_lower = map_text.lower()
            for dist in KP_DISTRICTS_DB.keys():
                if dist.lower() in map_text_lower:
                    detected_districts.append(dist)
            # Fuzzy check
            for alias, real_name in FUZZY_ALIASES.items():
                if alias in map_text_lower and real_name not in detected_districts:
                    detected_districts.append(real_name)
            
            if not detected_districts:
                detected_districts = ["Peshawar", "Mardan", "Swat"]  # Default
            st.session_state['detected_districts'] = detected_districts
            st.session_state['map_base_name'] = base_name_m
            st.success(f"✅ Detected Districts: {', '.join(detected_districts)}")
        except Exception as e:
            st.error(f"❌ PDF Read Error: {str(e)}")
            st.stop()

        # Generate Map Button
        if st.button("🗺️ Generate Interactive Map", type="primary"):
            with st.spinner("🗺️ Rendering map..."):
                detected = st.session_state.get('detected_districts', ["Peshawar"])
                center_coords = KP_DISTRICTS_DB.get(detected[0], [34.0151, 71.5249])
                m = folium.Map(location=center_coords, zoom_start=8)
                for dist in detected:
                    folium.Marker(
                        KP_DISTRICTS_DB[dist], 
                        popup=f"📍 Project Site: {dist}", 
                        tooltip=dist, 
                        icon=folium.Icon(color="red", icon="info-sign")
                    ).add_to(m)
                # Save map HTML to session state
                st.session_state.map_html = m._repr_html_()
                st.session_state.map_filename = st.session_state.get('map_base_name', 'map')
                st.success("✅ Map generated successfully!")

        # Display map if available in session
        if st.session_state.map_html:
            st.components.v1.html(st.session_state.map_html, height=450)
            st.download_button(
                "📥 Download Interactive Map (.html)", 
                data=st.session_state.map_html.encode('utf-8'), 
                file_name=f"{st.session_state.map_filename}_Map.html", 
                mime="text/html"
            )
    else:
        st.info("👆 Geographic mapping ke liye PDF upload karein.")

# ============ TAB 3: DEEP SCANNER ============
with tab3:
    st.markdown("### 🔍 Scan Massive PDFs for Hidden Loopholes")
    scan_file = st.file_uploader("Upload PDF Document for Deep Scanning", type=["pdf"], key="scan_pdf")
    scan_focus = st.text_input("Specific area to focus on (Optional):", placeholder="Cost escalation, timeline delays...")
    
    if st.button("🚀 Run Deep Issue Scanner", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide API Key!")
        elif not scan_file:
            st.error("⚠️ Please upload a PDF!")
        else:
            try:
                with st.status("🔎 Scanning document for issues...", expanded=True) as status:
                    pdf_text = validate_and_extract_pdf(scan_file.getvalue(), scan_file.name)
                    st.session_state.extracted_texts[scan_file.name] = pdf_text
                    status.update(label="🧠 Analyzing with AI...")
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    scanner_prompt = f"""
Act as Chief Planning and Audit Inspector for Govt of KP.
Find hidden loopholes, financial inconsistencies, missing approvals.
Specific Focus: {scan_focus}

Document Text:
{safe_truncate_text(pdf_text)}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[{"role": "system", "content": "Meticulous govt auditor."}, {"role": "user", "content": scanner_prompt}],
                        max_tokens=4000, temperature=0.1
                    )
                    status.update(label="✅ Scan Complete", state="complete")
                    st.markdown("### 📋 Deep Audit Findings")
                    st.markdown(completion.choices[0].message.content)
            except Exception as e:
                st.error(f"❌ Scan Error: {str(e)}")

# ============ TAB 4: DEEP CHAT ============
with tab4:
    st.markdown("### 💬 Deep Chat with Massive PDFs")
    chat_file = st.file_uploader("Upload Large PDF Document", type=["pdf"], key="chat_pdf")
    if chat_file is not None:
        try:
            # Extract and cache
            if chat_file.name not in st.session_state.extracted_texts:
                chat_pdf_text = validate_and_extract_pdf(chat_file.getvalue(), chat_file.name)
                st.session_state.extracted_texts[chat_file.name] = chat_pdf_text
            else:
                chat_pdf_text = st.session_state.extracted_texts[chat_file.name]
            
            user_question = st.text_input("💬 Aap is document mein kya talaash kar rahe hain?")
            if st.button("🔍 Search & Answer from PDF", type="primary"):
                if not api_key:
                    st.error("⚠️ API Key missing!")
                elif not user_question:
                    st.warning("⚠️ Type your question.")
                else:
                    try:
                        with st.spinner("🔍 Searching across pages..."):
                            client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                            chat_completion = client.chat.completions.create(
                                model=model_name,
                                messages=[
                                    {"role": "system", "content": "Expert document research assistant."},
                                    {"role": "user", "content": f"Document Text:\n{safe_truncate_text(chat_pdf_text)}\n\nQuestion: {user_question}"}
                                ],
                                max_tokens=2000, temperature=0.2
                            )
                            st.markdown("### 🤖 AI Precise Answer:")
                            st.markdown(chat_completion.choices[0].message.content)
                    except Exception as e:
                        st.error(f"❌ Chat Error: {str(e)}")
        except Exception as e:
            st.error(f"❌ PDF Load Error: {str(e)}")
    else:
        st.info("👆 Pehle upar PDF upload karein.")

# ============ TAB 5: COMPARATIVE ANALYSIS ============
with tab5:
    st.markdown("### ⚖️ Compare Two Documents")
    col_orig, col_rev = st.columns(2)
    with col_orig:
        file_original = st.file_uploader("Upload First File", type=["pdf"], key="orig_file")
    with col_rev:
        file_revised = st.file_uploader("Upload Second File", type=["pdf"], key="rev_file")
    
    if st.button("🔍 Run Comparative Analysis", type="primary"):
        if not api_key:
            st.error("⚠️ API Key missing!")
        elif not file_original or not file_revised:
            st.error("⚠️ Upload both files!")
        else:
            base_name_c2 = file_revised.name.rsplit('.', 1)[0]
            try:
                with st.status("⚖️ Comparing documents...", expanded=True) as status:
                    status.update(label="📄 Extracting text from both files...")
                    text_orig = validate_and_extract_pdf(file_original.getvalue(), file_original.name)
                    text_rev = validate_and_extract_pdf(file_revised.getvalue(), file_revised.name)
                    st.session_state.extracted_texts[file_original.name] = text_orig
                    st.session_state.extracted_texts[file_revised.name] = text_rev
                    
                    status.update(label="🧠 AI generating comparison...")
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    comparison_prompt = f"""
Compare these two documents:
1. Exact changes in cost/budget.
2. Scope modifications.
3. Added/removed conditions.

File 1: {safe_truncate_text(text_orig, 20000)}
File 2: {safe_truncate_text(text_rev, 20000)}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[{"role": "system", "content": "Planning officer."}, {"role": "user", "content": comparison_prompt}],
                        max_tokens=3000, temperature=0.1
                    )
                    comparison_result = completion.choices[0].message.content
                    status.update(label="✅ Comparison Ready", state="complete")
                    
                    st.markdown("### 📊 Detailed Comparative Report")
                    st.markdown(comparison_result)
                    
                    # DOCX generation
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

# ============ TAB 6: DRAFT GENERATOR ============
with tab6:
    st.markdown("### 📝 Generate Official Reply (.docx)")
    rep_file = st.file_uploader("Upload Observation Letter / PC-1 PDF", type=["pdf"], key="rep_pdf")
    task_option = st.selectbox("Choose Draft Type:", ["Official Reply", "Executive Summary", "Project Justification"])
    user_instructions = st.text_area("📝 Additional instructions:")
    
    if st.button("🚀 Generate Official Word Document", type="primary"):
        if not api_key:
            st.error("⚠️ API Key missing!")
        elif not rep_file:
            st.error("⚠️ Upload source PDF!")
        else:
            base_name = rep_file.name.rsplit('.', 1)[0]
            try:
                with st.status("📝 Drafting official document...", expanded=True) as status:
                    status.update(label="📄 Extracting source text...")
                    pdf_text = validate_and_extract_pdf(rep_file.getvalue(), rep_file.name)
                    st.session_state.extracted_texts[rep_file.name] = pdf_text
                    
                    status.update(label="🧠 Generating draft with AI...")
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    system_prompt = "Expert Section Officer in KP Secretariat. Write formal govt correspondence."
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"Task: {task_option}\nInstructions: {user_instructions}\nSource Document:\n{safe_truncate_text(pdf_text)}"}
                        ],
                        max_tokens=4000, temperature=0.2
                    )
                    response_text = completion.choices[0].message.content
                    status.update(label="✅ Draft generated", state="complete")
                    
                    st.markdown("### 📋 Generated Draft Preview")
                    st.markdown(response_text)
                    
                    # DOCX generation
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
