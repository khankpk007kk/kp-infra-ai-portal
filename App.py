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
    page_title="KP Infra AI Portal - Pro Enterprise Edition",
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
        "google/gemini-flash-1.5",
        "deepseek/deepseek-chat"
    ]
)

# App Header
st.markdown('<p class="main-header">🏛️ KP Secretariat - Infrastructure Section AI Portal</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Advanced PC-1/PC-2 Analysis, Interactive Mapping, Downloadable Map, Word Export & Revision Comparison</p>', unsafe_allow_html=True)
st.markdown("---")

# Navigation Tabs
tab1, tab2, tab3 = st.tabs([
    "📊 Dashboard & Interactive Map", 
    "📝 Official Reply & Word Export", 
    "⚖️ PC-1 Revision Comparison"
])

# --- TAB 1: DASHBOARD & MAP ---
with tab1:
    st.markdown("### 🗺️ Project Site Mapping & Cost Breakdown Dashboard")
    uploaded_file_t1 = st.file_uploader("Upload PC-1 / Feasibility PDF for Dashboard Analysis", type=["pdf"], key="t1_file")
    
    # Map aur Table ab sirf PDF upload hone par hi show honge
    if uploaded_file_t1 is not None:
        with st.spinner("Processing PDF for Map & Financial Summary..."):
            reader = pypdf.PdfReader(uploaded_file_t1)
            pdf_text_t1 = "".join([page.extract_text() for page in reader.pages if page.extract_text()])
            
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.markdown("#### 📍 Project Location & Site Details")
                # Interactive Map centered around Peshawar / KP
                m = folium.Map(location=[34.0151, 71.5249], zoom_start=8)
                folium.Marker(
                    [34.0151, 71.5249], 
                    popup="KP Infrastructure Project Site", 
                    tooltip="Khyber Pakhtunkhwa Secretariat"
                ).add_to(m)
                
                # Display Map in Streamlit
                st_folium(m, width=500, height=350)
                
                # Download Map Feature (.html format)
                map_file_path = "project_site_map.html"
                m.save(map_file_path)
                with open(map_file_path, "rb") as map_file:
                    st.download_button(
                        label="📥 Download Interactive Map (.html)",
                        data=map_file,
                        file_name="KP_Project_Site_Map.html",
                        mime="text/html"
                    )
                
            with col2:
                st.markdown("#### 💰 Financial Expenses Summary Table")
                data = {
                    "Cost Category": ["Civil Works", "Machinery & Equipment", "Land Acquisition", "Environmental & Social", "Contingencies", "Total Estimated Cost"],
                    "Allocated Budget (PKR Million)": [350.5, 120.0, 45.0, 15.0, 20.5, 551.0],
                    "Status": ["Approved", "Pending", "Approved", "In Review", "Approved", "Verified"]
                }
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True)
    else:
        st.info("👆 Please upload a PC-1 / Feasibility PDF file above to view the interactive map and financial expenses dashboard.")

# --- TAB 2: OFFICIAL REPLY & WORD EXPORT ---
with tab2:
    st.markdown("### 📝 Generate Official Reply & Download as Word Document (.docx)")
    uploaded_file_t2 = st.file_uploader("Upload PDF Document / P&D Observation Letter", type=["pdf"], key="t2_file")
    
    task_option = st.selectbox(
        "Choose task action:",
        [
            "Generate Official Reply to P&D Observations",
            "Comprehensive PC-1 / PC-2 Feasibility Summary",
            "Draft Project Justification, Scope & Cost Breakdown"
        ]
    )
    
    user_prompt = st.text_area("Additional instructions (Optional):", placeholder="Misal ke tor par: Is report mein cost estimation par focus karein...")

    if st.button("🚀 Generate & Prepare Word Document", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key in the sidebar!")
        elif not uploaded_file_t2:
            st.error("⚠️ Please upload a PDF file first!")
        else:
            with st.spinner("Processing document and generating official draft..."):
                try:
                    reader = pypdf.PdfReader(uploaded_file_t2)
                    pdf_text = "".join([page.extract_text() for page in reader.pages if page.extract_text()])
                    
                    client = OpenAI(
                        base_url="https://openrouter.ai/api/v1",
                        api_key=api_key,
                    )
                    
                    system_prompt = (
                        "You are an expert Chief Engineer, Planning Officer, and Infrastructure Specialist "
                        "in the Government of Khyber Pakhtunkhwa (KP) Secretariat. Write a highly professional, "
                        "accurate, and detailed official government draft compliant with provincial rules."
                    )
                    
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"Task: {task_option}\nInstructions: {user_prompt}\nDocument:\n{pdf_text[:35000]}"}
                        ],
                        max_tokens=4000,
                        temperature=0.2
                    )
                    
                    response_text = completion.choices[0].message.content
                    st.markdown("### 📋 Generated Official Draft Preview")
                    st.markdown(response_text)
                    
                    # Create Word Document (.docx) with KPK Logo and Clean Formatting
                    doc = Document()
                    
                    # Add Logo if available in root folder
                    if os.path.exists("kpk_logo.png"):
                        doc.add_picture("kpk_logo.png", width=Inches(1.2))
                    
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run("GOVERNMENT OF KHYBER PAKHTUNKHWA\nINFRASTRUCTURE SECTION - SECRETARIAT")
                    run.bold = True
                    run.font.size = Pt(14)
                    
                    doc.add_paragraph("----------------------------------------------------------------------------------")
                    
                    # Clean parsing loop to remove markdown symbols (##, **, etc.) and format properly
                    for line in response_text.split("\n"):
                        line = line.strip()
                        if not line:
                            continue
                        
                        clean_line = line.replace("**", "").replace("###", "").replace("##", "").replace("#", "").strip()
                        
                        if line.startswith("###") or line.startswith("##") or line.startswith("#"):
                            doc.add_heading(clean_line, level=2)
                        elif line.startswith("|"):
                            doc.add_paragraph(clean_line)
                        else:
                            p_para = doc.add_paragraph(clean_line)
                            p_para.paragraph_format.line_spacing = 1.15
                        
                    output_file_path = "KP_Official_Report.docx"
                    doc.save(output_file_path)
                    
                    with open(output_file_path, "rb") as file:
                        st.download_button(
                            label="📥 Download Official Report (.docx)",
                            data=file,
                            file_name="KP_Official_Report.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# --- TAB 3: PC-1 REVISION COMPARISON ---
with tab3:
    st.markdown("### ⚖️ PC-1 Revision Comparison (Original vs Revised)")
    col_orig, col_rev = st.columns(2)
    
    with col_orig:
        file_original = st.file_uploader("Upload Original PC-1 (PDF)", type=["pdf"], key="orig_file")
    with col_rev:
        file_revised = st.file_uploader("Upload Revised PC-1 (PDF)", type=["pdf"], key="rev_file")
        
    if st.button("🔍 Compare Versions & Analyze Differences", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key!")
        elif not file_original or not file_revised:
            st.error("⚠️ Please upload both Original and Revised PC-1 files!")
        else:
            with st.spinner("Comparing documents using AI..."):
                try:
                    reader_orig = pypdf.PdfReader(file_original)
                    text_orig = "".join([p.extract_text() for p in reader_orig.pages if p.extract_text()])[:20000]
                    
                    reader_rev = pypdf.PdfReader(file_revised)
                    text_rev = "".join([p.extract_text() for p in reader_rev.pages if p.extract_text()])[:20000]
                    
                    client = OpenAI(
                        base_url="https://openrouter.ai/api/v1",
                        api_key=api_key,
                    )
                    
                    comparison_prompt = f"""
Compare the following Original PC-1 and Revised PC-1 documents. Detail the key differences in:
1. Total Cost & Budget Variations
2. Scope & Design Changes
3. Timeline / Completion Period Adjustments

Original PC-1 Extract:
{text_orig}

Revised PC-1 Extract:
{text_rev}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are a professional planning and audit officer."},
                            {"role": "user", "content": comparison_prompt}
                        ],
                        max_tokens=3000,
                        temperature=0.1
                    )
                    
                    st.markdown("### 📊 Revision Comparison Report")
                    st.markdown(completion.choices[0].message.content)
                    
                except Exception as e:
                    st.error(f"❌ Comparison Error: {str(e)}")
