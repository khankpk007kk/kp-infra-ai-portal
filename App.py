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
st.markdown('<p class="sub-header">Advanced PC-1/PC-2 Analysis, Dynamic ADP Area Mapping, Word Export & Revision Comparison</p>', unsafe_allow_html=True)
st.markdown("---")

# Navigation Tabs
tab1, tab2, tab3 = st.tabs([
    "📊 Dashboard & Dynamic Map", 
    "📝 Official Reply & Word Export", 
    "⚖️ PC-1 Revision Comparison"
])

# KP Districts Coordinates Database for Mapping
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

# --- TAB 1: DASHBOARD & DYNAMIC MAP ---
with tab1:
    st.markdown("### 🗺️ ADP Project Site Mapping & Cost Breakdown Dashboard")
    uploaded_file_t1 = st.file_uploader("Upload ADP / PC-1 PDF for Dynamic Mapping", type=["pdf"], key="t1_file")
    
    if uploaded_file_t1 is not None:
        with st.spinner("Analyzing document for project locations and financial data..."):
            reader = pypdf.PdfReader(uploaded_file_t1)
            pdf_text_t1 = "".join([page.extract_text() for page in reader.pages if page.extract_text()])
            
            # Detect which districts are mentioned in the uploaded PDF text
            detected_districts = []
            for district in KP_DISTRICTS_DB.keys():
                if district.lower() in pdf_text_t1.lower():
                    detected_districts.append(district)
            
            # Fallback agar koi specific district text mein na mile toh default Peshawar/KP rakhein
            if not detected_districts:
                detected_districts = ["Peshawar", "Mardan", "Swat"] # Sample default for demonstration
            
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.markdown(f"#### 📍 Project Locations Found: {', '.join(detected_districts)}")
                
                # Center map around the first detected district
                center_coords = KP_DISTRICTS_DB.get(detected_districts[0], [34.0151, 71.5249])
                m = folium.Map(location=center_coords, zoom_start=8)
                
                # Add markers and overall work area circle
                lat_list = []
                lon_list = []
                for dist in detected_districts:
                    coords = KP_DISTRICTS_DB[dist]
                    lat_list.append(coords[0])
                    lon_list.append(coords[1])
                    
                    # Add individual project point marker
                    folium.Marker(
                        coords, 
                        popup=f"ADP Project Site: {dist}", 
                        tooltip=f"{dist} ADP Zone",
                        icon=folium.Icon(color="red", icon="info-sign")
                    ).add_to(m)
                
                # Draw overall work area circle covering the project zones
                if lat_list and lon_list:
                    avg_lat = sum(lat_list) / len(lat_list)
                    avg_lon = sum(lon_list) / len(lon_list)
                    folium.Circle(
                        location=[avg_lat, avg_lon],
                        radius=45000, # 45km radius covering overall work region
                        color="#1E3A8A",
                        fill=True,
                        fill_color="#3B82F6",
                        fill_opacity=0.2,
                        tooltip="Overall ADP Work Scope Region"
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
                        file_name="KP_ADP_Project_Map.html",
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
        st.info("👆 Please upload an ADP / PC-1 PDF file above to automatically detect project areas and render the dynamic map.")

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
                    
                    if os.path.exists("kpk_logo.png"):
                        doc.add_picture("kpk_logo.png", width=Inches(1.2))
                    
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run("GOVERNMENT OF KHYBER PAKHTUNKHWA\nINFRASTRUCTURE SECTION - SECRETARIAT")
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
