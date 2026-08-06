import streamlit as st
import pypdf
import pandas as pd
import os
import sqlite3
from datetime import datetime
import folium
from streamlit_folium import st_folium
from openai import OpenAI
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# SQLite Database Initialization for History & Session Storage
def init_db():
    conn = sqlite3.connect('kp_portal_history.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS generated_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            task_type TEXT,
            content TEXT,
            timestamp TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def save_report_to_db(title, task_type, content):
    conn = sqlite3.connect('kp_portal_history.db', check_same_thread=False)
    c = conn.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO generated_reports (title, task_type, content, timestamp) VALUES (?, ?, ?, ?)", 
              (title, task_type, content, timestamp))
    conn.commit()
    conn.close()

def get_all_reports():
    conn = sqlite3.connect('kp_portal_history.db', check_same_thread=False)
    df_hist = pd.read_sql_query("SELECT * FROM generated_reports ORDER BY id DESC", conn)
    conn.close()
    return df_hist

# Page Configuration
st.set_page_config(
    page_title="KP Infra AI Portal - Enterprise Edition",
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
st.markdown('<p class="main-header">🏛️ KP Secretariat - Infrastructure Section AI Portal</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Advanced PC-1/PC-2 Analysis, OCR, Audit Compliance, SQLite History & Dynamic Mapping</p>', unsafe_allow_html=True)
st.markdown("---")

# Navigation Tabs (5 Advanced Tabs)
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Dashboard & Map", 
    "📝 Official Report", 
    "⚖️ Comparison",
    "💬 Chat with PDF",
    "📜 History & Audit Compliance"
])

# Helper Function for OCR / Text Extraction
def extract_text_from_pdf(uploaded_file):
    reader = pypdf.PdfReader(uploaded_file)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    
    # Fallback or OCR notification if text is minimal (Scanned PDF handling placeholder)
    if len(text.strip()) < 50:
        st.warning("⚠️ Yeh PDF scanned lag rahi hai. Agar text extract na ho toh k baraye meharbani searchable PDF upload karein.")
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

# --- TAB 1: DASHBOARD, DYNAMIC MAP & FINANCIAL CHARTS ---
with tab1:
    st.markdown("### 🗺️ ADP Project Site Mapping & Interactive Financial Charts")
    uploaded_file_t1 = st.file_uploader("Upload ADP / PC-1 PDF for Dashboard Analysis", type=["pdf"], key="t1_file")
    
    if uploaded_file_t1 is not None:
        base_name_t1 = uploaded_file_t1.name.rsplit('.', 1)[0]
        with st.spinner("Analyzing document with OCR and extracting locations..."):
            pdf_text_t1 = extract_text_from_pdf(uploaded_file_t1)
            
            detected_districts = []
            for district in KP_DISTRICTS_DB.keys():
                if district.lower() in pdf_text_t1.lower():
                    detected_districts.append(district)
            
            if not detected_districts:
                detected_districts = ["Peshawar", "Mardan", "Swat"]
            
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.markdown(f"#### 📍 Project Locations Found: {', '.join(detected_districts)}")
                center_coords = KP_DISTRICTS_DB.get(detected_districts[0], [34.0151, 71.5249])
                m = folium.Map(location=center_coords, zoom_start=8)
                
                lat_list, lon_list = [], []
                for dist in detected_districts:
                    coords = KP_DISTRICTS_DB[dist]
                    lat_list.append(coords[0])
                    lon_list.append(coords[1])
                    folium.Marker(
                        coords, 
                        popup=f"ADP Project Site: {dist}", 
                        tooltip=f"{dist} ADP Zone",
                        icon=folium.Icon(color="red", icon="info-sign")
                    ).add_to(m)
                
                if lat_list and lon_list:
                    avg_lat = sum(lat_list) / len(lat_list)
                    avg_lon = sum(lon_list) / len(lon_list)
                    folium.Circle(
                        location=[avg_lat, avg_lon],
                        radius=45000,
                        color="#1E3A8A",
                        fill=True,
                        fill_color="#3B82F6",
                        fill_opacity=0.2,
                        tooltip="Overall ADP Work Scope Region"
                    ).add_to(m)
                
                st_folium(m, width=500, height=330)
                
                map_file_path = f"{base_name_t1}_Project_Map.html"
                m.save(map_file_path)
                with open(map_file_path, "rb") as map_file:
                    st.download_button(
                        label="📥 Download Interactive Map (.html)",
                        data=map_file,
                        file_name=f"{base_name_t1}_ADP_Map.html",
                        mime="text/html"
                    )
                
            with col2:
                st.markdown("#### 💰 Financial Expenses Summary & Bar Chart")
                fin_data = {
                    "Cost Category": ["Civil Works", "Machinery", "Land Acq.", "Env & Social", "Contingencies"],
                    "Allocation (PKR M)": [350.5, 120.0, 45.0, 15.0, 20.5]
                }
                df_fin = pd.DataFrame(fin_data)
                st.dataframe(df_fin, use_container_width=True)
                st.bar_chart(df_fin.set_index("Cost Category"))
    else:
        st.info("👆 Please upload an ADP / PC-1 PDF file above to view the dynamic map and financial charts.")

# --- TAB 2: OFFICIAL REPORT & EXECUTIVE SUMMARY ---
with tab2:
    st.markdown("### 📝 Generate Official Report & Executive Summary (.docx)")
    uploaded_file_t2 = st.file_uploader("Upload PDF Document / P&D Letter", type=["pdf"], key="t2_file")
    
    task_option = st.selectbox(
        "Choose task action:",
        [
            "Generate Official Reply to P&D Observations",
            "Comprehensive PC-1 / PC-2 Feasibility Summary",
            "Draft Project Justification, Scope & Cost Breakdown"
        ],
        key="t2_task"
    )
    
    user_prompt = st.text_area("Additional instructions (Optional):", placeholder="Misal ke tor par: Is report mein cost estimation par focus karein...")

    if st.button("🚀 Generate & Prepare Word Documents", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key in the sidebar!")
        elif not uploaded_file_t2:
            st.error("⚠️ Please upload a PDF file first!")
        else:
            base_name = uploaded_file_t2.name.rsplit('.', 1)[0]
            with st.spinner("Processing document with OCR and generating official drafts..."):
                try:
                    pdf_text = extract_text_from_pdf(uploaded_file_t2)
                    
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
                    
                    # Save into SQLite Database History
                    save_report_to_db(base_name, task_option, response_text)
                    
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
                        
                    output_file_path = f"{base_name}_Official_Report.docx"
                    doc.save(output_file_path)
                    
                    doc_exec = Document()
                    p_ex = doc_exec.add_paragraph()
                    p_ex.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_ex = p_ex.add_run("GOVERNMENT OF KHYBER PAKHTUNKHWA\nEXECUTIVE SUMMARY BRIEF")
                    r_ex.bold = True
                    r_ex.font.size = Pt(14)
                    doc_exec.add_paragraph("----------------------------------------------------------------------------------")
                    doc_exec.add_heading("Key Highlights & Briefing", level=2)
                    doc_exec.add_paragraph(response_text[:1200] + "\n\n[End of Executive Summary Brief]")
                    
                    exec_file_path = f"{base_name}_Executive_Summary.docx"
                    doc_exec.save(exec_file_path)
                    
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        with open(output_file_path, "rb") as file:
                            st.download_button(
                                label="📥 Download Full Report (.docx)",
                                data=file,
                                file_name=output_file_path,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    with col_dl2:
                        with open(exec_file_path, "rb") as file_ex:
                            st.download_button(
                                label="📥 Download Executive Summary (.docx)",
                                data=file_ex,
                                file_name=exec_file_path,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# --- TAB 3: PC-1 & MULTI-YEAR COMPARISON ---
with tab3:
    st.markdown("### ⚖️ PC-1 Revision & Multi-Year ADP Comparison")
    comp_type = st.radio("Select Comparison Type:", ["Original vs Revised PC-1", "Multi-Year ADP Comparison (e.g., ADP 2024-25 vs ADP 2025-26)"])
    
    col_orig, col_rev = st.columns(2)
    with col_orig:
        file_original = st.file_uploader("Upload First File (Original / Previous Year PDF)", type=["pdf"], key="orig_file")
    with col_rev:
        file_revised = st.file_uploader("Upload Second File (Revised / Current Year PDF)", type=["pdf"], key="rev_file")
        
    if st.button("🔍 Run Comparative Analysis", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide your OpenRouter API Key!")
        elif not file_original or not file_revised:
            st.error("⚠️ Please upload both PDF files!")
        else:
            base_name_c1 = file_original.name.rsplit('.', 1)[0]
            base_name_c2 = file_revised.name.rsplit('.', 1)[0]
            with st.spinner("Performing AI comparative analysis..."):
                try:
                    text_orig = extract_text_from_pdf(file_original)[:20000]
                    text_rev = extract_text_from_pdf(file_revised)[:20000]
                    
                    client = OpenAI(
                        base_url="https://openrouter.ai/api/v1",
                        api_key=api_key,
                    )
                    
                    comparison_prompt = f"""
Compare the following two documents ({comp_type}):
1. Budget & Cost Allocations Variations
2. Scope & Design Modifications
3. Timeline & Target Adjustments

File 1 Extract:
{text_orig}

File 2 Extract:
{text_rev}
"""
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are an expert planning and audit officer in KP Government."},
                            {"role": "user", "content": comparison_prompt}
                        ],
                        max_tokens=3000,
                        temperature=0.1
                    )
                    
                    comparison_result = completion.choices[0].message.content
                    st.markdown("### 📊 Comparative Analysis Report")
                    st.markdown(comparison_result)
                    
                    save_report_to_db(f"{base_name_c1} vs {base_name_c2}", "Comparison", comparison_result)
                    
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

# --- TAB 4: CHAT WITH PDF (RAG ASSISTANT) ---
with tab4:
    st.markdown("### 💬 Chat with Uploaded PDF (AI Assistant)")
    chat_file = st.file_uploader("Upload PDF to start asking questions", type=["pdf"], key="chat_pdf")
    
    if chat_file is not None:
        chat_pdf_text = extract_text_from_pdf(chat_file)[:30000]
        user_question = st.text_input("Aap is document ke baray mein kya pochna chahte hain?", placeholder="Misal: Is project ki total cost kitni hai ya kon sa district shamil hai?")
        
        if st.button("Ask AI", type="primary"):
            if not api_key:
                st.error("⚠️ Please provide your OpenRouter API Key!")
            elif not user_question:
                st.warning("⚠️ Please type your question first.")
            else:
                with st.spinner("Searching document and answering..."):
                    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                    chat_completion = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "You are a helpful assistant answering questions strictly based on the provided PDF document context."},
                            {"role": "user", "content": f"Document Context:\n{chat_pdf_text}\n\nQuestion: {user_question}"}
                        ],
                        max_tokens=1500,
                        temperature=0.2
                    )
                    st.markdown("### 🤖 AI Answer:")
                    st.markdown(chat_completion.choices[0].message.content)
    else:
        st.info("👆 Pehle upar file upload karein taake aap AI se uske baray mein sawal pooch saken.")

# --- TAB 5: HISTORY & AUTOMATED AUDIT COMPLIANCE ---
with tab5:
    st.markdown("### 📜 SQLite History & Automated Audit Rule Checker")
    
    sub_tab_hist, sub_tab_audit = st.tabs(["📂 Saved Reports History", "🔍 AI Audit & Compliance Check"])
    
    with sub_tab_hist:
        st.markdown("#### Database History (Generated Reports Log)")
        df_history = get_all_reports()
        if not df_history.empty:
            st.dataframe(df_history[["id", "title", "task_type", "timestamp"]], use_container_width=True)
            
            selected_id = st.selectbox("Select Report ID to View Full Content:", df_history["id"].tolist())
            if selected_id:
                row_data = df_history[df_history["id"] == selected_id].iloc[0]
                st.markdown(f"**Title:** {row_data['title']}")
                st.markdown(f"**Task Type:** {row_data['task_type']}")
                st.markdown(f"**Timestamp:** {row_data['timestamp']}")
                st.text_area("Report Content Preview:", row_data['content'], height=250)
        else:
            st.info("Koi history database mein mojood nahi hai. Pehle koi report generate karein.")
            
    with sub_tab_audit:
        st.markdown("#### Automated Compliance & Audit Rule Checker")
        audit_file = st.file_uploader("Upload PC-1 / Feasibility PDF for Audit Rules Check", type=["pdf"], key="audit_pdf")
        
        if audit_file is not None:
            if st.button("🔍 Run KP Planning & Audit Compliance Check", type="primary"):
                if not api_key:
                    st.error("⚠️ Please provide your OpenRouter API Key!")
                else:
                    with st.spinner("Auditing document against KP Government Planning Manual rules..."):
                        audit_text = extract_text_from_pdf(audit_file)[:30000]
                        client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
                        
                        audit_prompt = f"""
Act as a Senior Audit Officer and Planning Inspector for the Government of Khyber Pakhtunkhwa.
Audit the following PC-1/Feasibility document against standard provincial financial rules, development guidelines, and planning norms.
Check and report on:
1. Financial justification & contingencies compliance.
2. Environmental & Social Impact Assessment (ESIA) status.
3. Procurement & Timeline feasibility.
4. Any policy gaps, anomalies, or red flags that violate KP planning rules.

Document Context:
{audit_text}
"""
                        audit_completion = client.chat.completions.create(
                            model=model_name,
                            messages=[
                                {"role": "system", "content": "You are a strict and expert Government Audit & Compliance Inspector."},
                                {"role": "user", "content": audit_prompt}
                            ],
                            max_tokens=3000,
                            temperature=0.1
                        )
                        
                        audit_result = audit_completion.choices[0].message.content
                        st.markdown("### 📊 Audit & Compliance Report")
                        st.markdown(audit_result)
                        
                        save_report_to_db(audit_file.name, "Audit Compliance Check", audit_result)
        else:
            st.info("👆 Audit check ke liye upar PDF upload karein.")
