import streamlit as st
import pypdf
import pandas as pd
from openai import OpenAI

# Page Configuration
st.set_page_config(
    page_title="KP Infra AI Portal - Pro Dashboard Edition",
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

# Built-in API Key (Khali rakhi gayi hai taake GitHub block na kare)
BUILTIN_API_KEY = "" 
api_key = st.sidebar.text_input("OpenRouter API Key", value=BUILTIN_API_KEY, type="password")

model_name = st.sidebar.selectbox(
    "Select AI Model via OpenRouter",
    [
        "meta-llama/llama-3.3-70b-instruct",
        "deepseek/deepseek-chat",
        "anthropic/claude-3.5-sonnet",
        "google/gemini-flash-1.5"
    ]
)

# App Header
st.markdown('<p class="main-header">🏛️ KP Secretariat - Infrastructure Section AI Portal</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Advanced PC-1/PC-2 Analysis, Cost Pivot Tables & P&D Automation Dashboard</p>', unsafe_allow_html=True)
st.markdown("---")

# Main File Uploader
uploaded_file = st.file_uploader(
    label="📁 Upload your PDF to get surprised",
    type=["pdf"],
    help="Upload your PC-1, PC-2, feasibility report or P&D observation letter."
)

if uploaded_file is not None:
    with st.spinner("Reading PDF document into memory..."):
        reader = pypdf.PdfReader(uploaded_file)
        pdf_text = ""
        total_pages = len(reader.pages)
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                pdf_text += f"\n--- Page {i+1} ---\n" + text
                
    st.success(f"✨ PDF uploaded successfully! Total pages read: {total_pages}")
    
    # Tabs for different sections (Dashboard vs AI Reply Generator)
    tab1, tab2 = st.tabs(["📊 Dashboard, Cost Pivot & Map Summary", "📝 Generate Official Reply & Reports"])
    
    with tab1:
        st.markdown("### 🔍 Executive Dashboard & Financial Pivot Summary")
        st.info("Yahan AI poori PDF se project ki cost breakdown, location, aur key observations extract karke display karega.")
        
        if st.button("📈 Extract Dashboard Analytics & Cost Pivot", type="primary"):
            if not api_key:
                st.error("⚠️ Pehle sidebar mein apni OpenRouter API Key enter karein!")
            else:
                with st.spinner("Analyzing document for cost pivot and location mapping..."):
                    try:
                        client = OpenAI(
                            base_url="https://openrouter.ai/api/v1",
                            api_key=api_key,
                        )
                        
                        dashboard_prompt = f"""
You are an expert Planning Officer in KP Secretariat. Analyze the following document text and extract:
1. Project Location details (District, Tehsil, Site).
2. Cost Breakdown (Civil Works, Machinery, Land, Contingencies, Total Cost) in a structured tabular format or bullet points.
3. Key P&D / Finance Observations and Comments.

Document Content:
{pdf_text[:30000]}
"""
                        completion = client.chat.completions.create(
                            model=model_name,
                            messages=[
                                {"role": "system", "content": "You are a professional financial and infrastructure data analyst."},
                                {"role": "user", "content": dashboard_prompt}
                            ],
                            max_tokens=3000,
                            temperature=0.1
                        )
                        
                        dash_result = completion.choices[0].message.content
                        st.markdown(dash_result)
                        
                        # Sample Pivot Table view for demonstration of expenses
                        st.markdown("#### 💰 Financial Expenses Summary Table (Pivot View)")
                        data = {
                            "Cost Category": ["Civil Works", "Machinery & Equipment", "Land Acquisition", "Environmental & Social", "Contingencies", "Total Estimated Cost"],
                            "Allocated Budget (PKR Million)": [350.5, 120.0, 45.0, 15.0, 20.5, 551.0],
                            "Status": ["Approved", "Pending", "Approved", "In Review", "Approved", "Verified"]
                        }
                        df = pd.DataFrame(data)
                        st.dataframe(df, use_container_width=True)
                        
                    except Exception as e:
                        st.error(f"❌ Error generating dashboard: {str(e)}")

    with tab2:
        st.markdown("### ⚙️ Official Report & Reply Generator")
        task_option = st.selectbox(
            "Choose what you want the AI to do with this document:",
            [
                "Generate Official Reply to P&D Observations",
                "Comprehensive PC-1 / PC-2 Feasibility Summary",
                "Draft Project Justification, Scope & Cost Breakdown"
            ]
        )

        user_prompt = st.text_area(
            "Additional instructions or specific sections to focus on (Optional):", 
            placeholder="Misal ke tor par: Is report mein cost estimation aur CSR rates par focus karein..."
        )

        if st.button("🚀 Process & Generate Official Draft", type="primary"):
            if not api_key:
                st.error("⚠️ Please provide a valid OpenRouter API Key!")
            else:
                with st.spinner(f"Analyzing document using {model_name}..."):
                    try:
                        client = OpenAI(
                            base_url="https://openrouter.ai/api/v1",
                            api_key=api_key,
                        )
                        
                        system_prompt = (
                            "You are an expert Chief Engineer, Planning Officer, and Infrastructure Specialist "
                            "in the Government of Khyber Pakhtunkhwa (KP) Secretariat. Your job is to review "
                            "PC-1, PC-2 documents, feasibility reports, and P&D/Finance department observations, "
                            "and write highly professional, accurate, and detailed official government drafts "
                            "compliant with provincial development rules and CSR standards."
                        )
                        
                        user_content = f"""
Selected Task: {task_option}
Additional User Instructions: {user_prompt}

Document Content:
{pdf_text[:35000]}

Please provide a comprehensive, complete, and professionally structured response without cutting short.
"""
                        
                        completion = client.chat.completions.create(
                            model=model_name,
                            messages=[
                                {"role": "system", "content": system_prompt},
                                {"role": "user", "content": user_content}
                            ],
                            max_tokens=4000,
                            temperature=0.2
                        )
                        
                        response_text = completion.choices[0].message.content
                        
                        st.markdown("---")
                        st.markdown("### 📋 AI Generated Official Report / Draft")
                        st.markdown(response_text)
                        
                        st.download_button(
                            label="📥 Download Complete Report (.txt)",
                            data=response_text,
                            file_name="KP_Infra_Official_Report.txt",
                            mime="text/plain"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ An error occurred during processing: {str(e)}")
