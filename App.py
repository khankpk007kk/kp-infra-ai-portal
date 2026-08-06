import streamlit as st
import pypdf
from openai import OpenAI

# Page Configuration
st.set_page_config(
    page_title="KP Infra AI Portal - Pro Edition",
    page_icon="🏛️",
    layout="wide"
)

# Custom Styling for Professional Look
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

# Built-in OpenRouter API Key
BUILTIN_API_KEY = "" 
api_key = st.sidebar.text_input("OpenRouter API Key", value=BUILTIN_API_KEY, type="password")

# OpenRouter Model Selection
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
st.markdown('<p class="sub-header">Powered by OpenRouter & Advanced Open-Source Models for PC-1 & P&D Documents</p>', unsafe_allow_html=True)
st.markdown("---")

# Main File Uploader with the exact required label
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
    
    # Text Preview with expander
    with st.expander("🔍 View Extracted Document Content Preview"):
        st.text(pdf_text[:2000] + "\n... [Content Truncated for Preview] ...")

    st.markdown("### ⚙️ Select Action Task")
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

    # Action Button
    if st.button("🚀 Get Surprised (Process & Generate Details)", type="primary"):
        if not api_key:
            st.error("⚠️ Please provide a valid OpenRouter API Key!")
        else:
            with st.spinner(f"Analyzing document using {model_name}..."):
                
                try:
                    # OpenRouter client configuration using OpenAI SDK base_url
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
                    
                    # Download Button
                    st.download_button(
                        label="📥 Download Complete Report (.txt)",
                        data=response_text,
                        file_name="KP_Infra_Official_Report.txt",
                        mime="text/plain"
                    )
                    
                except Exception as e:
                    st.error(f"❌ An error occurred during processing: {str(e)}")
