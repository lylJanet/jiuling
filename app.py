import streamlit as st
import pandas as pd
from openai import OpenAI
import io
import time

# -----------------------------------------------------------------------------
# 安全导入可选依赖
# -----------------------------------------------------------------------------
try:
    import anthropic
except ImportError:
    anthropic = None

try:
    import docx
except ImportError:
    docx = None

# -----------------------------------------------------------------------------
# 1. 页面配置与 CSS 美化
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="九灵AI 智能办公平台",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义 CSS 样式
st.markdown("""
<style>
    /* 全局字体与背景 */
    .stApp {
        background-color: #f8f9fa;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* 标题样式 */
    h1, h2, h3 {
        color: #0d47a1;
        font-weight: 700;
    }
    
    /* 调整主区域 padding */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* 按钮美化 */
    .stButton > button {
        background-color: #1e88e5;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #1565c0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    /* 报告卡片样式 */
    .report-card {
        background: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border: 1px solid #e0e0e0;
        margin-bottom: 20px;
    }
    
    /* 隐藏默认 Footer 和 Hamburger Menu */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Sidebar 美化 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
    }
    
    /* 自定义 Footer */
    .custom-footer {
        text-align: center;
        color: #757575;
        padding: 20px;
        font-size: 0.85rem;
        border-top: 1px solid #e0e0e0;
        margin-top: 40px;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 核心逻辑函数
# -----------------------------------------------------------------------------
def process_multiple_files(files):
    """处理多个上传的文件，提取内容并汇总"""
    combined_text = ""
    file_summaries = []
    data_types = set()
    
    for file in files:
        combined_text += f"\n\n### 【文件: {file.name}】\n"
        try:
            if file.name.endswith('.csv') or file.name.endswith('.xlsx'):
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file)
                
                data_types.add("dataframe")
                file_summaries.append({
                    "name": file.name,
                    "type": "dataframe",
                    "shape": df.shape,
                    "preview": df.head(5)
                })
                
                combined_text += f"- **数据维度**: {df.shape[0]}行 x {df.shape[1]}列\n"
                combined_text += f"- **列名**: {list(df.columns)}\n"
                combined_text += f"- **数据预览 (前5行)**:\n{df.head(5).to_markdown(index=False)}\n"
                
            elif file.name.endswith('.docx'):
                if docx is None:
                    combined_text += "错误: 环境缺少 python-docx 依赖，无法读取。\n"
                    continue
                doc = docx.Document(file)
                full_text = '\n'.join([para.text for para in doc.paragraphs])
                
                data_types.add("text")
                file_summaries.append({
                    "name": file.name,
                    "type": "text",
                    "preview": full_text[:300] + "..." if len(full_text) > 300 else full_text
                })
                
                # 限制单个文档传入大模型的长度，避免超长
                combined_text += full_text[:20000] + "\n"
        except Exception as e:
            combined_text += f"读取文件失败: {e}\n"
            
    # 确定整体数据类型，以便提供默认的分析提示词
    if "dataframe" in data_types and "text" in data_types:
        overall_type = "mixed"
    elif "dataframe" in data_types:
        overall_type = "dataframe"
    else:
        overall_type = "text"
        
    return combined_text, overall_type, file_summaries

def call_llm(prompt, provider, key, base_url=None):
    try:
        if provider == "OpenAI (GPT-4o)":
            client_kwargs = {"api_key": key}
            if base_url:
                client_kwargs["base_url"] = base_url
            client = OpenAI(**client_kwargs)
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "system", "content": "你是一个资深的商业分析师和文档专家。"},
                          {"role": "user", "content": prompt}],
                temperature=0.7
            )
            return response.choices[0].message.content
            
        elif provider == "Anthropic (Claude 3.5)":
            if anthropic is None:
                return "当前环境未安装 anthropic 依赖，请先安装 requirements.txt。"
            client_kwargs = {"api_key": key}
            if base_url:
                client_kwargs["base_url"] = base_url
            client = anthropic.Anthropic(**client_kwargs)
            response = client.messages.create(
                model="claude-3-5-sonnet-20240620",
                max_tokens=2000,
                temperature=0.7,
                system="你是一个资深的商业分析师和文档专家。",
                messages=[{"role": "user", "content": prompt}]
            )
            return response.content[0].text
        
        elif provider == "Kimi (Moonshot AI)":
            client_kwargs = {"api_key": key, "base_url": base_url or "https://api.moonshot.cn/v1"}
            client = OpenAI(**client_kwargs)
            response = client.chat.completions.create(
                model="moonshot-v1-32k",
                messages=[{"role": "system", "content": "你是一个资深的商业分析师和文档专家。"},
                          {"role": "user", "content": prompt}],
                temperature=0.7
            )
            return response.choices[0].message.content
    except Exception as e:
        return f"调用大模型失败: {str(e)}"

def get_analysis_report(combined_content, query, provider, key, base_url=None):
    prompt = f"""
你是一个资深的商业分析师、数据科学家和文档专家。请根据以下我提供的【一个或多个文件的数据概览/内容片段】，回答用户的问题并生成一份详细的中文分析报告。

{combined_content}

---
### 用户分析指令
"{query}"

### 报告要求
请生成一份结构化的 Markdown 报告，排版清晰美观。请包含以下部分：
1. **全局洞察**: 综合所有提供的文件内容，给出核心摘要或数据概况。
2. **关键发现**: 提取文档中的重要观点，或分析数据中的趋势与异常值。
3. **交叉分析**: 若提供了多个文件，请寻找并分析不同文件之间的关联性与共性。
4. **总结与建议**: 基于分析结果，给出具有可落地性的业务建议或后续行动指南。
请直接输出 Markdown 格式的内容，不要包含无关的寒暄。
"""
    return call_llm(prompt, provider, key, base_url)

# -----------------------------------------------------------------------------
# 3. 页面布局实现 (科技感 Demo 风格)
# -----------------------------------------------------------------------------

# --- Sidebar 配置面板 ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/bot.png", width=60)
    st.title("配置面板")
    st.markdown("---")
    
    st.markdown("**模型设置**")
    api_provider = st.selectbox("选择模型引擎", ["Kimi (Moonshot AI)", "OpenAI (GPT-4o)", "Anthropic (Claude 3.5)"])
    
    default_api_key = ""
    default_base_url = ""
    if api_provider == "Kimi (Moonshot AI)":
        default_base_url = "https://api.moonshot.cn/v1"

    api_key = st.text_input("API Key", value=default_api_key, type="password")
    
    with st.expander("高级设置（可选）", expanded=False):
        base_url = st.text_input("Base URL", value=default_base_url)

    st.markdown("---")
    st.info("基于 Kimi K2.5 模型，提供全流程文档与数据智能分析")

# --- 主区域 ---

# 顶部 Header
col_header_1, col_header_2 = st.columns([1, 6])
with col_header_1:
    st.markdown("# 🚀")
with col_header_2:
    st.title("九灵AI 智能办公平台 - 数据与文档分析助手")
    st.markdown("##### Professional AI Office Demo | Data & Document Intelligence")

st.markdown("---")

# 主内容分栏布局
col_left, col_right = st.columns([1.25, 1], gap="large")

with col_left:
    st.subheader("📋 报告预览")
    if 'report' in st.session_state:
        st.markdown('<div class="report-card">', unsafe_allow_html=True)
        st.markdown("### 🧠 智能分析结果")
        with st.expander("完整报告", expanded=True):
            st.markdown(st.session_state['report'], unsafe_allow_html=True)
        st.download_button(
            label="下载报告 (Markdown)",
            data=st.session_state['report'],
            file_name="jiuling_multi_document_analysis.md",
            mime="text/markdown",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("暂无报告，请在右侧上传文件并生成分析报告。")

with col_right:
    st.subheader("📂 多文件上传与分析")
    # 开启 accept_multiple_files=True 支持多文件
    uploaded_files = st.file_uploader(
        "支持 Excel, CSV, Word - 可同时选择多个文件", 
        type=['xlsx', 'csv', 'docx'], 
        accept_multiple_files=True
    )
    
    combined_content = None
    data_type = None
    
    if uploaded_files:
        combined_content, data_type, summaries = process_multiple_files(uploaded_files)
        if combined_content:
            st.success(f"成功读取 {len(uploaded_files)} 个文件！")
            
            # 多文件内容预览
            with st.expander("👀 文件内容预览", expanded=False):
                for idx, summary in enumerate(summaries):
                    st.markdown(f"**📄 {summary['name']}**")
                    if summary['type'] == 'dataframe':
                        st.caption(f"数据形状: {summary['shape']}")
                        st.dataframe(summary['preview'], use_container_width=True)
                    else:
                        st.text_area("文档片段", value=summary['preview'], height=120, disabled=True, key=f"prev_{idx}")
                    st.markdown("---")
    
    if combined_content:
        # 根据上传的文件类型自动调整默认指令
        if data_type == "dataframe":
            default_query = "请对比分析这些数据表，重点关注各维度的分布情况和异常值，并给出业务改进建议。"
        elif data_type == "text":
            default_query = "请总结这些文档的核心观点，提取关键信息，并分析它们之间的关联性。"
        else:
            default_query = "请综合分析上传的数据表和文本文档，提取核心信息并给出交叉分析的结论。"
            
        user_query = st.text_area("请输入分析指令", value=default_query, height=120)
        
        if st.button("✨ 生成多文档联合分析报告", type="primary", use_container_width=True):
            if not api_key:
                st.error("请先在左侧配置 API Key")
            else:
                with st.spinner(f"🤖 AI 正在深度阅读并联合分析 {len(uploaded_files)} 份文件，请稍候..."):
                    report = get_analysis_report(combined_content, user_query, api_provider, api_key, base_url)
                if isinstance(report, str) and ("失败" in report or "未安装" in report):
                    st.error(report)
                else:
                    st.session_state['report'] = report
                    st.success("多文档联合报告生成成功！")
                    st.rerun()
    else:
        st.info("请先上传一个或多个文件，然后输入分析指令。")

st.markdown("---")
current_time = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
st.markdown(
    f"""
    <div class="custom-footer">
        九灵AI © 2026 | 专注办公场景AI数字员工<br>
        当前时间：{current_time} | Version 1.3.0 (多文档增强版)
    </div>
    """, 
    unsafe_allow_html=True
)
