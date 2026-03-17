import streamlit as st
import pandas as pd
from openai import OpenAI
import io
import time
import numpy as np
import streamlit_authenticator as stauth
from streamlit_authenticator.utilities.hasher import Hasher

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
# 登录功能（加在最前面）
# -----------------------------------------------------------------------------
# 用户凭证（简单硬编码，密码用 Hasher 生成，实际生产建议用数据库）
credentials = {
    "usernames": {
        "test": {
            "name": "测试用户",
            "password": Hasher(["test123"]).generate()[0]  # 密码：test123
        },
        "yulin": {
            "name": "煜麟",
            "password": Hasher(["你的密码123"]).generate()[0]  # 改成你想用的密码
        }
    }
}

# cookie 配置（保持登录状态）
cookie_name = "jiuling_cookie"
cookie_key = "random_key_change_me_2026"  # 改成随机字符串（防篡改）
cookie_expiry_days = 30

authenticator = stauth.Authenticate(
    credentials,
    cookie_name,
    cookie_key,
    cookie_expiry_days
)

# 登录界面（必须放最前面）
name, authentication_status, username = authenticator.login("登录九灵AI 智能办公平台", "main")

if authentication_status:
    st.sidebar.success(f"欢迎回来，{name}！")
    
    # 登出按钮放 sidebar
    authenticator.logout("登出", "sidebar")
    
    # 登录成功后，才显示下面全部内容
    # -----------------------------------------------------------------------------
    # 1. 页面配置与 CSS 美化
    # -----------------------------------------------------------------------------
    st.set_page_config(
        page_title="九灵AI 智能办公平台",
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 自定义 CSS 样式（原样保留）
    st.markdown("""
    <style>
        /* 你的原 CSS 全部保留 */
        .stApp { background-color: #f8f9fa; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
        h1, h2, h3 { color: #0d47a1; font-weight: 700; }
        .block-container { padding-top: 2rem; padding-bottom: 2rem; }
        .stButton > button { background-color: #1e88e5; color: white; border-radius: 8px; border: none; padding: 0.5rem 1rem; font-weight: 600; transition: all 0.3s ease; }
        .stButton > button:hover { background-color: #1565c0; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .report-card { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); border: 1px solid #e0e0e0; margin-bottom: 20px; }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e0e0e0; }
        .custom-footer { text-align: center; color: #757575; padding: 20px; font-size: 0.85rem; border-top: 1px solid #e0e0e0; margin-top: 40px; }
    </style>
    """, unsafe_allow_html=True)

    # -----------------------------------------------------------------------------
    # 2. 核心逻辑函数（原样保留）
    # -----------------------------------------------------------------------------
    def process_multiple_files(files):
        """处理多个上传的文件，提取内容并汇总"""
        # ... 你的原函数代码保持不变 ...
        pass  # 这里粘贴你原来的 process_multiple_files 函数完整代码

    def call_llm(prompt, provider, key, base_url=None):
        """调用大模型函数（原样保留）"""
        # ... 你的原 call_llm 函数完整代码 ...
        pass

    def get_analysis_report(combined_content, query, provider, key, base_url=None):
        """生成报告函数（原样保留）"""
        # ... 你的原 get_analysis_report 函数完整代码 ...
        pass

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
        
        # 插件生态仪表盘（原样保留）
        with st.expander("📊 插件生态概览（模拟）", expanded=False):
            # ... 你的原模拟数据代码保持不变 ...
            pass

        st.markdown("---")
        st.info("基于 Kimi K2.5 模型，提供全流程文档与数据智能分析")

    # --- 主区域 ---
    col_header_1, col_header_2 = st.columns([1, 6])
    with col_header_1:
        st.markdown("# 🚀")
    with col_header_2:
        st.title("九灵AI 智能办公平台 - 数据与文档分析助手")
        st.markdown("##### Professional AI Office Demo | Data & Document Intelligence")
    st.markdown("---")

    # 主内容分栏布局（原样保留）
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
            当前时间：{current_time} | Version 1.4.0 (生态演示版)
        </div>
        """,
        unsafe_allow_html=True
    )

elif authentication_status is False:
    st.error("用户名或密码错误")
elif authentication_status is None:
    st.warning("请登录使用")
    st.stop()  # 登录前不显示任何内容