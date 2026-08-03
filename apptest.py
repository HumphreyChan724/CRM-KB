import os
os.environ["PYTHONUTF8"] = "1"  # 强制使用 UTF-8，防止 Windows 乱码/解码报错
import time
import zipfile
import requests
import shutil
import re
import jieba  # 💡 引入中文分词库
import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from langchain_community.vectorstores import FAISS
from langchain_community.embeddings import FastEmbedEmbeddings
from langchain_community.retrievers import BM25Retriever  # 💡 引入 BM25 精准字面检索器

# ==================== 🛠️ 自定义配置区 ====================
BASE_ZIP_URL = "https://raw.githubusercontent.com/HumphreyChan724/CRM-KB/main/storage.zip"
TIMEOUT_LIMIT = 1800  

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORAGE_DIR = os.path.join(BASE_DIR, "storage")

DEFAULT_QUICK_QUERIES = [
    "How to reset CRM login password?",
    "Steps for processing customer refund operations",
    "How to check live warehouse inventory status?"
]
# ========================================================

# --- 🧠 全域拦截 JS (精准放行 Ctrl+C / Cmd+C 复制，封杀单按 C/R 弹窗) ---
components.html("""
    <script>
    const parentWin = window.parent;
    const parentDoc = parentWin.document;

    function blockStreamlitShortcuts(e) {
        const target = e.target;
        if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable) {
            return;
        }
        
        if (e.ctrlKey || e.metaKey) {
            if (e.key.toLowerCase() === 'c' || e.key.toLowerCase() === 'r') {
                e.stopImmediatePropagation();
            }
            return;
        }

        if (e.key.toLowerCase() === 'c' || e.key.toLowerCase() === 'r') {
            e.stopImmediatePropagation();
            e.preventDefault();
        }
    }

    parentDoc.addEventListener('keydown', blockStreamlitShortcuts, true);
    parentDoc.addEventListener('keyup', blockStreamlitShortcuts, true);
    parentWin.addEventListener('keydown', blockStreamlitShortcuts, true);
    parentWin.addEventListener('keyup', blockStreamlitShortcuts, true);
    </script>
""", height=0)

# --- 🎯 智能解压与全自动同步 (支持多分卷流式下载并自动拼接) ---
def download_and_update_knowledge():
    try:
        temp_combined_zip = os.path.join(BASE_DIR, "temp_combined_storage.zip")
        if os.path.exists(temp_combined_zip):
            os.remove(temp_combined_zip)

        part_num = 1
        downloaded_any = False

        # 💡 顺序下载所有分卷 (storage.zip.001, storage.zip.002...) 并流式拼接到本地
        with open(temp_combined_zip, "wb") as merged_file:
            while True:
                part_url = f"{BASE_ZIP_URL}.{part_num:03d}?t={int(time.time())}"
                response = requests.get(part_url, stream=True, timeout=60)
                
                # 如果遇到 404，说明所有分卷已下载完毕
                if response.status_code != 200:
                    break
                
                for chunk in response.iter_content(chunk_size=1024 * 1024):  # 每次写入 1MB
                    if chunk:
                        merged_file.write(chunk)
                
                downloaded_any = True
                part_num += 1

        if downloaded_any:
            temp_extract_dir = os.path.join(BASE_DIR, "temp_extract")
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir)
            os.makedirs(temp_extract_dir, exist_ok=True)

            # 💡 将拼合后的完整 Zip 解压
            with zipfile.ZipFile(temp_combined_zip, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_dir)

            real_data_dir = None
            for root, dirs, files in os.walk(temp_extract_dir):
                if "index.faiss" in files:
                    real_data_dir = root
                    break

            if real_data_dir:
                if os.path.exists(STORAGE_DIR):
                    shutil.rmtree(STORAGE_DIR)
                shutil.copytree(real_data_dir, STORAGE_DIR)

            # 💡 清理临时文件
            shutil.rmtree(temp_extract_dir)
            if os.path.exists(temp_combined_zip):
                os.remove(temp_combined_zip)

            st.cache_resource.clear()
            st.cache_data.clear()
            return True
    except Exception as e:
        print(f"Sync failed: {e}")
    return False

st.set_page_config(page_title="CRM AI Assistant", page_icon="🤖", layout="centered")

# --- 🎨 布局与 UI 美化 CSS ---
hide_streamlit_style = """
            <style>
            [data-testid="stHeader"], [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], footer { 
                display: none !important; 
                visibility: hidden !important; 
            }
            .block-container {
                max-width: 1000px !important;
                padding-left: 2rem !important;
                padding-right: 2rem !important;
                padding-top: 3rem !important;
            }
            [data-testid="stBottom"] {
                max-width: 1000px !important;
                margin: 0 auto !important;
                left: 0 !important;
                right: 0 !important;
            }
            [data-testid="stBottomBlockContainer"] {
                max-width: 1000px !important;
                margin: 0 auto !important;
                padding-left: 2rem !important;
                padding-right: 2rem !important;
            }
            [data-testid="stChatInput"] textarea {
                padding-left: 12px !important;
            }

            [data-testid="stChatMessage"] h1 {
                font-size: 1.25rem !important;
                font-weight: 700 !important;
                margin-top: 0.8rem !important;
                margin-bottom: 0.4rem !important;
                line-height: 1.3 !important;
            }
            [data-testid="stChatMessage"] h2 {
                font-size: 1.1rem !important;
                font-weight: 600 !important;
                margin-top: 0.6rem !important;
                margin-bottom: 0.3rem !important;
                line-height: 1.3 !important;
            }
            [data-testid="stChatMessage"] h3 {
                font-size: 1.0rem !important;
                font-weight: 600 !important;
                margin-top: 0.5rem !important;
                margin-bottom: 0.2rem !important;
            }
            [data-testid="stChatMessage"] p, [data-testid="stChatMessage"] li {
                font-size: 0.95rem !important;
                line-height: 1.6 !important;
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

if "last_activity_time" not in st.session_state:
    st.session_state.last_activity_time = time.time()

if time.time() - st.session_state.last_activity_time > TIMEOUT_LIMIT:
    st.error("⏳ Session timeout due to inactivity. Please refresh the page to reconnect.")
    st.stop()

st.session_state.last_activity_time = time.time()

if "has_updated_kb" not in st.session_state:
    with st.spinner("Synchronizing CRM core systems... Please wait..."):
        download_and_update_knowledge()
    st.session_state.has_updated_kb = True

title_col, btn_col = st.columns([7.5, 2.5])
with title_col:
    st.title("I am your CRM AI Assistant 🤖")
with btn_col:
    st.write("")  
    st.write("") 
    if st.button("➕ New Chat", use_container_width=True, type="primary"):
        st.session_state.chat_history = []
        st.rerun()

st.caption("**Service Scope:** How to use CRM | CRM Operation Issue")
st.caption("**Service Targets:** Receptionist | Technician | Shop Manager | Warehouse Keeper | Regional Manager | Country Manager")
st.write("---")

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# --- ⚙️ 双引擎加载区域 (FAISS 向量语义 + BM25 精准字面) ---
@st.cache_resource
def load_knowledge_base():
    if not os.path.exists(STORAGE_DIR): return None, None
    try:
        embeddings = FastEmbedEmbeddings(model_name="BAAI/bge-small-zh-v1.5")
        vector_store = FAISS.load_local(STORAGE_DIR, embeddings, allow_dangerous_deserialization=True)
        
        all_docs = list(vector_store.docstore._dict.values())
        
        def chinese_tokenizer(text):
            return list(jieba.cut(text))
            
        bm25_retriever = BM25Retriever.from_documents(
            all_docs,
            preprocess_func=chinese_tokenizer
        )
        bm25_retriever.k = 8
        
        return vector_store, bm25_retriever
    except Exception:
        return None, None

vector_store, bm25_retriever = load_knowledge_base()

# 💡 渲染历史聊天记录（历史记录纯粹干净，不渲染历史调试框）
chat_container = st.container()

with chat_container:
    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

@st.cache_resource
def init_agent():
    return OpenAI(
        base_url="https://tokenhub.tencentmaas.com/v1",
        api_key="sk-9hFfbE2MHPClZC31RBJwiFx7Hx8L0cFNZWhfBMneGU2eDPsU"
    )

client = init_agent()

# --- 🧠 动态快捷词 ---
@st.cache_data(show_spinner=False)
def generate_dynamic_queries(_vs, _client):
    if not _vs: return DEFAULT_QUICK_QUERIES
    try:
        docs = list(_vs.docstore._dict.values())[:6]
        combined_text = "\n".join([d.page_content for d in docs])[:2000]
        prompt = """You are a CRM system business analyst. Extract exactly 3 concise common operational questions from the text. One per line, no bullets."""
        response = _client.chat.completions.create(
            model="deepseek-v4-flash-202605",
            messages=[{"role": "user", "content": prompt + "\n" + combined_text}],
            temperature=0.4
        )
        lines = [line.strip().lstrip("1234567890. -*•") for line in response.choices[0].message.content.strip().split("\n") if line.strip()]
        while len(lines) < 3: lines.append(DEFAULT_QUICK_QUERIES[len(lines)])
        return lines[:3]
    except Exception: return DEFAULT_QUICK_QUERIES

suggestions = generate_dynamic_queries(vector_store, client)

if not st.session_state.chat_history:
    st.write("💡 **Quick Queries:**")
    cols = st.columns(3)
    for i, sug in enumerate(suggestions):
        if cols[i].button(sug, key=f"sug_btn_{i}", use_container_width=True):
            st.session_state.chat_history.append({"role": "user", "content": sug})
            st.session_state.last_activity_time = time.time()
            st.rerun()

user_input = st.chat_input("Enter your questions...")
if user_input:
    st.session_state.chat_history.append({"role": "user", "content": user_input})
    st.session_state.last_activity_time = time.time()
    st.rerun()

# --- 🚀 核心混合检索（Hybrid Search）与响应流程 ---
if st.session_state.chat_history and st.session_state.chat_history[-1]["role"] == "user":
    latest_query = st.session_state.chat_history[-1]["content"]
    
    with chat_container:
        with st.chat_message("assistant"):
            placeholder = st.empty()
            placeholder.markdown("AI Assistant is organizing thoughts...")
            
            context_str = ""
            retrieved_docs_preview = []
            
            if vector_store is not None:
                # 💡 1. 语义搜索 (FAISS 召回 Top 8)
                faiss_docs = vector_store.similarity_search(latest_query, k=8)
                
                # 💡 2. 字面精准搜索 (BM25 召回 Top 8)
                bm25_docs = []
                if bm25_retriever is not None:
                    try:
                        bm25_docs = bm25_retriever.invoke(latest_query)
                    except Exception:
                        bm25_docs = []
                
                # 💡 3. 双路合并 & 交叉去重
                combined_docs = []
                seen_hashes = set()
                
                for doc in bm25_docs + faiss_docs:
                    raw_text = doc.page_content
                    clean_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ufffd]', '', raw_text)
                    clean_text = re.sub(r'[\u200b-\u200d\ufeff]', '', clean_text)
                    clean_text = re.sub(r'\s+', ' ', clean_text).strip()
                    
                    dedup_key = clean_text[:100]
                    if dedup_key not in seen_hashes:
                        seen_hashes.add(dedup_key)
                        doc.page_content = clean_text
                        combined_docs.append(doc)
                
                final_docs = combined_docs[:12]
                
                for idx, doc in enumerate(final_docs):
                    context_str += f"{doc.page_content}\n\n"
                    retrieved_docs_preview.append(f"**[片段 {idx+1}]**:\n{doc.page_content}") 
            
            system_prompt = f"""# Role

You are an expert CRM Frontline Operations Coach with 10 years of experience. Your target audience consists of frontline staff (Receptionists, Technicians, Warehouse Keepers, Shop Managers).

# CRITICAL LANGUAGE DIRECTIVE (最高优先级语种指令)
- **Primary Rule**: Detect the language of the user's query: "{latest_query}".
- **Language Matching**: You MUST respond in the EXACT SAME LANGUAGE as the user's prompt!
  - User query is in English -> Reply COMPLETELY in English.
  - User query is in Spanish -> Reply COMPLETELY in Spanish.
  - User query is in Portuguese -> Reply COMPLETELY in Portuguese.
  - User query is in French -> Reply COMPLETELY in French.
  - User query is in Russian -> Reply COMPLETELY in Russian.
  - User query is in Chinese -> Reply COMPLETELY in Chinese.
- ⚠️ **IMPORTANT**: The Reference Content below is written in Chinese. You must read it for facts, but **YOU MUST TRANSLATE AND PRESENT YOUR ENTIRE RESPONSE IN THE USER'S QUERY LANGUAGE**.

# 硬性约束 (Highest Priority)
1. 绝对零成本原则：不得推荐任何付费 API、付费服务。
2. 保持精度：必须基于 CHUNK_SIZE=500 的检索分块进行回答。

# Style & Forbidden Terms
- **NO AI Boilerplate**: NEVER start your answer with phrases like "According to the knowledge base", "Based on SOP", "Provided documents state".
- **NO Backend IT Codes**: STRICTLY FORBIDDEN to display technical module IDs (e.g. M03_BU03_BF08).
- **Frontline UI Naming ONLY**: Translate technical module names into clear, human-readable CRM UI menu names.

# Reference Content (Knowledge Base in Chinese):
{context_str}

# Workflow & RAG Rules
1. Extract answers strictly from the Reference Content above.
2. Translate all steps and UI menu names into the user's language smoothly and naturally.
3. Standard Fallback Rule (When no operational steps exist in reference):
   If reference data is completely irrelevant or missing operational guidance, output ONLY the standard fallback in the USER'S LANGUAGE:
   - 🇬🇧 **English**: 💡 Sorry, the current official knowledge base has not recorded the detailed steps for this specific operation. To avoid misoperation, it is recommended to contact the system administrator or check the latest standard operating manual.
   - 🇨🇳 **Chinese**: 💡 抱歉，当前官方知识库中暂未收录该特定操作的详细步骤。为了避免误操作，建议您联系系统管理员或查阅最新的标准操作手册。

# Output Format & Visual Constraints
- Use clean Markdown with ordered numbered steps.
- **DO NOT use large headers like `#` (H1) or `##` (H2)**. Use `###` (H3) or **bold text**.
- Use **bold text** for UI elements, menu names, and buttons.

# FINAL CHECK BEFORE OUTPUT:
Target Language: Match the language of "{latest_query}"."""
            
            try:
                stream = client.chat.completions.create(
                    model="deepseek-v4-flash-202605",
                    messages=[{"role": "system", "content": system_prompt}] + [{"role": m["role"], "content": m["content"]} for m in st.session_state.chat_history],
                    stream=True
                )
                
                full_response = ""
                for chunk in stream:
                    if chunk.choices[0].delta.content is not None:
                        full_response += chunk.choices[0].delta.content
                        placeholder.markdown(full_response + "▌")
                
                placeholder.markdown(full_response)
                st.session_state.chat_history.append({"role": "assistant", "content": full_response})

                # 💡 只在【最近一次回答】下方渲染调试框，发起下一次提问后自动隐去
                if retrieved_docs_preview:
                    with st.expander(f"🔍 调试专用：点击查看向量库本次抽取的 {len(retrieved_docs_preview)} 段原文）："):
                        for doc_text in retrieved_docs_preview:
                            st.text(doc_text)
                            st.divider()
                
            except Exception as e:
                placeholder.markdown("❌ Connection error, please check your network.")
                st.error(f"Error: {str(e)}")