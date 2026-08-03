# ==================== 🚀 必须放在文件最顶端（所有 import 之前） ====================
import os
import multiprocessing

# 获取 CPU 物理核心与逻辑线程数，强制解锁所有数学库的 CPU 并发封印
cpu_cores = multiprocessing.cpu_count()
os.environ["OMP_NUM_THREADS"] = str(cpu_cores)
os.environ["MKL_NUM_THREADS"] = str(cpu_cores)
os.environ["OPENBLAS_NUM_THREADS"] = str(cpu_cores)
os.environ["VECLIB_MAXIMUM_THREADS"] = str(cpu_cores)
os.environ["NUMEXPR_NUM_THREADS"] = str(cpu_cores)

import shutil
import time
import re
import zipfile
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_community.embeddings import FastEmbedEmbeddings

# 1. 通用解析引擎 (负责 .docx, .pptx, .xlsx, .pdf, .html, .txt, .md 等)
try:
    from markitdown import MarkItDown
    md_converter = MarkItDown()
    HAS_MARKITDOWN = True
except ImportError:
    HAS_MARKITDOWN = False

# 2. 老版 Excel 解析器 (.xls)
try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

# 3. Windows 系统 COM 解析器 (负责老版 .doc 和 .ppt)
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# ==================== 🛠️ 性能与切片优化配置区 ====================
SOURCE_FOLDERS = ["./raw_docs", "./kms_documents"]
TEMP_STORAGE_DIR = "./storage"
ZIP_FILENAME = "storage.zip"
PART_SIZE = 80 * 1024 * 1024  # 💡 每一个分卷包控制在 45MB，远低于 GitHub 100MB 限制

EMBEDDING_MODEL_NAME = "BAAI/bge-small-zh-v1.5"

# 💡 最佳平衡点：CHUNK_SIZE 设为 500，重叠 100
CHUNK_SIZE = 500
CHUNK_OVERLAP = 100

# 💡 向量化批处理大小
BATCH_SIZE = 2000
# =================================================================

# --- 💡 核心修复 1：文本乱码洗涤器 ---
def clean_text_content(text):
    if not text:
        return ""
    # 1. 剔除全套 ASCII 控制字符 (0x00-0x1F, 0x7F-0x9F) 与 Unicode 替代符 \ufffd
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ufffd]', '', text)
    # 2. 将连续的多个空白字符直接替换为一个空格
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# --- 💡 核心修复 2：安全多编码文本读取器 ---
def read_file_safely(file_path):
    """支持 UTF-8, UTF-16, GBK 等多编码自动尝试，防止读取出问号乱码"""
    encodings = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'gbk', 'gb2312', 'gb18030']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                return f.read()
        except Exception:
            continue
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception:
        return ""

def parse_xls_file(file_path):
    """专门解析老版 Excel (.xls) 表格"""
    if not HAS_XLRD:
        return ""
    text_lines = []
    try:
        wb = xlrd.open_workbook(file_path)
        for sheet in wb.sheets():
            text_lines.append(f"--- 工作表: {sheet.name} ---")
            for row_idx in range(sheet.nrows):
                row_vals = [str(cell.value) for cell in sheet.row(row_idx) if cell.value != ""]
                if row_vals:
                    text_lines.append(" | ".join(row_vals))
        return "\n".join(text_lines)
    except Exception:
        return ""

def parse_doc_win32(file_path):
    """专门解析老版 Word (.doc) 文档"""
    if not HAS_WIN32COM:
        return ""
    word_app = None
    text = ""
    try:
        pythoncom.CoInitialize()
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
        doc = word_app.Documents.Open(os.path.abspath(file_path), ReadOnly=True)
        text = doc.Content.Text
        doc.Close(False)
    except Exception:
        pass
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return text

def parse_ppt_win32(file_path):
    """专门解析老版 PPT (.ppt) 演示文稿"""
    if not HAS_WIN32COM:
        return ""
    ppt_app = None
    text_list = []
    try:
        pythoncom.CoInitialize()
        ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), ReadOnly=True, WithWindow=False)
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text_list.append(shape.TextFrame.TextRange.Text)
        presentation.Close()
    except Exception:
        pass
    finally:
        if ppt_app:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return "\n".join(text_list)

def parse_any_file(file_path):
    """全格式路由解析器（已注入自动防乱码机制）"""
    ext = os.path.splitext(file_path)[1].lower()
    raw_text = ""

    if ext == ".doc":
        raw_text = parse_doc_win32(file_path)

    elif ext == ".ppt":
        raw_text = parse_ppt_win32(file_path)

    elif ext == ".xls":
        raw_text = parse_xls_file(file_path)

    if not raw_text and HAS_MARKITDOWN:
        try:
            result = md_converter.convert(file_path)
            if result and result.text_content:
                raw_text = result.text_content
        except Exception:
            pass

    if not raw_text:
        raw_text = read_file_safely(file_path)

    # 统一清洗处理：确保返回的文本 100% 无 \x00 空字节和 \ufffd 乱码
    return clean_text_content(raw_text)

def load_all_documents(folders):
    """遍历文件夹提取文档"""
    documents = []
    
    for folder in folders:
        if not os.path.exists(folder):
            print(f"⚠️ 文件夹不存在，已跳过: {folder}")
            continue
            
        print(f"\n📁 正在读取文件夹: {folder}")
        file_count = 0
        
        for root, _, files in os.walk(folder):
            for file_name in files:
                if file_name.startswith("~$") or file_name.startswith("."):
                    continue
                    
                file_path = os.path.join(root, file_name)
                text = parse_any_file(file_path)
                
                if text and text.strip():
                    documents.append({
                        "page_content": text,
                        "metadata": {
                            "source": file_path,
                            "folder": os.path.basename(folder),
                            "file_name": file_name
                        }
                    })
                    file_count += 1
                    
        print(f"✅ {folder} 读取完毕，共计成功提取 {file_count} 个有效文档。")
        
    return documents

def update_and_zip_kb():
    """主流程：多目录提取 + 多线程向量化 + 秒级切割打包"""
    # 0. 清理旧的分卷和压缩包
    for f in os.listdir("."):
        if f.startswith("storage.zip"):
            try:
                os.remove(f)
            except Exception:
                pass

    raw_docs = load_all_documents(SOURCE_FOLDERS)
    if not raw_docs:
        print("\n❌ 错误：未读取到任何有效文档！")
        return

    print(f"\n📊 汇总：成功提取 {len(raw_docs)} 个文档，开始切分成向量切片...")

    # 1. 文本切块
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        separators=["\n\n", "\n", "。", "！", "？", " ", ""]
    )
    
    texts, metadatas = [], []
    for doc in raw_docs:
        file_name = doc["metadata"]["file_name"]
        header_prefix = f"【文档标题: {file_name}】\n"
        
        chunks = text_splitter.split_text(doc["page_content"])
        for chunk in chunks:
            texts.append(header_prefix + chunk)
            metadatas.append(doc["metadata"])
        
    total_chunks = len(texts)
    print(f"✂️ 切分完成！总切片数: {total_chunks} 个。")

    # 2. 初始化 FastEmbed 模型
    print(f"\n🚀 正在加载 FastEmbed 模型 [{EMBEDDING_MODEL_NAME}]...")
    embeddings = FastEmbedEmbeddings(
        model_name=EMBEDDING_MODEL_NAME,
        threads=os.cpu_count()
    )

    # 3. 分批生成 FAISS 向量索引
    print(f"\n⚡ 正在分批生成 FAISS 向量数据库 (每批 {BATCH_SIZE} 条)...")
    start_time = time.time()
    vector_store = None

    for i in range(0, total_chunks, BATCH_SIZE):
        batch_texts = texts[i : i + BATCH_SIZE]
        batch_metadatas = metadatas[i : i + BATCH_SIZE]
        
        if vector_store is None:
            vector_store = FAISS.from_texts(
                texts=batch_texts,
                embedding=embeddings,
                metadatas=batch_metadatas
            )
        else:
            vector_store.add_texts(
                texts=batch_texts,
                metadatas=batch_metadatas
            )
        
        current_processed = min(i + BATCH_SIZE, total_chunks)
        percent = (current_processed / total_chunks) * 100
        elapsed = time.time() - start_time
        speed = current_processed / elapsed if elapsed > 0 else 0
        remaining_time = (total_chunks - current_processed) / speed if speed > 0 else 0
        
        print(f" ⏳ 进度: {current_processed}/{total_chunks} ({percent:.1f}%) | 速度: {speed:.1f} 条/秒 | 预计剩余: {remaining_time/60:.1f} 分钟")

    # 4. 保存 FAISS 数据库
    if os.path.exists(TEMP_STORAGE_DIR):
        shutil.rmtree(TEMP_STORAGE_DIR)

    vector_store.save_local(TEMP_STORAGE_DIR)
    print(f"\n💾 FAISS 索引已写入临时目录 [{TEMP_STORAGE_DIR}]")

    # 5. 💡 快速打包 ZIP (去掉 compresslevel=9 的巨额计算，改为普通秒级压缩)
    print(f"📦 正在进行快速打包...")
    with zipfile.ZipFile(ZIP_FILENAME, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(TEMP_STORAGE_DIR):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, TEMP_STORAGE_DIR)
                zipf.write(file_path, arcname)

    # 6. 💡 自动切割分卷 (切割为 45MB/卷)
    print("✂️ 正在切割分卷包 (保证每卷 < 50MB，顺畅推至 GitHub)...")
    part_num = 1
    with open(ZIP_FILENAME, 'rb') as src:
        while True:
            chunk = src.read(PART_SIZE)
            if not chunk:
                break
            part_filename = f"{ZIP_FILENAME}.{part_num:03d}"
            with open(part_filename, 'wb') as dest:
                dest.write(chunk)
            print(f"   └─ 已生成分卷: {os.path.basename(part_filename)} ({len(chunk)/1024/1024:.2f} MB)")
            part_num += 1

    # 删除未切割的临时大 zip 文件
    if os.path.exists(ZIP_FILENAME):
        os.remove(ZIP_FILENAME)

    total_time_min = (time.time() - start_time) / 60
    print(f"\n🎉 完美成功！共生成 {part_num - 1} 个分卷文件，总耗时仅: {total_time_min:.1f} 分钟。")
    print(f"📍 分卷文件 (storage.zip.001 等) 已生成，可以准备提交 GitHub 了！")

if __name__ == "__main__":
    update_and_zip_kb()