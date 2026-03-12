import os
import sys
import pandas as pd
from pypdf import PdfReader
from docx import Document
import re

def get_pdf_page_count(file_path):
    """获取 PDF 文件页数"""
    try:
        reader = PdfReader(file_path)
        return len(reader.pages)
    except Exception as e:
        print(f"读取 PDF {file_path} 出错: {e}")
        return 0

import zipfile
import xml.etree.ElementTree as ET

def get_docx_page_count(file_path):
    """强力获取 Word 文件页数 (支持 python-docx 和 XML 直接解析)"""
    try:
        # 第一步：尝试直接用 python-docx 的 core_properties (注意：不同版本属性名不同)
        doc = Document(file_path)
        props = doc.core_properties
        # 尝试常见的页码属性名
        pages = getattr(props, 'pages', 0)
        
        if pages and pages > 0:
            return pages
            
        # 第二步：如果上面失败了或为0，直接解压读取 docProps/app.xml (这是 OpenXML 的标准)
        with zipfile.ZipFile(file_path) as z:
            # 获取所有文件名，确认 app.xml 是否存在
            if 'docProps/app.xml' in z.namelist():
                app_xml = z.read('docProps/app.xml')
                root = ET.fromstring(app_xml)
                # 定义命名空间
                ns = {'ns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
                pages_node = root.find('.//ns:Pages', ns)
                if pages_node is not None:
                    return int(pages_node.text)
        return 0
    except Exception as e:
        print(f"读取 Word {file_path} 出错: {e}")
        return 0


def get_doc_page_count(file_path):
    """处理旧版 .doc 文件的页数 (较难直接获取，尝试读取元数据)"""
    # 注意：旧版 .doc 格式与 docx 完全不同，python-docx 无法读取。
    # 这里我们返回一个提示，或者尝试简单的字节流搜索（不推荐用于生产）
    return "格式不支持统计"

def process_matching(excel_path, doc_dir, output_path):
    # 1. 读取 Excel
    print(f"正在读取 Excel: {excel_path}...")
    df = pd.read_excel(excel_path)
    
    target_col = '标题'
    if target_col not in df.columns:
        print(f"错误：Excel 中未找到 '{target_col}' 列。存在的列有: {list(df.columns)}")
        return

    # 2. 扫描 doc 文件夹 (包含子文件夹)
    print(f"正在扫描目录: {doc_dir}...")
    doc_files = []
    for root, dirs, files in os.walk(doc_dir):
        for f in files:
            # 增加对 .doc 的支持
            if f.lower().endswith(('.pdf', '.docx', '.doc')):
                doc_files.append({
                    'name': f,
                    'path': os.path.join(root, f)
                })

    # 3. 匹配与合并结果
    results_match = []
    results_file = []
    results_pages = []

    for index, row in df.iterrows():
        title = str(row[target_col]).strip()
        
        # 寻找该标题的所有匹配项
        matched_infos = []
        for doc_file in doc_files:
            file_name_without_ext = os.path.splitext(doc_file['name'])[0]
            if title == file_name_without_ext:
                f_name = doc_file['name']
                # 提取页码
                pages = 0
                if f_name.lower().endswith('.pdf'):
                    pages = get_pdf_page_count(doc_file['path'])
                elif f_name.lower().endswith('.docx'):
                    pages = get_docx_page_count(doc_file['path'])
                elif f_name.lower().endswith('.doc'):
                    pages = get_doc_page_count(doc_file['path'])
                matched_infos.append((f_name, pages))

        if not matched_infos:
            results_match.append('否')
            results_file.append('')
            results_pages.append('')
        else:
            results_match.append('是')
            # 使用 " - " 作为切分符号，增强可视化效果
            results_file.append(" | ".join([info[0] for info in matched_infos]))
            results_pages.append(" | ".join([str(info[1]) for info in matched_infos]))

    # 4. 更新 DataFrame 并保存
    df['匹配状态'] = results_match
    df['匹配文件名'] = results_file
    df['文件实际页码'] = results_pages

    print(f"正在保存结果到: {output_path}...")
    df.to_excel(output_path, index=False)
    print("处理完成！")

if __name__ == "__main__":
    # 配置路径：需要适配 PyInstaller 打包后的路径获取
    if hasattr(sys, '_MEIPASS'):
        # 如果是打包后的环境，sys.executable 是 exe 的路径
        BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
    else:
        # 如果是源代码运行环境
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    EXCEL_FILE = os.path.join(BASE_DIR, "样本.xlsx")
    DOC_FOLDER = os.path.join(BASE_DIR, "doc")
    OUTPUT_FILE = os.path.join(BASE_DIR, "匹配结果_汇总.xlsx")

    process_matching(EXCEL_FILE, DOC_FOLDER, OUTPUT_FILE)
