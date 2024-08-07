import os
import shutil
import streamlit as st
import pdfplumber
import zipfile
import csv
from docx import Document
import patoolib  # 添加patoolib库以处理多种压缩格式


# 定义函数，读取并统计PDF文件中指定词语的个数
def count_word_in_pdf(file_path, word):
    count = 0
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    count += text.count(word)
    except Exception as e:
        st.error(f"Error reading PDF file {file_path}: {e}")
    return count


# 定义函数，读取并统计TXT文件中指定词语的个数
def count_word_in_txt(file_path, word):
    count = 0
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
            count = text.count(word)
    except Exception as e:
        st.error(f"Error reading TXT file {file_path}: {e}")
    return count


# 定义函数，读取并统计CSV文件中指定词语的个数
def count_word_in_csv(file_path, word):
    count = 0
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row in reader:
                row_text = ','.join(row)
                count += row_text.count(word)
    except Exception as e:
        st.error(f"Error reading CSV file {file_path}: {e}")
    return count


# 定义函数，读取并统计Word文件中指定词语的个数
def count_word_in_docx(file_path, word):
    count = 0
    try:
        doc = Document(file_path)
        for paragraph in doc.paragraphs:
            count += paragraph.text.count(word)
    except Exception as e:
        st.error(f"Error reading DOCX file {file_path}: {e}")
    return count


# 定义函数，遍历文件夹中的所有文件，并统计指定词语的总数
def count_word_in_folder(folder_path, word):
    word_count = 0
    results = []
    seen_files = set()
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if filename in seen_files:
                continue
            seen_files.add(filename)
            file_path = os.path.join(root, filename)
            if filename.endswith('.pdf'):
                count = count_word_in_pdf(file_path, word)
            elif filename.endswith('.txt'):
                count = count_word_in_txt(file_path, word)
            elif filename.endswith('.csv'):
                count = count_word_in_csv(file_path, word)
            elif filename.endswith('.docx'):
                count = count_word_in_docx(file_path, word)
            else:
                continue
            word_count += count
            results.append((filename, count))
    return word_count, results


# 解压上传的文件夹，并处理乱码问题
def unzip_and_filter_folders(zip_file, extract_to, file_type):
    if file_type == "zip":
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
    elif file_type == "rar":
        patoolib.extract_archive(zip_file, outdir=extract_to)

    extracted_files = os.listdir(extract_to)

    normal_folders = []
    for folder in extracted_files:
        folder_path = os.path.join(extract_to, folder)
        if os.path.isdir(folder_path):
            normal_folders.append(folder_path)

    if len(normal_folders) == 0:
        st.error("未找到正常显示的文件夹")
        return None

    return normal_folders[0]  # 返回第一个正常文件夹的路径


# 尝试多种编码方案来解码文件名
def try_decode(filename):
    encodings = ['utf-8', 'gbk', 'latin1', 'cp437']
    for enc in encodings:
        try:
            return filename.encode('cp437', errors='ignore').decode(enc)
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return filename  # 如果解码失败，返回原始文件名


# Streamlit 应用程序
def main():
    st.markdown("<h3 style='text-align: center; color: black;'>📚在多个文件里检索某个词语的个数</h3>",
                unsafe_allow_html=True)
    st.write('\n')
    st.write('\n')
    st.write('\n')
    st.write('\n')
    uploaded_file = st.file_uploader("请选择一个包含多个文件的压缩文件【zip或者rar】", type=["zip", "rar"])
    word = st.text_input("请输入要检索的词语:")
    search_button = st.button("检索")

    if search_button and uploaded_file is not None:
        with st.spinner("正在处理..."):
            folder_path = "extracted_files"
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            # 检测文件类型
            file_type = uploaded_file.name.split(".")[-1]

            # 解压并过滤ZIP/RAR文件
            normal_folder_path = unzip_and_filter_folders(uploaded_file, folder_path, file_type)

            if normal_folder_path:
                total_count, results = count_word_in_folder(normal_folder_path, word)
                st.write(f"总共找到 '{word}' {total_count} 次")
                for filename, count in results:
                    decoded_filename = try_decode(filename)
                    st.write(f"文件: {decoded_filename}, 包含 '{word}' {count} 次")

    # 添加分割线和版权信息，固定在页面底部
    st.markdown("""
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: white;
        text-align: center;
        padding: 10px 0;
        border-top: 1px solid #ddd;
        font-family: KaiTi;
    }
    .email {
        font-family: KaiTi,Times New Roman;
    }
    </style>
    <div class="footer">
        Copyright © 2024-长期 版权所有：华师沈威，在使用中如果有任何问题可以发邮件至：sw@ccnu.edu.cn
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
