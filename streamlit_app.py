import os
import tempfile
import streamlit as st
import gdown
import pandas as pd

from pypdf import PdfReader
from docx import Document
from google import genai


# =========================
# CẤU HÌNH
# =========================
st.set_page_config(
    page_title="CHATBOT TRA CỨU TÀI LIỆU IC PXVH1",
    page_icon="🤖",
    layout="wide"
)

DRIVE_FOLDER_ID = "1P44hHly9bSdVZps4oqIgeReclxQWCzIm"

client = genai.Client(
    api_key=st.secrets["GEMINI_API_KEY"]
)


# =========================
# ĐỌC TÀI LIỆU
# =========================
def read_pdf(path):
    text = ""
    reader = PdfReader(path)

    for i, page in enumerate(reader.pages):
        try:
            page_text = page.extract_text() or ""
            text += f"\n\n--- Trang {i+1} ---\n{page_text}"
        except:
            pass

    return text


def read_docx(path):
    doc = Document(path)

    paragraphs = []

    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append(p.text)

    return "\n".join(paragraphs)


def read_excel(path):
    text = ""

    try:
        excel = pd.ExcelFile(path)

        for sheet in excel.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet)

            text += f"\n\n--- Sheet: {sheet} ---\n"
            text += df.fillna("").astype(str).to_csv(index=False)

    except Exception as e:
        text += f"\nLỗi đọc Excel: {e}"

    return text


def read_file(path):
    name = os.path.basename(path)
    ext = os.path.splitext(name)[1].lower()

    if ext == ".pdf":
        content = read_pdf(path)

    elif ext == ".docx":
        content = read_docx(path)

    elif ext in [".xlsx", ".xls"]:
        content = read_excel(path)

    elif ext in [".txt", ".csv"]:
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
        except:
            content = ""

    else:
        content = ""

    return content


# =========================
# TẢI GOOGLE DRIVE
# =========================
@st.cache_resource
def load_documents():

    temp_dir = tempfile.mkdtemp()

    url = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"

    gdown.download_folder(
        url=url,
        output=temp_dir,
        quiet=True,
        use_cookies=False
    )

    documents = []

    for root, dirs, files in os.walk(temp_dir):

        for file in files:

            path = os.path.join(root, file)

            ext = os.path.splitext(file)[1].lower()

            if ext in [".pdf", ".docx", ".xlsx", ".xls", ".txt", ".csv"]:

                try:

                    text = read_file(path)

                    if text.strip():

                        documents.append({
                            "file": file,
                            "text": text
                        })

                except Exception as e:

                    pass

    return documents


# =========================
# GIAO DIỆN
# =========================
st.title("🤖 CHATBOT TRA CỨU TÀI LIỆU IC PXVH1")

st.caption(
    "Tra cứu tài liệu PDF, Word, Excel từ Google Drive."
)

with st.spinner("Đang đọc tài liệu từ Google Drive..."):

    documents = load_documents()


if documents:

    st.success(
        f"Đã tải {len(documents)} tài liệu."
    )

else:

    st.error(
        "Không đọc được tài liệu. Kiểm tra quyền chia sẻ Google Drive."
    )


with st.expander("📁 Danh sách tài liệu"):

    for doc in documents:

        st.write("•", doc["file"])


# =========================
# CHAT
# =========================
question = st.chat_input(
    "Nhập nội dung cần tra cứu..."
)


if question:

    with st.chat_message("user"):

        st.write(question)


    # Gom nội dung tài liệu
    context_parts = []

    for doc in documents:

        text = doc["text"]

        # Giới hạn để tránh prompt quá lớn
        if len(text) > 30000:
            text = text[:30000]

        context_parts.append(
            f"""
=========================
TÊN FILE: {doc['file']}
=========================

{text}
"""
        )

    context = "\n".join(context_parts)


    prompt = f"""
Bạn là trợ lý tra cứu tài liệu kỹ thuật của PXVH1.

CHỈ trả lời dựa trên tài liệu được cung cấp dưới đây.

Yêu cầu:

1. Không tự suy diễn khi tài liệu không có thông tin.
2. Nếu không tìm thấy, trả lời:
   "Không tìm thấy nội dung này trong tài liệu hiện có."
3. Ưu tiên trả lời ngắn gọn, rõ ràng, đúng kỹ thuật.
4. Cuối câu trả lời phải ghi:
   Nguồn: tên file liên quan.
5. Nếu nội dung từ PDF có thông tin số trang thì ghi thêm trang.
6. Nếu nội dung từ Excel thì ghi tên file và Sheet nếu xác định được.

CÂU HỎI:

{question}


TÀI LIỆU:

{context}
"""


    with st.chat_message("assistant"):

        with st.spinner("Đang tra cứu tài liệu..."):

            try:

                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt
                )

                st.write(response.text)

            except Exception as e:

                st.error(
                    f"Lỗi Gemini API: {e}"
                )
