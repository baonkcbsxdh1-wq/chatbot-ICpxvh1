import os
import tempfile
import streamlit as st
import gdown
import pandas as pd

from pypdf import PdfReader
from docx import Document
from google import genai


# =========================================================
# CẤU HÌNH APP
# =========================================================
st.set_page_config(
    page_title="CHATBOT TRA CỨU TÀI LIỆU IC PXVH1",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 CHATBOT TRA CỨU TÀI LIỆU IC PXVH1")
st.caption("Tra cứu tài liệu PDF, Word, Excel từ Google Drive")


# =========================================================
# GOOGLE DRIVE FOLDER ID
# =========================================================
DRIVE_FOLDER_ID = "1P44hHly9bSdVZps4oqIgeReclxQWCzIm"


# =========================================================
# KẾT NỐI GEMINI
# =========================================================
try:
    client = genai.Client(
        api_key=st.secrets["GEMINI_API_KEY"]
    )
except Exception as e:
    st.error(f"Lỗi Gemini API Key: {e}")
    st.stop()


# =========================================================
# ĐỌC PDF
# =========================================================
def read_pdf(path):
    text = ""

    reader = PdfReader(path)

    for i, page in enumerate(reader.pages):
        try:
            page_text = page.extract_text() or ""

            if page_text.strip():
                text += f"\n\n===== TRANG {i + 1} =====\n"
                text += page_text

        except Exception:
            pass

    return text


# =========================================================
# ĐỌC WORD
# =========================================================
def read_docx(path):
    text = ""

    doc = Document(path)

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            text += paragraph.text + "\n"

    # Đọc bảng trong Word
    for table_index, table in enumerate(doc.tables):
        text += f"\n===== BẢNG {table_index + 1} =====\n"

        for row in table.rows:
            values = []

            for cell in row.cells:
                values.append(cell.text.strip())

            text += " | ".join(values) + "\n"

    return text


# =========================================================
# ĐỌC EXCEL
# =========================================================
def read_excel(path):
    text = ""

    excel_file = pd.ExcelFile(path)

    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(
                path,
                sheet_name=sheet_name,
                header=None
            )

            df = df.fillna("")

            text += f"\n\n===== SHEET: {sheet_name} =====\n"

            for index, row in df.iterrows():
                values = []

                for value in row.tolist():
                    value_text = str(value).strip()

                    if value_text:
                        values.append(value_text)

                if values:
                    text += (
                        f"Dòng {index + 1}: "
                        + " | ".join(values)
                        + "\n"
                    )

        except Exception as e:
            text += (
                f"\nKhông đọc được sheet {sheet_name}: {e}\n"
            )

    return text


# =========================================================
# ĐỌC TXT / CSV
# =========================================================
def read_text_file(path):
    try:
        with open(
            path,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as f:
            return f.read()

    except Exception:
        return ""


# =========================================================
# NHẬN DIỆN LOẠI FILE
# =========================================================
def read_file(path):
    extension = os.path.splitext(path)[1].lower()

    if extension == ".pdf":
        return read_pdf(path)

    elif extension == ".docx":
        return read_docx(path)

    elif extension in [".xlsx", ".xls"]:
        return read_excel(path)

    elif extension in [".txt", ".csv"]:
        return read_text_file(path)

    return ""


# =========================================================
# TẢI FILE TỪ GOOGLE DRIVE
# =========================================================
@st.cache_resource
def load_documents():

    documents = []

    temp_dir = tempfile.mkdtemp()

    try:
        downloaded_files = gdown.download_folder(
            id=DRIVE_FOLDER_ID,
            output=temp_dir,
            quiet=False,
            use_cookies=False
        )

    except Exception as e:
        st.error(
            f"Lỗi tải Google Drive: {e}"
        )
        return []


    if not downloaded_files:
        st.error(
            "Google Drive không trả về file nào."
        )
        return []


    for root, dirs, files in os.walk(temp_dir):

        for filename in files:

            path = os.path.join(
                root,
                filename
            )

            extension = os.path.splitext(
                filename
            )[1].lower()


            if extension not in [
                ".pdf",
                ".docx",
                ".xlsx",
                ".xls",
                ".txt",
                ".csv"
            ]:
                continue


            try:
                content = read_file(path)

                if content.strip():
                    documents.append(
                        {
                            "file": filename,
                            "text": content
                        }
                    )

            except Exception as e:
                st.warning(
                    f"Không đọc được file {filename}: {e}"
                )


    return documents


# =========================================================
# NẠP TÀI LIỆU
# =========================================================
with st.spinner(
    "Đang đọc tài liệu từ Google Drive..."
):
    documents = load_documents()


# =========================================================
# HIỂN THỊ TRẠNG THÁI
# =========================================================
if documents:
    st.success(
        f"✅ Đã đọc được {len(documents)} tài liệu."
    )

else:
    st.error(
        "❌ Chưa đọc được tài liệu từ Google Drive."
    )


# =========================================================
# DANH SÁCH FILE
# =========================================================
with st.expander(
    "📁 Danh sách tài liệu đã đọc"
):

    if documents:
        for doc in documents:
            st.write(
                "•",
                doc["file"]
            )

    else:
        st.write(
            "Chưa có tài liệu."
        )


# =========================================================
# LỊCH SỬ CHAT
# =========================================================
if "messages" not in st.session_state:
    st.session_state.messages = []


for message in st.session_state.messages:

    with st.chat_message(
        message["role"]
    ):
        st.markdown(
            message["content"]
        )


# =========================================================
# Ô CHAT
# =========================================================
question = st.chat_input(
    "Nhập nội dung cần tra cứu..."
)


if question:

    st.session_state.messages.append(
        {
            "role": "user",
            "content": question
        }
    )


    with st.chat_message("user"):
        st.markdown(question)


    # =====================================================
    # TẠO DỮ LIỆU NGỮ CẢNH
    # =====================================================
    context_parts = []


    for doc in documents:

        document_text = doc["text"]

        # Giới hạn mỗi tài liệu
        if len(document_text) > 40000:
            document_text = document_text[:40000]


        context_parts.append(
            f"""
==================================================
TÊN FILE: {doc["file"]}
==================================================

{document_text}
"""
        )


    context = "\n".join(
        context_parts
    )


    # =====================================================
    # PROMPT
    # =====================================================
    prompt = f"""
Bạn là CHATBOT TRA CỨU TÀI LIỆU IC PXVH1.

Nhiệm vụ:
Tra cứu thông tin trong các tài liệu PDF, Word, Excel
được cung cấp bên dưới.

QUY TẮC:

1. Chỉ trả lời dựa trên tài liệu được cung cấp.
2. Không tự bịa thông tin.
3. Nếu không tìm thấy, trả lời:
   "Không tìm thấy nội dung này trong tài liệu hiện có."
4. Trả lời ngắn gọn, rõ ràng, đúng kỹ thuật.
5. Nếu có nhiều nội dung liên quan thì trình bày theo từng ý.
6. Cuối câu trả lời phải ghi nguồn tài liệu.
7. Nếu là PDF, cố gắng ghi số trang.
8. Nếu là Excel, ghi tên file và tên Sheet nếu xác định được.
9. Có thể trích xuất thông số, điều kiện, liên động, bảo vệ.
10. Không sử dụng kiến thức bên ngoài thay thế tài liệu.

CÂU HỎI:

{question}


DỮ LIỆU TÀI LIỆU:

{context}
"""


    # =====================================================
    # GỌI GEMINI
    # =====================================================
    with st.chat_message(
        "assistant"
    ):

        with st.spinner(
            "Đang tra cứu..."
        ):

            try:
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )

                answer = response.text

                st.markdown(
                    answer
                )

                st.session_state.messages.append(
                    {
                        "role": "assistant",
                        "content": answer
                    }
                )

            except Exception as e:
                st.error(
                    f"Lỗi Gemini API: {e}"
                )
