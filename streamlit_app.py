import os
import time
import tempfile
import streamlit as st
import gdown
import pandas as pd

from pypdf import PdfReader
from docx import Document
from google import genai
from google.genai import types


# =========================================================
# CẤU HÌNH APP
# =========================================================
st.set_page_config(
    page_title="CHATBOT TRA CỨU TÀI LIỆU IC PXVH1",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 CHATBOT TRA CỨU RANGE/VALUE IC PXVH1")
st.caption("Tra cứu PDF, Word, Excel và hình ảnh từ Google Drive")


# =========================================================
# GOOGLE DRIVE
# =========================================================
DRIVE_FOLDER_ID = "1P44hHly9bSdVZps4oqIgeReclxQWCzIm"


# =========================================================
# GEMINI API
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

    # Đọc đoạn văn
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            text += paragraph.text + "\n"

    # Đọc bảng
    for table_index, table in enumerate(doc.tables):
        text += f"\n===== BẢNG {table_index + 1} =====\n"

        for row in table.rows:
            values = []

            for cell in row.cells:
                values.append(
                    cell.text.strip()
                )

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

            text += (
                f"\n\n"
                f"===== SHEET: {sheet_name} =====\n"
            )

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
                f"\nKhông đọc được Sheet "
                f"{sheet_name}: {e}\n"
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
# NHẬN DIỆN FILE
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
# TẢI GOOGLE DRIVE
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
                    f"Không đọc được file "
                    f"{filename}: {e}"
                )

    return documents


# =========================================================
# TÌM CHÍNH XÁC TAG / TỪ KHÓA
# =========================================================
def find_exact_matches(question, documents):
    question_clean = question.strip().upper()

    matches = []

    if not question_clean:
        return matches

    for doc in documents:
        text_upper = doc["text"].upper()

        if question_clean in text_upper:
            matches.append(doc)

    return matches


# =========================================================
# RETRY GEMINI KHI 503
# =========================================================
def ask_gemini(contents):
    last_error = None

    for attempt in range(4):
        try:
            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=contents
            )

            return response.text

        except Exception as e:
            last_error = e

            error_text = str(e)

            if (
                "503" in error_text
                or "UNAVAILABLE" in error_text
                or "high demand" in error_text
            ):
                wait_time = 3 + attempt * 2

                time.sleep(wait_time)
                continue

            raise e

    raise last_error


# =========================================================
# NẠP TÀI LIỆU
# =========================================================
with st.spinner(
    "Đang đọc tài liệu từ Google Drive..."
):
    documents = load_documents()


# =========================================================
# TRẠNG THÁI
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
# DANH SÁCH TÀI LIỆU
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
# HÌNH ẢNH
# =========================================================
st.divider()

st.subheader("📷 Tra cứu bằng hình ảnh")

col1, col2 = st.columns(2)


with col1:
    uploaded_image = st.file_uploader(
        "🖼️ Chọn ảnh từ máy",
        type=[
            "jpg",
            "jpeg",
            "png",
            "webp"
        ]
    )


with col2:
    camera_image = st.camera_input(
        "📷 Hoặc chụp ảnh trực tiếp"
    )


image_file = None

if camera_image is not None:
    image_file = camera_image

elif uploaded_image is not None:
    image_file = uploaded_image


if image_file is not None:
    st.image(
        image_file,
        caption="Ảnh dùng để tra cứu",
        width=500
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
    # ƯU TIÊN TÌM CHÍNH XÁC
    # =====================================================
    exact_matches = find_exact_matches(
        question,
        documents
    )


    # =====================================================
    # TẠO CONTEXT
    # =====================================================
    context_parts = []


    # Nếu có kết quả chính xác, ưu tiên đưa lên đầu
    if exact_matches:

        for doc in exact_matches:

            context_parts.append(
                f"""
==================================================
KẾT QUẢ KHỚP CHÍNH XÁC
TÊN FILE: {doc["file"]}
==================================================

{doc["text"]}
"""
            )


    # Sau đó thêm toàn bộ tài liệu còn lại
    for doc in documents:

        if doc not in exact_matches:

            context_parts.append(
                f"""
==================================================
TÊN FILE: {doc["file"]}
==================================================

{doc["text"]}
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

Bạn hỗ trợ tra cứu tài liệu kỹ thuật
PDF, Word, Excel và hình ảnh.

QUY TẮC BẮT BUỘC:

1. Chỉ trả lời dựa trên nội dung tài liệu
   hoặc hình ảnh người dùng cung cấp.

2. Không được tự bịa hoặc tự suy diễn.

3. Nếu người dùng nhập mã TAG như:
   A0HTG14CP004
   10HLS11CP401H
   thì phải ưu tiên tìm chính xác toàn bộ mã TAG.

4. Nếu tìm thấy mã TAG:
   - ghi đầy đủ thông tin liên quan;
   - ghi tên file;
   - ghi Sheet nếu là Excel;
   - ghi trang nếu là PDF.

5. Nếu một TAG xuất hiện ở nhiều Sheet
   hoặc nhiều file, phải liệt kê tất cả.

6. Không được nói "không tìm thấy"
   nếu mã đó thực tế có trong dữ liệu.

7. Nếu không tìm thấy thật sự thì trả lời:
   "Không tìm thấy nội dung này trong tài liệu hiện có."

8. Với Excel:
   chú ý các tiêu đề:
   ===== SHEET: ... =====
   và phải xác định đúng Sheet.

9. Với PDF:
   chú ý:
   ===== TRANG ... =====
   và ghi đúng số trang nếu có.

10. Nếu có hình ảnh:
    - đọc mã TAG;
    - đọc chữ;
    - đọc thông số;
    - đối chiếu với tài liệu.

11. Trả lời bằng tiếng Việt,
    rõ ràng, ngắn gọn, đúng kỹ thuật.

12. Cuối câu trả lời phải có:

Nguồn:
- Tên file
- Sheet hoặc Trang nếu xác định được


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

                if image_file is not None:

                    image_bytes = image_file.getvalue()

                    image_part = types.Part.from_bytes(
                        data=image_bytes,
                        mime_type=image_file.type
                    )

                    answer = ask_gemini(
                        [
                            prompt,
                            image_part
                        ]
                    )

                else:

                    answer = ask_gemini(
                        prompt
                    )


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
                    f"Lỗi Gemini API sau khi thử lại: {e}"
                )
