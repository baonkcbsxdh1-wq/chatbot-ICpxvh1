import os
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

st.title("🤖 CHATBOT TRA CỨU TÀI LIỆU IC PXVH1")
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
# TXT / CSV
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
# PHẦN HÌNH ẢNH
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


# Ưu tiên ảnh chụp nếu có
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
    # TẠO CONTEXT TÀI LIỆU
    # =====================================================
    context_parts = []


    for doc in documents:

        document_text = doc["text"]


        # Giới hạn mỗi file
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

Bạn có nhiệm vụ hỗ trợ tra cứu tài liệu kỹ thuật
và phân tích hình ảnh do người dùng cung cấp.

Hình ảnh có thể là:

- màn hình DCS/HMI;
- bảng liên động;
- sơ đồ kỹ thuật;
- nameplate thiết bị;
- mã TAG thiết bị;
- bảng thông số;
- trang tài liệu;
- hình chụp thiết bị;
- hình chụp lỗi hoặc alarm.

QUY TẮC:

1. Nếu có hình ảnh, hãy đọc kỹ chữ, mã TAG,
   thông số và nội dung kỹ thuật trong ảnh.

2. Sau đó đối chiếu nội dung trong ảnh với
   tài liệu được cung cấp bên dưới.

3. Chỉ kết luận dựa trên:
   - nội dung nhìn thấy trong ảnh;
   - tài liệu được cung cấp.

4. Không tự bịa thông tin.

5. Nếu không tìm thấy trong tài liệu, trả lời:
   "Không tìm thấy nội dung này trong tài liệu hiện có."

6. Nếu người dùng gửi mã TAG,
   hãy ưu tiên tìm chính xác mã TAG đó.

7. Nếu tìm thấy nhiều vị trí,
   hãy liệt kê đầy đủ các kết quả liên quan.

8. Với PDF:
   ghi tên file và số trang nếu xác định được.

9. Với Excel:
   ghi tên file và Sheet nếu xác định được.

10. Trả lời bằng tiếng Việt,
    rõ ràng và đúng thuật ngữ kỹ thuật.

11. Cuối câu trả lời ghi:

Nguồn:
- Tên tài liệu
- Trang hoặc Sheet nếu có.


CÂU HỎI CỦA NGƯỜI DÙNG:

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
            "Đang phân tích và tra cứu..."
        ):

            try:

                # ==========================================
                # CÓ HÌNH ẢNH
                # ==========================================
                if image_file is not None:

                    image_bytes = image_file.getvalue()

                    mime_type = image_file.type

                    image_part = types.Part.from_bytes(
                        data=image_bytes,
                        mime_type=mime_type
                    )


                    response = client.models.generate_content(
                        model="gemini-3.8-flash",
                        contents=[
                            prompt,
                            image_part
                        ]
                    )


                # ==========================================
                # KHÔNG CÓ HÌNH
                # ==========================================
                else:

                    response = client.models.generate_content(
                        model="gemini-3.8-flash",
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
