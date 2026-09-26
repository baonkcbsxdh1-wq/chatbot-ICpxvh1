import os
import re
import time
import tempfile
import unicodedata
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
    page_title="CHATBOT TRA CỨU RANGE/VALUE IC PXVH1",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 CHATBOT TRA CỨU RANGE/VALUE IC PXVH1")
st.caption(
    "Tra cứu PDF, Word, Excel và hình ảnh từ Google Drive"
)


# =========================================================
# GOOGLE DRIVE
# =========================================================
DRIVE_FOLDER_ID = "1P44hHly9bSdVZps4oqIgeReclxQWCzIm"


# =========================================================
# GIỚI HẠN DỮ LIỆU GỬI GEMINI
# =========================================================

# Số đoạn liên quan tối đa gửi Gemini
MAX_RESULTS = 12

# Mỗi đoạn tối đa bao nhiêu ký tự
MAX_CHUNK_CHARS = 2500

# Tổng context tối đa
MAX_CONTEXT_CHARS = 28000


# =========================================================
# GEMINI
# =========================================================
try:
    client = genai.Client(
        api_key=st.secrets["GEMINI_API_KEY"]
    )

except Exception as e:
    st.error(
        f"Lỗi Gemini API Key: {e}"
    )
    st.stop()


# =========================================================
# CHUẨN HÓA CHỮ ĐỂ TÌM KIẾM
# =========================================================
def normalize_text(text):

    text = str(text).lower()

    text = unicodedata.normalize(
        "NFD",
        text
    )

    text = "".join(
        char
        for char in text
        if unicodedata.category(char) != "Mn"
    )

    text = text.replace("đ", "d")

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================================================
# TÁCH TỪ KHÓA
# =========================================================
def get_keywords(question):

    normalized = normalize_text(question)

    words = re.findall(
        r"[a-z0-9_.\-/]+",
        normalized
    )

    stop_words = {
        "tim",
        "giup",
        "toi",
        "cho",
        "biet",
        "thong",
        "tin",
        "cua",
        "ve",
        "la",
        "bao",
        "nhieu",
        "nhung",
        "nao",
        "trong",
        "tai",
        "lieu",
        "range",
        "value",
        "gia",
        "tri",
        "thiet",
        "bi",
        "do"
    }

    keywords = []

    for word in words:

        if len(word) >= 2 and word not in stop_words:
            keywords.append(word)

    return list(dict.fromkeys(keywords))


# =========================================================
# NHẬN DIỆN TAG
# =========================================================
def extract_possible_tags(text):

    candidates = re.findall(
        r"\b[A-Za-z0-9][A-Za-z0-9._\-/]{5,}\b",
        text
    )

    result = []

    for item in candidates:

        has_letter = bool(
            re.search(r"[A-Za-z]", item)
        )

        has_number = bool(
            re.search(r"\d", item)
        )

        if has_letter and has_number:
            result.append(
                item.upper()
            )

    return list(
        dict.fromkeys(result)
    )


# =========================================================
# ĐỌC PDF THÀNH CÁC ĐOẠN
# =========================================================
def read_pdf_chunks(path, filename):

    chunks = []

    reader = PdfReader(path)

    for page_index, page in enumerate(reader.pages):

        try:
            text = page.extract_text() or ""

        except Exception:
            text = ""

        if not text.strip():
            continue

        lines = [
            line.strip()
            for line in text.splitlines()
            if line.strip()
        ]

        block = []

        for line in lines:

            block.append(line)

            if len("\n".join(block)) >= 1800:

                chunks.append(
                    {
                        "file": filename,
                        "location": f"Trang {page_index + 1}",
                        "text": "\n".join(block)
                    }
                )

                block = []

        if block:

            chunks.append(
                {
                    "file": filename,
                    "location": f"Trang {page_index + 1}",
                    "text": "\n".join(block)
                }
            )

    return chunks


# =========================================================
# ĐỌC WORD THÀNH CÁC ĐOẠN
# =========================================================
def read_docx_chunks(path, filename):

    chunks = []

    doc = Document(path)

    block = []

    block_number = 1

    for paragraph in doc.paragraphs:

        text = paragraph.text.strip()

        if not text:
            continue

        block.append(text)

        if len("\n".join(block)) >= 1800:

            chunks.append(
                {
                    "file": filename,
                    "location": f"Nội dung {block_number}",
                    "text": "\n".join(block)
                }
            )

            block = []

            block_number += 1

    if block:

        chunks.append(
            {
                "file": filename,
                "location": f"Nội dung {block_number}",
                "text": "\n".join(block)
            }
        )


    # Đọc bảng Word
    for table_index, table in enumerate(doc.tables):

        rows = []

        for row_index, row in enumerate(table.rows):

            values = [
                cell.text.strip()
                for cell in row.cells
            ]

            row_text = " | ".join(values)

            if row_text.strip():
                rows.append(
                    f"Dòng {row_index + 1}: {row_text}"
                )

        if rows:

            chunks.append(
                {
                    "file": filename,
                    "location": f"Bảng {table_index + 1}",
                    "text": "\n".join(rows)
                }
            )

    return chunks


# =========================================================
# ĐỌC EXCEL
# MỖI DÒNG TRỞ THÀNH 1 ĐƠN VỊ TRA CỨU
# =========================================================
def read_excel_chunks(path, filename):

    chunks = []

    excel_file = pd.ExcelFile(path)

    for sheet_name in excel_file.sheet_names:

        try:
            df = pd.read_excel(
                path,
                sheet_name=sheet_name,
                header=None
            )

            df = df.fillna("")

        except Exception:
            continue


        for row_index, row in df.iterrows():

            values = []

            for col_index, value in enumerate(
                row.tolist()
            ):

                value_text = str(value).strip()

                if (
                    value_text
                    and value_text.lower() != "nan"
                ):

                    column_name = get_excel_column_name(
                        col_index + 1
                    )

                    values.append(
                        f"{column_name}={value_text}"
                    )


            if values:

                chunks.append(
                    {
                        "file": filename,
                        "location": (
                            f"Sheet: {sheet_name} | "
                            f"Dòng: {row_index + 1}"
                        ),
                        "text": " | ".join(values)
                    }
                )

    return chunks


# =========================================================
# ĐỔI SỐ CỘT THÀNH A, B, C...
# =========================================================
def get_excel_column_name(number):

    result = ""

    while number:

        number, remainder = divmod(
            number - 1,
            26
        )

        result = (
            chr(65 + remainder)
            + result
        )

    return result


# =========================================================
# TXT / CSV
# =========================================================
def read_text_chunks(path, filename):

    chunks = []

    try:

        with open(
            path,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as f:

            text = f.read()

    except Exception:

        return []


    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    block = []

    block_number = 1

    for line in lines:

        block.append(line)

        if len("\n".join(block)) >= 1800:

            chunks.append(
                {
                    "file": filename,
                    "location": f"Đoạn {block_number}",
                    "text": "\n".join(block)
                }
            )

            block = []

            block_number += 1


    if block:

        chunks.append(
            {
                "file": filename,
                "location": f"Đoạn {block_number}",
                "text": "\n".join(block)
            }
        )

    return chunks


# =========================================================
# ĐỌC FILE
# =========================================================
def read_file_chunks(path, filename):

    extension = os.path.splitext(
        filename
    )[1].lower()


    if extension == ".pdf":

        return read_pdf_chunks(
            path,
            filename
        )


    elif extension == ".docx":

        return read_docx_chunks(
            path,
            filename
        )


    elif extension in [
        ".xlsx",
        ".xls"
    ]:

        return read_excel_chunks(
            path,
            filename
        )


    elif extension in [
        ".txt",
        ".csv"
    ]:

        return read_text_chunks(
            path,
            filename
        )


    return []


# =========================================================
# TẢI GOOGLE DRIVE
#
# KHÔNG CÓ TTL
# CHỈ CẬP NHẬT KHI BẤM NÚT
# =========================================================
@st.cache_data(
    show_spinner=False
)
def load_documents():

    all_chunks = []

    files_loaded = []

    temp_dir = tempfile.mkdtemp()


    try:

        downloaded_files = gdown.download_folder(
            id=DRIVE_FOLDER_ID,
            output=temp_dir,
            quiet=True,
            use_cookies=False
        )


    except Exception as e:

        return (
            [],
            [],
            f"Lỗi tải Google Drive: {e}"
        )


    if not downloaded_files:

        return (
            [],
            [],
            "Google Drive không trả về file nào."
        )


    for root, dirs, files in os.walk(
        temp_dir
    ):

        for filename in files:

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


            path = os.path.join(
                root,
                filename
            )


            try:

                chunks = read_file_chunks(
                    path,
                    filename
                )


                if chunks:

                    all_chunks.extend(
                        chunks
                    )

                    files_loaded.append(
                        filename
                    )


            except Exception:
                pass


    files_loaded = list(
        dict.fromkeys(files_loaded)
    )


    return (
        all_chunks,
        files_loaded,
        None
    )


# =========================================================
# NÚT CẬP NHẬT DRIVE
# =========================================================
if st.button(
    "🔄 Cập nhật tài liệu từ Google Drive"
):

    st.cache_data.clear()

    st.success(
        "Đang cập nhật lại tài liệu..."
    )

    time.sleep(0.4)

    st.rerun()


# =========================================================
# NẠP DỮ LIỆU
# =========================================================
with st.spinner(
    "Đang đọc tài liệu..."
):

    chunks, files_loaded, drive_error = (
        load_documents()
    )


# =========================================================
# HIỂN THỊ TRẠNG THÁI
# =========================================================
if drive_error:

    st.error(
        drive_error
    )


elif files_loaded:

    st.success(
        f"✅ Đã đọc được "
        f"{len(files_loaded)} tài liệu "
        f"({len(chunks)} vùng dữ liệu)."
    )


else:

    st.error(
        "❌ Chưa đọc được tài liệu."
    )


# =========================================================
# DANH SÁCH FILE
# =========================================================
with st.expander(
    "📁 Danh sách tài liệu đã đọc"
):

    for filename in files_loaded:

        st.write(
            "•",
            filename
        )


# =========================================================
# TRA CỨU HÌNH ẢNH
# =========================================================
st.divider()

st.subheader(
    "🖼️ Tra cứu bằng hình ảnh"
)


uploaded_image = st.file_uploader(
    "Chọn ảnh từ máy",
    type=[
        "jpg",
        "jpeg",
        "png",
        "webp"
    ]
)


image_file = uploaded_image


if image_file is not None:

    st.image(
        image_file,
        caption="Ảnh dùng để tra cứu",
        width=500
    )


# =========================================================
# CHẤM ĐIỂM MỨC LIÊN QUAN
# =========================================================
def score_chunk(
    chunk,
    question,
    keywords,
    tags
):

    text_normalized = normalize_text(
        chunk["text"]
    )

    question_normalized = normalize_text(
        question
    )


    score = 0


    # ---------------------------------------------
    # ƯU TIÊN TAG
    # ---------------------------------------------
    for tag in tags:

        if tag.lower() in chunk["text"].lower():

            score += 100


    # ---------------------------------------------
    # KHỚP NGUYÊN CÂU
    # ---------------------------------------------
    if (
        question_normalized
        and question_normalized
        in text_normalized
    ):

        score += 80


    # ---------------------------------------------
    # KHỚP TỪ KHÓA
    # ---------------------------------------------
    matched_keywords = 0

    for keyword in keywords:

        if keyword in text_normalized:

            matched_keywords += 1

            score += 10


    # Nhiều từ cùng xuất hiện thì cộng thêm
    if matched_keywords >= 2:

        score += (
            matched_keywords * 5
        )


    # ---------------------------------------------
    # RANGE / VALUE
    # ---------------------------------------------
    question_lower = question_normalized

    if (
        "range" in question_lower
        or "value" in question_lower
        or "dai do" in question_lower
        or "gia tri" in question_lower
    ):

        if (
            "range" in text_normalized
            or "value" in text_normalized
            or "ma" in text_normalized
            or "bar" in text_normalized
            or "mpa" in text_normalized
            or "kpa" in text_normalized
            or "mm" in text_normalized
            or "%" in text_normalized
        ):

            score += 10


    return score


# =========================================================
# TÌM CÁC ĐOẠN PHÙ HỢP
# =========================================================
def retrieve_relevant_chunks(
    question,
    chunks,
    max_results=MAX_RESULTS
):

    keywords = get_keywords(
        question
    )

    tags = extract_possible_tags(
        question
    )


    ranked = []


    for chunk in chunks:

        score = score_chunk(
            chunk,
            question,
            keywords,
            tags
        )


        if score > 0:

            ranked.append(
                (
                    score,
                    chunk
                )
            )


    ranked.sort(
        key=lambda item: item[0],
        reverse=True
    )


    selected = []


    for score, chunk in ranked:

        selected.append(
            {
                "score": score,
                **chunk
            }
        )


        if len(selected) >= max_results:

            break


    return selected


# =========================================================
# TẠO CONTEXT NHỎ GỬI GEMINI
# =========================================================
def build_context(results):

    parts = []

    total_chars = 0


    for number, result in enumerate(
        results,
        start=1
    ):

        text = result["text"]


        if len(text) > MAX_CHUNK_CHARS:

            text = text[
                :MAX_CHUNK_CHARS
            ]


        block = f"""
===== KẾT QUẢ {number} =====
FILE: {result["file"]}
VỊ TRÍ: {result["location"]}

{text}
"""


        if (
            total_chars
            + len(block)
            > MAX_CONTEXT_CHARS
        ):

            break


        parts.append(block)

        total_chars += len(block)


    return "\n".join(parts)


# =========================================================
# GEMINI
# =========================================================
def call_gemini(contents):

    try:

        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=contents
        )

        return response.text


    except Exception as e:

        error_text = str(e)


        if (
            "429" in error_text
            or "RESOURCE_EXHAUSTED"
            in error_text
        ):

            raise Exception(
                "Đã chạm giới hạn Gemini miễn phí. "
                "Hãy chờ khoảng 1 phút rồi hỏi lại."
            )


        elif (
            "503" in error_text
            or "UNAVAILABLE"
            in error_text
        ):

            time.sleep(3)

            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=contents
            )

            return response.text


        raise e


# =========================================================
# ĐỌC HÌNH ĐỂ LẤY TAG / TỪ KHÓA
# =========================================================
def analyze_image_for_search(
    image_file,
    user_question
):

    image_part = types.Part.from_bytes(
        data=image_file.getvalue(),
        mime_type=image_file.type
    )


    prompt = f"""
Hãy đọc hình ảnh kỹ thuật này.

Chỉ thực hiện nhiệm vụ sau:

1. Đọc các mã TAG thiết bị nhìn thấy.
2. Đọc các từ khóa kỹ thuật quan trọng.
3. Đọc tên thiết bị nếu nhìn thấy.
4. Không giải thích dài dòng.

Câu hỏi của người dùng:
{user_question}

Hãy trả về một dòng ngắn chứa
TAG và từ khóa dùng để tìm tài liệu.
"""


    try:

        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=[
                prompt,
                image_part
            ]
        )

        return response.text


    except Exception:

        return ""


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
# CHAT
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

        st.markdown(
            question
        )


    # =====================================================
    # CÂU TÌM KIẾM
    # =====================================================
    search_question = question


    # Nếu có ảnh:
    # Gemini đọc ảnh trước để lấy TAG
    if image_file is not None:

        with st.spinner(
            "Đang đọc thông tin trong ảnh..."
        ):

            image_keywords = (
                analyze_image_for_search(
                    image_file,
                    question
                )
            )


        if image_keywords:

            search_question = (
                question
                + " "
                + image_keywords
            )


    # =====================================================
    # TÌM LOCAL TRƯỚC
    # =====================================================
    results = retrieve_relevant_chunks(
        search_question,
        chunks
    )


    # =====================================================
    # KHÔNG CÓ KẾT QUẢ
    # =====================================================
    if not results:

        answer = (
            "Không tìm thấy nội dung này "
            "trong tài liệu hiện có."
        )


        with st.chat_message(
            "assistant"
        ):

            st.markdown(
                answer
            )


        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": answer
            }
        )


    # =====================================================
    # CÓ KẾT QUẢ
    # =====================================================
    else:

        context = build_context(
            results
        )


        prompt = f"""
Bạn là CHATBOT TRA CỨU RANGE/VALUE IC PXVH1.

Người dùng hỏi:

{question}


Dưới đây là các kết quả đã được
Python tìm trước trong tài liệu.

CHỈ sử dụng các kết quả này để trả lời.

{context}


YÊU CẦU:

1. Trả lời bằng tiếng Việt.
2. Ưu tiên trả lời trực tiếp câu hỏi.
3. Không tự bịa dữ liệu.
4. Nếu hỏi RANGE/VALUE:
   phải nêu chính xác giá trị và đơn vị
   nếu tài liệu có.
5. Nếu hỏi TAG:
   nêu thông tin liên quan đến đúng TAG.
6. Nếu có nhiều kết quả phù hợp,
   liệt kê rõ từng kết quả.
7. Với Excel:
   ghi tên file, Sheet và dòng.
8. Với PDF:
   ghi tên file và Trang.
9. Không đưa nguồn không liên quan.
10. Cuối câu trả lời phải có mục:

Nguồn:
- ...
"""


        with st.chat_message(
            "assistant"
        ):

            with st.spinner(
                "Đang tra cứu..."
            ):

                try:

                    # =====================================
                    # CÓ ẢNH
                    # =====================================
                    if image_file is not None:

                        image_part = (
                            types.Part.from_bytes(
                                data=image_file.getvalue(),
                                mime_type=image_file.type
                            )
                        )


                        answer = call_gemini(
                            [
                                prompt,
                                image_part
                            ]
                        )


                    # =====================================
                    # KHÔNG CÓ ẢNH
                    # =====================================
                    else:

                        answer = call_gemini(
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
                        str(e)
                    )


        # =================================================
        # CHO PHÉP XEM KẾT QUẢ PYTHON TÌM ĐƯỢC
        # =================================================
        with st.expander(
            "🔎 Xem các vị trí chatbot đã tìm thấy"
        ):

            for i, result in enumerate(
                results,
                start=1
            ):

                st.markdown(
                    f"**{i}. {result['file']}**"
                )

                st.caption(
                    result["location"]
                )

                st.text(
                    result["text"][:1000]
                )

                st.divider()
