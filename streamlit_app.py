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
st.caption("Tra cứu PDF, Word, Excel và hình ảnh từ Google Drive")


# =========================================================
# GOOGLE DRIVE
# =========================================================
DRIVE_FOLDER_ID = "1P44hHly9bSdVZps4oqIgeReclxQWCzIm"


# =========================================================
# TỐI ƯU QUOTA GEMINI
# =========================================================
MAX_RESULTS = 6
MAX_CHUNK_CHARS = 1800
MAX_CONTEXT_CHARS = 12000


# =========================================================
# GEMINI
# =========================================================
try:
    client = genai.Client(
        api_key=st.secrets["GEMINI_API_KEY"]
    )
except Exception as e:
    st.error(f"Lỗi Gemini API Key: {e}")
    st.stop()


# =========================================================
# CHUẨN HÓA TEXT
# =========================================================
def normalize_text(text):
    text = str(text).lower()

    text = unicodedata.normalize("NFD", text)

    text = "".join(
        ch for ch in text
        if unicodedata.category(ch) != "Mn"
    )

    text = text.replace("đ", "d")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


# =========================================================
# TỪ KHÓA
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
        "do",
        "hay",
        "liet",
        "ke"
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

    tags = []

    for item in candidates:
        if (
            re.search(r"[A-Za-z]", item)
            and re.search(r"\d", item)
        ):
            tags.append(item.upper())

    return list(dict.fromkeys(tags))


# =========================================================
# EXCEL COLUMN A/B/C...
# =========================================================
def excel_column_name(number):
    result = ""

    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result

    return result


# =========================================================
# PDF
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
            x.strip()
            for x in text.splitlines()
            if x.strip()
        ]

        block = []

        for line in lines:
            block.append(line)

            if len("\n".join(block)) >= 1600:
                chunks.append({
                    "file": filename,
                    "location": f"Trang {page_index + 1}",
                    "text": "\n".join(block)
                })

                block = []

        if block:
            chunks.append({
                "file": filename,
                "location": f"Trang {page_index + 1}",
                "text": "\n".join(block)
            })

    return chunks


# =========================================================
# WORD
# =========================================================
def read_docx_chunks(path, filename):
    chunks = []

    doc = Document(path)

    block = []
    block_no = 1

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        block.append(text)

        if len("\n".join(block)) >= 1600:
            chunks.append({
                "file": filename,
                "location": f"Nội dung {block_no}",
                "text": "\n".join(block)
            })

            block = []
            block_no += 1

    if block:
        chunks.append({
            "file": filename,
            "location": f"Nội dung {block_no}",
            "text": "\n".join(block)
        })

    for table_index, table in enumerate(doc.tables):
        for row_index, row in enumerate(table.rows):
            values = [
                cell.text.strip()
                for cell in row.cells
            ]

            row_text = " | ".join(values)

            if row_text.strip():
                chunks.append({
                    "file": filename,
                    "location": (
                        f"Bảng {table_index + 1} | "
                        f"Dòng {row_index + 1}"
                    ),
                    "text": row_text
                })

    return chunks


# =========================================================
# EXCEL
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

            for col_index, value in enumerate(row.tolist()):
                value_text = str(value).strip()

                if (
                    value_text
                    and value_text.lower() != "nan"
                ):
                    col = excel_column_name(col_index + 1)

                    values.append(
                        f"{col}={value_text}"
                    )

            if values:
                chunks.append({
                    "file": filename,
                    "location": (
                        f"Sheet: {sheet_name} | "
                        f"Dòng: {row_index + 1}"
                    ),
                    "text": " | ".join(values)
                })

    return chunks


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
        x.strip()
        for x in text.splitlines()
        if x.strip()
    ]

    for index, line in enumerate(lines):
        chunks.append({
            "file": filename,
            "location": f"Dòng {index + 1}",
            "text": line
        })

    return chunks


# =========================================================
# ĐỌC FILE
# =========================================================
def read_file_chunks(path, filename):
    ext = os.path.splitext(filename)[1].lower()

    if ext == ".pdf":
        return read_pdf_chunks(path, filename)

    if ext == ".docx":
        return read_docx_chunks(path, filename)

    if ext in [".xlsx", ".xls"]:
        return read_excel_chunks(path, filename)

    if ext in [".txt", ".csv"]:
        return read_text_chunks(path, filename)

    return []


# =========================================================
# GOOGLE DRIVE
# KHÔNG TỰ UPDATE
# =========================================================
@st.cache_data(show_spinner=False)
def load_documents():
    chunks = []
    files_loaded = []

    temp_dir = tempfile.mkdtemp()

    try:
        downloaded = gdown.download_folder(
            id=DRIVE_FOLDER_ID,
            output=temp_dir,
            quiet=True,
            use_cookies=False
        )

    except Exception as e:
        return [], [], f"Lỗi tải Google Drive: {e}"

    if not downloaded:
        return [], [], "Google Drive không trả về file nào."

    for root, dirs, files in os.walk(temp_dir):
        for filename in files:
            ext = os.path.splitext(filename)[1].lower()

            if ext not in [
                ".pdf",
                ".docx",
                ".xlsx",
                ".xls",
                ".txt",
                ".csv"
            ]:
                continue

            path = os.path.join(root, filename)

            try:
                file_chunks = read_file_chunks(
                    path,
                    filename
                )

                if file_chunks:
                    chunks.extend(file_chunks)
                    files_loaded.append(filename)

            except Exception:
                pass

    files_loaded = list(
        dict.fromkeys(files_loaded)
    )

    return chunks, files_loaded, None


# =========================================================
# NÚT UPDATE THỦ CÔNG
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
# LOAD DATA
# =========================================================
with st.spinner(
    "Đang đọc tài liệu..."
):
    chunks, files_loaded, drive_error = (
        load_documents()
    )


if drive_error:
    st.error(drive_error)

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
# UPLOAD ẢNH
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
# CHẤM ĐIỂM KẾT QUẢ
# =========================================================
def score_chunk(
    chunk,
    question,
    keywords,
    tags
):
    text_norm = normalize_text(
        chunk["text"]
    )

    question_norm = normalize_text(
        question
    )

    score = 0

    # TAG exact
    for tag in tags:
        if tag.lower() in chunk["text"].lower():
            score += 150

    # nguyên câu
    if (
        question_norm
        and question_norm in text_norm
    ):
        score += 80

    # từ khóa
    matched = 0

    for keyword in keywords:
        if keyword in text_norm:
            matched += 1
            score += 12

    if matched >= 2:
        score += matched * 6

    # range/value
    if (
        "range" in question_norm
        or "value" in question_norm
        or "dai do" in question_norm
        or "gia tri" in question_norm
    ):
        if any(
            x in text_norm
            for x in [
                "range",
                "value",
                "ma",
                "vdc",
                "bar",
                "mpa",
                "kpa",
                "mm",
                "°c",
                "%",
                "rpm"
            ]
        ):
            score += 15

    return score


# =========================================================
# RETRIEVAL LOCAL
# =========================================================
def retrieve(
    question,
    chunks
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
        key=lambda x: x[0],
        reverse=True
    )

    selected = []

    for score, chunk in ranked[:MAX_RESULTS]:
        selected.append({
            "score": score,
            **chunk
        })

    return selected


# =========================================================
# CONTEXT TỐI ƯU
# =========================================================
def build_context(results):
    parts = []
    total = 0

    for index, result in enumerate(
        results,
        start=1
    ):
        text = result["text"]

        if len(text) > MAX_CHUNK_CHARS:
            text = text[:MAX_CHUNK_CHARS]

        block = f"""
===== KẾT QUẢ {index} =====
FILE: {result["file"]}
VỊ TRÍ: {result["location"]}

{text}
"""

        if total + len(block) > MAX_CONTEXT_CHARS:
            break

        parts.append(block)

        total += len(block)

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
            or "RESOURCE_EXHAUSTED" in error_text
        ):
            raise Exception(
                "Đã chạm giới hạn Gemini miễn phí. "
                "Hãy chờ khoảng 1 phút rồi hỏi lại."
            )

        if (
            "503" in error_text
            or "UNAVAILABLE" in error_text
        ):
            time.sleep(3)

            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=contents
            )

            return response.text

        raise e


# =========================================================
# ĐỌC ẢNH ĐỂ LẤY TAG
# =========================================================
def analyze_image(
    image_file,
    question
):
    image_part = types.Part.from_bytes(
        data=image_file.getvalue(),
        mime_type=image_file.type
    )

    prompt = f"""
Đọc ảnh kỹ thuật này.

Chỉ trả về:
- mã TAG nhìn thấy;
- tên thiết bị;
- từ khóa kỹ thuật chính.

Không giải thích dài.

Câu hỏi:
{question}
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
# HISTORY
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
    st.session_state.messages.append({
        "role": "user",
        "content": question
    })

    with st.chat_message("user"):
        st.markdown(question)


    # =====================================================
    # CÂU SEARCH
    # =====================================================
    search_question = question

    if image_file is not None:
        with st.spinner(
            "Đang đọc thông tin trong ảnh..."
        ):
            image_keywords = analyze_image(
                image_file,
                question
            )

        if image_keywords:
            search_question = (
                question
                + " "
                + image_keywords
            )


    # =====================================================
    # TÌM LOCAL
    # =====================================================
    results = retrieve(
        search_question,
        chunks
    )


    # =====================================================
    # KHÔNG TÌM THẤY
    # =====================================================
    if not results:
        answer = (
            "Không tìm thấy nội dung này "
            "trong tài liệu hiện có."
        )

        with st.chat_message("assistant"):
            st.markdown(answer)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })


    # =====================================================
    # CÓ KẾT QUẢ
    # =====================================================
    else:
        context = build_context(
            results
        )

        prompt = f"""
Bạn là CHATBOT TRA CỨU RANGE/VALUE IC PXVH1.

CÂU HỎI:

{question}


Python đã tìm sẵn các kết quả liên quan dưới đây.

CHỈ dùng dữ liệu này để trả lời.

{context}


YÊU CẦU:

1. Trả lời bằng tiếng Việt.
2. Trả lời trực tiếp câu hỏi.
3. Không tự bịa dữ liệu.
4. Nếu hỏi RANGE/VALUE:
   nêu rõ giá trị và đơn vị.
5. Nếu hỏi TAG:
   chỉ trả lời đúng TAG liên quan.
6. Nếu nhiều kết quả:
   liệt kê từng kết quả.
7. Với Excel:
   ghi File, Sheet, dòng.
8. Với PDF:
   ghi File, Trang.
9. Không đưa nguồn không liên quan.
10. Trả lời ngắn gọn, kỹ thuật.
11. Cuối câu trả lời phải có:

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

                    else:
                        answer = call_gemini(
                            prompt
                        )

                    st.markdown(answer)

                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": answer
                    })

                except Exception as e:
                    st.error(
                        str(e)
                    )


        # =================================================
        # XEM KẾT QUẢ LOCAL
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
