# =========================
# TẢI GOOGLE DRIVE
# =========================
@st.cache_resource
def load_documents():

    temp_dir = tempfile.mkdtemp()

    try:
        downloaded_files = gdown.download_folder(
            id=DRIVE_FOLDER_ID,
            output=temp_dir,
            quiet=False,
            use_cookies=False,
            remaining_ok=True
        )

    except Exception as e:
        st.error(f"Lỗi tải Google Drive: {e}")
        return []

    documents = []

    for root, dirs, files in os.walk(temp_dir):

        for file in files:

            path = os.path.join(root, file)
            ext = os.path.splitext(file)[1].lower()

            if ext in [".pdf", ".docx", ".xlsx", ".xls", ".txt", ".csv"]:

                try:
                    file_text = read_file(path)

                    if file_text.strip():
                        documents.append({
                            "file": file,
                            "text": file_text
                        })

                except Exception as e:
                    st.warning(f"Không đọc được file {file}: {e}")

    return documents
