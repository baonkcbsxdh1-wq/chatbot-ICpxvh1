import streamlit as st

st.set_page_config(page_title="Chatbot PXVH1", page_icon="🤖")

st.title("🤖 CHATBOT TRA CỨU TÀI LIỆU IC PXVH1")
st.write("Hệ thống đang được khởi tạo...")

question = st.text_input("Nhập câu hỏi cần tra cứu:")

if question:
    st.success(f"Bạn vừa hỏi: {question}")
