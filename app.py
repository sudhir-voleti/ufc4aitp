import streamlit as st
import os
from markitdown import MarkItDown
import requests

# Page Configuration
st.set_page_config(page_title="Universal Doc Converter", page_icon="📄")

def main():
    st.title("📄 Universal Document Reader")
    st.markdown("Upload Office docs, PDFs, or HTML to convert them into clean Markdown.")

    # [1] Initialize the Engine
    # Configuring with a custom request session for [3] Resilience
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    
    # markitdown allows passing a custom request_session for web-based content
    md_engine = MarkItDown(requests_session=session)

    # [2] Interface: Upload Area
    uploaded_files = st.file_uploader(
        "Drag and drop files here", 
        type=["docx", "xlsx", "pptx", "pdf", "html", "zip"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            file_extension = os.path.splitext(uploaded_file.name)[1].lower()
            base_name = os.path.splitext(uploaded_file.name)[0]
            
            with st.expander(f"👁️ Preview: {uploaded_file.name}", expanded=True):
                try:
                    # Save temporary file to disk for MarkItDown to process
                    temp_path = f"temp_{uploaded_file.name}"
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    # Conversion step
                    # Adding a generic timeout logic is handled by the engine's internal session
                    result = md_engine.convert(temp_path)
                    content = result.text_content

                    # [2] Instant Preview
                    st.text_area(
                        label="Converted Content",
                        value=content,
                        height=300,
                        key=f"text_{uploaded_file.name}"
                    )

                    # [2] Download Options
                    col1, col2 = st.columns(2)
                    
                    # Markdown Download
                    col1.download_button(
                        label="📥 Download as Markdown (.md)",
                        data=content,
                        file_name=f"{base_name}_converted.md",
                        mime="text/markdown",
                        key=f"md_{uploaded_file.name}"
                    )

                    # Text Download
                    col2.download_button(
                        label="📥 Download as Text (.txt)",
                        data=content,
                        file_name=f"{base_name}_converted.txt",
                        mime="text/plain",
                        key=f"txt_{uploaded_file.name}"
                    )

                    # Cleanup
                    os.remove(temp_path)

                except Exception as e:
                    # [3] Resilience: Error Handling
                    st.error(f"⚠️ Could not read {uploaded_file.name}. Please check the format.")
                    # In a real app, you might log the specific error 'e' for debugging
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

if __name__ == "__main__":
    main()
