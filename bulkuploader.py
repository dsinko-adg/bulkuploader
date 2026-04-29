import streamlit as st
from openpyxl import Workbook
import requests
import re
import io
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Set page configuration
st.set_page_config(page_title="Ad Script Converter", page_icon="📝", layout="centered")

def extract_comment(content):
    match = re.search(r'<!--(.*?)-->', content, re.DOTALL)
    return match.group(1).strip() if match else ""

def extract_size(name_draft):
    match = re.search(r'(\d+x\d+)', name_draft)
    return match.group(1) if match else ""

def process_content(content):
    return content.replace('/redir="', '/redir=%%c1;cpdir="')

def extract_name(name_draft):
    blocks = name_draft.split("/")
    last_block = blocks[-1]
    # Use regex for case-insensitive replacement
    last_block = re.sub(r'(?i)crimtan[_ ]', '', last_block)
    return last_block.strip()

def extract_links(html_content):
    try:
        pattern = r'document\.write\(\'<scr\'\+\'ipt src="([^"]+)"'
        matches = re.findall(pattern, html_content)
        
        landing_pages = []
        for match in matches:
            url_part = match.split("url=")
            link_to_check = url_part[1] if len(url_part) > 1 else match
            
            try:
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.head(link_to_check, allow_redirects=True, timeout=5, headers=headers)
                if response.status_code == 200:
                    landing_pages.append(response.url)
                else:
                    logging.error(f"Redirect failed for {link_to_check}: Status {response.status_code}")
                    landing_pages.append(link_to_check)
            except Exception as e:
                logging.error(f"Network error following redirect for {link_to_check}: {e}")
                landing_pages.append(link_to_check)
        
        return landing_pages
    except Exception as e:
        logging.error(f"An error occurred while extracting links: {e}")
        return []

def main():
    st.title("📝 Ad Script Converter")
    st.write("Upload your `.txt` ad scripts to generate an Excel tracker and an HTML preview.")

    # Initialize session state to store generated files
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'html_data' not in st.session_state:
        st.session_state.html_data = None

    uploaded_files = st.file_uploader("Select .txt files", type=['txt'], accept_multiple_files=True)

    if uploaded_files:
        st.write(f"**{len(uploaded_files)} files selected.**")
        
        if st.button("Process Files", type="primary"):
            with st.spinner("Processing scripts and extracting landing pages..."):
                try:
                    # Setup Excel workbook
                    workbook = Workbook()
                    sheet = workbook.active
                    headers = [
                        "name", "Click URL", "Content", "size", "name_draft", 
                        "Original script without adform clickmacro", 
                        "script with dv360 macros", "Landing Page"
                    ]
                    sheet.append(headers)

                    html_output = ""

                    for uploaded_file in uploaded_files:
                        # Streamlit uploaded files are bytes; decode to string
                        content = uploaded_file.read().decode('utf-8')
                        
                        comment = extract_comment(content)
                        name_draft = comment.split(',')[0] if ',' in comment else comment
                        size = extract_size(name_draft)
                        
                        if comment:
                            content = content.replace(f"<!--{comment}-->", "").strip()
                        
                        processed_content = process_content(content)
                        name = extract_name(name_draft)
                        
                        modified_script = (
                            content
                            .replace('gdpr=0', 'gdpr=${GDPR}')
                            .replace('gdpr_consent=', 'gdpr_consent=${GDPR_CONSENT_328}')
                            .replace('redir="', 'redir=${CLICK_URL}"')
                        )
                        
                        sheet.append([
                            name, "", processed_content, size, name_draft, 
                            content, modified_script, "" 
                        ])
                        
                        html_output += content + "\n" + "_" * 30 + "\n\n"

                    # Extract landing pages
                    links = extract_links(html_output)
                    if links:
                        for row_idx, link in enumerate(links, start=2):
                            if row_idx <= sheet.max_row:
                                sheet.cell(row=row_idx, column=8).value = link

                    # Save Excel to memory buffer
                    excel_buffer = io.BytesIO()
                    workbook.save(excel_buffer)
                    excel_buffer.seek(0)
                    
                    # Save HTML to memory buffer
                    html_buffer = io.BytesIO(html_output.encode('utf-8'))
                    
                    # Store in session state so downloads persist
                    st.session_state.excel_data = excel_buffer.getvalue()
                    st.session_state.html_data = html_buffer.getvalue()

                    st.success("Files processed successfully!")

                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")

    # Show download buttons if data exists in session state
    if st.session_state.excel_data and st.session_state.html_data:
        st.markdown("### Download Results")
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download Excel File",
                data=st.session_state.excel_data,
                file_name="processed_ad_scripts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            st.download_button(
                label="🌐 Download HTML File",
                data=st.session_state.html_data,
                file_name="ad_scripts_preview.html",
                mime="text/html"
            )

if __name__ == "__main__":
    main()