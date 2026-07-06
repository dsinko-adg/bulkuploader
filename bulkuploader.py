import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
import re
import io
import os
import logging
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Set page configuration
st.set_page_config(
    page_title="Ad Script Converter Pro", 
    page_icon="📝", 
    layout="wide"
)

# Custom Styles
st.markdown("""
<style>
    .reportview-container {
        background: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)


# Core parsing and transformation helpers
def extract_comment(content):
    """Extracts first HTML comment block."""
    match = re.search(r'<!--(.*?)-->', content, re.DOTALL)
    return match.group(1).strip() if match else ""


def extract_size(name_draft):
    """Extracts resolution sizes from draft name (e.g., 300x250)."""
    match = re.search(r'(\d+x\d+)', name_draft)
    return match.group(1) if match else ""


def process_adform_content(content):
    """Applies Adform specific replacements."""
    return content.replace('/redir="', '/redir=%%c1;cpdir="')


def extract_name(name_draft):
    """Cleans up names by taking the final branch block and stripping Crimson/Crimtan decorators."""
    blocks = name_draft.split("/")
    last_block = blocks[-1]
    last_block = re.sub(r'(?i)crimtan[_ ]', '', last_block)
    return last_block.strip()


def extract_single_script_url(html_content):
    """Finds the source URL embedded in document.write expressions within the specific script block."""
    try:
        pattern = r'document\.write\(\'<scr\'\+\'ipt src="([^"]+)"'
        match = re.search(pattern, html_content)
        if match:
            url = match.group(1)
            # Break off redirect payloads if they are chained
            url_part = url.split("url=")
            return url_part[1] if len(url_part) > 1 else url
    except Exception as e:
        logging.error(f"Error extracting script URL: {e}")
    return None


def resolve_landing_page(url, timeout=5):
    """Resolves landing pages using HEAD with a smart GET fallback for anti-crawler servers."""
    if not url:
        return ""
    
    # Prepend schema if protocol-relative URL
    if url.startswith("//"):
        url = "https:" + url
        
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    try:
        # 1. Attempt lightweight HEAD request
        response = requests.head(url, allow_redirects=True, timeout=timeout, headers=headers)
        
        # 2. If server blocks HEAD requests (e.g., 405, 403), fallback to lightweight GET
        if response.status_code in [403, 404, 405]:
            response = requests.get(url, allow_redirects=True, timeout=timeout, headers=headers, stream=True)
            
        if response.status_code == 200:
            return response.url
        else:
            return url
    except Exception as e:
        logging.error(f"Network redirection error for {url}: {e}")
        return url


def parse_and_process_file_content(raw_bytes, filename, parent_folder_name, macro_gdpr, macro_consent, macro_redir):
    """Decodes, parses, and creates the structured metadata mapping for each script."""
    # Robust character encoding detection fallback
    try:
        content = raw_bytes.decode('utf-8')
    except UnicodeDecodeError:
        try:
            content = raw_bytes.decode('latin1')
        except UnicodeDecodeError:
            content = raw_bytes.decode('cp1252', errors='replace')
    
    comment = extract_comment(content)
    name_draft = comment.split(',')[0] if ',' in comment else comment
    
    # If no draft name in comment, fallback to filename (without extension)
    if not name_draft:
        name_draft = os.path.splitext(filename)[0]
        
    # Replicate your desktop renamer logic: Prepend the immediate parent folder name
    if parent_folder_name:
        name_draft = f"{parent_folder_name}_{name_draft}"
        
    size = extract_size(name_draft)
    
    # Strip out processing comments to find original clean code
    clean_original = content
    if comment:
        clean_original = content.replace(f"<!--{comment}-->", "").strip()
    
    # Run processed transforms
    processed_adform = process_adform_content(clean_original)
    name = extract_name(name_draft)
    
    # Modify DV360 tracking variables based on config
    processed_dv360 = (
        clean_original
        .replace('gdpr=0', macro_gdpr)
        .replace('gdpr_consent=', macro_consent)
        .replace('redir="', macro_redir)
    )
    
    # Extract script-level redirection paths
    extracted_url = extract_single_script_url(clean_original)
    
    return {
        "name": name,
        "size": size,
        "name_draft": name_draft,
        "clean_original": clean_original,
        "processed_adform": processed_adform,
        "processed_dv360": processed_dv360,
        "extracted_url": extracted_url,
        "landing_page": ""  # Resolved next
    }


def generate_styled_html_preview(records):
    """Generates a premium, clean, and highly usable preview HTML file with copying built-in."""
    cards_html = ""
    for idx, r in enumerate(records):
        cards_html += f"""
        <div class="card">
            <div class="card-header">
                <div>
                    <span class="badge">{r['size'] if r['size'] else 'Unknown Size'}</span>
                    <strong style="font-size: 1.15rem; color: #1e293b; margin-left: 8px;">{r['name']}</strong>
                </div>
                <div style="font-size: 0.85rem; color: #64748b;">
                    Draft: {r['name_draft']}
                </div>
            </div>
            <div class="card-body">
                <div class="section-title">Original Cleaned Script</div>
                <div class="textarea-container">
                    <textarea id="orig-{idx}" readonly>{r['clean_original']}</textarea>
                    <button class="btn-copy" onclick="copyText('orig-{idx}')">Copy Clean</button>
                </div>
                
                <div class="section-title">DV360 Modified Script</div>
                <div class="textarea-container">
                    <textarea id="dv360-{idx}" readonly>{r['processed_dv360']}</textarea>
                    <button class="btn-copy" onclick="copyText('dv360-{idx}')">Copy DV360</button>
                </div>
                
                {f'<div class="landing-page"><strong>Destination Landing Page:</strong> <a href="{r["landing_page"]}" target="_blank">{r["landing_page"]}</a></div>' if r['landing_page'] else ''}
            </div>
        </div>
        """

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ad Scripts Live Preview Dashboard</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #f1f5f9;
            color: #334155;
            margin: 0;
            padding: 40px 20px;
        }}
        .container {{
            max-width: 900px;
            margin: 0 auto;
        }}
        header {{
            margin-bottom: 30px;
            text-align: center;
        }}
        h1 {{
            color: #0f172a;
            font-size: 2rem;
            margin-bottom: 8px;
        }}
        .card {{
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
            margin-bottom: 24px;
            border: 1px solid #e2e8f0;
            overflow: hidden;
        }}
        .card-header {{
            background-color: #f8fafc;
            padding: 16px 20px;
            border-bottom: 1px solid #e2e8f0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }}
        .badge {{
            background-color: #3b82f6;
            color: white;
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: bold;
            text-transform: uppercase;
        }}
        .card-body {{
            padding: 20px;
        }}
        .section-title {{
            font-size: 0.85rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #475569;
            margin-top: 14px;
            margin-bottom: 6px;
        }}
        .textarea-container {{
            position: relative;
            margin-bottom: 16px;
        }}
        textarea {{
            width: 100%;
            height: 100px;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            font-size: 0.85rem;
            padding: 10px;
            box-sizing: border-box;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            background-color: #f8fafc;
            color: #334155;
            resize: vertical;
        }}
        .btn-copy {{
            position: absolute;
            top: 8px;
            right: 8px;
            background: #0f172a;
            color: #ffffff;
            border: none;
            padding: 4px 10px;
            font-size: 0.75rem;
            border-radius: 4px;
            cursor: pointer;
            transition: all 0.2s;
            font-weight: 500;
        }}
        .btn-copy:hover {{
            background: #1e293b;
        }}
        .btn-copy.success {{
            background: #10b981;
        }}
        .landing-page {{
            margin-top: 15px;
            background-color: #eff6ff;
            padding: 10px 14px;
            border-radius: 6px;
            border: 1px solid #bfdbfe;
            font-size: 0.9rem;
            word-break: break-all;
        }}
        .landing-page a {{
            color: #2563eb;
            text-decoration: none;
            font-weight: 500;
        }}
        .landing-page a:hover {{
            text-decoration: underline;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Ad Scripts Preview Directory</h1>
            <p style="color: #64748b;">Generated dynamically via Ad Script Converter Pro</p>
        </header>
        {cards_html}
    </div>

    <script>
        function copyText(elementId) {{
            const textarea = document.getElementById(elementId);
            textarea.select();
            textarea.setSelectionRange(0, 99999); // For mobile devices
            
            // Safe copy utility fallback for iframe contexts
            try {{
                document.execCommand('copy');
                const btn = textarea.nextElementSibling;
                const baseText = btn.innerText;
                btn.innerText = "Copied!";
                btn.classList.add('success');
                setTimeout(() => {{
                    btn.innerText = baseText;
                    btn.classList.remove('success');
                }}, 1500);
            }} catch (err) {{
                console.error('Copy failed', err);
            }}
        }}
    </script>
</body>
</html>
"""
    return html_template


def update_outputs_from_records(records):
    """Regenerates files in session state dynamically when records are altered by Find/Replace operations."""
    # Write outputs to high-fidelity styled Excel structure
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Ad Scripts Data"
    
    # Show grid lines
    sheet.views.sheetView[0].showGridLines = True
    
    headers = [
        "Creative Name", "Click URL", "Content (Adform Tag)", "Creative Size", 
        "Campaign Manager Placement/Draft Name", "Clean Original Script", 
        "DV360 Modified Script Tag", "Destination Landing Page URL"
    ]
    sheet.append(headers)
    
    # Styled premium elements for columns and headings
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    border_thin = Side(border_style="thin", color="D9D9D9")
    cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    
    for col_num, header_title in enumerate(headers, 1):
        cell = sheet.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
    
    for r in records:
        sheet.append([
            r["name"],
            "",  # Click URL placeholder
            r["processed_adform"],
            r["size"],
            r["name_draft"],
            r["clean_original"],
            r["processed_dv360"],
            r["landing_page"]
        ])
    
    # Format spreadsheet layouts and adjust column widths
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=8):
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.border = cell_border
            # Align large textual script code boxes gracefully
            if cell.column in [3, 6, 7]:
                cell.alignment = Alignment(vertical="top", wrap_text=False)
            else:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                
    # Determine best-fit column sizing
    for col in sheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        # Cap script block widths to prevent horizontal overstretching
        if col[0].column in [3, 6, 7]:
            sheet.column_dimensions[col_letter].width = 35
        else:
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    # Excel save to binary buffer
    excel_buffer = io.BytesIO()
    workbook.save(excel_buffer)
    excel_buffer.seek(0)
    
    # Generate styled raw preview outputs
    html_output = generate_styled_html_preview(records)
    html_buffer = io.BytesIO(html_output.encode('utf-8'))
    
    # Update Session states
    st.session_state.excel_bytes = excel_buffer.getvalue()
    st.session_state.html_bytes = html_buffer.getvalue()
    
    # Refresh live previews
    st.session_state.df_preview = [
        {
            "Name": r["name"],
            "Size": r["size"],
            "Landing Page": r["landing_page"],
            "Adform Script Preview": r["processed_adform"][:120] + "..." if len(r["processed_adform"]) > 120 else r["processed_adform"],
            "DV360 Script Preview": r["processed_dv360"][:120] + "..." if len(r["processed_dv360"]) > 120 else r["processed_dv360"]
        }
        for r in records
    ]


def main():
    st.title("📝 Ad Script Converter Pro")
    st.write("Convert agency-specific raw `.txt` script tags (either uploaded directly or nested in `.zip` folders) into clean, macro-enabled configurations ready for upload.")

    # Highly optimized default configurations
    resolve_redirects = True
    timeout_limit = 5
    concurrent_threads = 8
    macro_gdpr = "gdpr=${GDPR}"
    macro_consent = "gdpr_consent=${GDPR_CONSENT_328}"
    macro_redir = "redir=${CLICK_URL}\""

    # State setups
    if 'excel_bytes' not in st.session_state:
        st.session_state.excel_bytes = None
    if 'html_bytes' not in st.session_state:
        st.session_state.html_bytes = None
    if 'df_preview' not in st.session_state:
        st.session_state.df_preview = None
    if 'records' not in st.session_state:
        st.session_state.records = None

    uploaded_files = st.file_uploader(
        "Upload Raw Ad Script Files (.txt or .zip archives)", 
        type=['txt', 'zip'], 
        accept_multiple_files=True
    )

    if uploaded_files:
        st.info(f"📂 {len(uploaded_files)} item(s) selected.")
        
        if st.button("Process Scripts", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            records = []
            
            # Step 1: Parse and unpack files (including deep directory structures inside ZIP files)
            for idx, uploaded_file in enumerate(uploaded_files):
                file_name_lower = uploaded_file.name.lower()
                
                if file_name_lower.endswith('.zip'):
                    status_text.text(f"📦 Extracting directory structure from: {uploaded_file.name}...")
                    try:
                        zip_data = uploaded_file.read()
                        with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
                            for member in z.infolist():
                                # Skip directory entries and non-txt files
                                if member.filename.endswith('/') or member.is_dir():
                                    continue
                                if not member.filename.lower().endswith('.txt'):
                                    continue
                                
                                # Retrieve folder hierarchy details
                                parent_folder_name = os.path.basename(os.path.dirname(member.filename))
                                filename_only = os.path.basename(member.filename)
                                
                                with z.open(member) as f:
                                    raw_bytes = f.read()
                                    
                                record = parse_and_process_file_content(
                                    raw_bytes,
                                    filename_only,
                                    parent_folder_name,
                                    macro_gdpr,
                                    macro_consent,
                                    macro_redir
                                )
                                records.append(record)
                    except Exception as e:
                        st.error(f"Error unpacking ZIP archive '{uploaded_file.name}': {str(e)}")
                
                elif file_name_lower.endswith('.txt'):
                    status_text.text(f"📄 Reading single script file: {uploaded_file.name}...")
                    raw_bytes = uploaded_file.read()
                    
                    record = parse_and_process_file_content(
                        raw_bytes,
                        uploaded_file.name,
                        "",  # No parent folder from simple direct file upload
                        macro_gdpr,
                        macro_consent,
                        macro_redir
                    )
                    records.append(record)

            if not records:
                st.warning("⚠️ No valid `.txt` script tags were discovered in the uploaded content.")
                return

            # Step 2: Resolve redirect links concurrently to avoid sequential blocking delays
            if resolve_redirects:
                status_text.text("🔗 Running parallel checks on redirection loops...")
                
                # Create maps to process only unique URLs (reduces network redundancy)
                unique_urls = list({r["extracted_url"] for r in records if r["extracted_url"]})
                resolved_url_map = {}
                
                if unique_urls:
                    with ThreadPoolExecutor(max_workers=concurrent_threads) as executor:
                        # Map future tasks to original urls
                        future_to_url = {
                            executor.submit(resolve_landing_page, url, timeout_limit): url 
                            for url in unique_urls
                        }
                        
                        completed_count = 0
                        total_futures = len(future_to_url)
                        
                        for future in as_completed(future_to_url):
                            orig_url = future_to_url[future]
                            try:
                                resolved_url = future.result()
                                resolved_url_map[orig_url] = resolved_url
                            except Exception as e:
                                resolved_url_map[orig_url] = orig_url
                                logging.error(f"Error handling future result: {e}")
                            
                            completed_count += 1
                            progress_val = int((completed_count / total_futures) * 100)
                            progress_bar.progress(progress_val)
                            status_text.text(f"Resolved {completed_count}/{total_futures} redirect URLs...")
                
                # Map resolved configurations back to records
                for r in records:
                    url_to_lookup = r["extracted_url"]
                    if url_to_lookup in resolved_url_map:
                        r["landing_page"] = resolved_url_map[url_to_lookup]
            
            # Step 3: Populate primary structures & trigger update loops
            status_text.text("📊 Formatting premium spreadsheet outputs...")
            st.session_state.records = records
            update_outputs_from_records(records)
            
            progress_bar.progress(100)
            status_text.empty()
            st.success(f"Successfully processed {len(records)} script configs!")

    # Step 4: Display Find and Replace, on-screen previews & export downloads
    if st.session_state.records is not None:
        
        # New Bulk Find and Replace Utility Block
        st.markdown("### 🛠️ Find & Replace Naming Utility")
        st.info("💡 Adjust and align naming patterns dynamically (similar to Ctrl+H in Excel) before generating final outputs.")
        
        col_find, col_replace, col_opt = st.columns([3, 3, 4])
        with col_find:
            find_val = st.text_input("Find string", placeholder="e.g., OldBrandName")
        with col_replace:
            replace_val = st.text_input("Replace with", placeholder="e.g., NewBrandName")
        with col_opt:
            st.write("Target Fields:")
            sub_col1, sub_col2 = st.columns(2)
            with sub_col1:
                replace_in_name = st.checkbox("Apply to Creative Name", value=True)
            with sub_col2:
                replace_in_draft = st.checkbox("Apply to Placement/Draft Name", value=False)
                
        if st.button("Apply Replacements", type="secondary"):
            if find_val:
                match_count = 0
                for r in st.session_state.records:
                    target_updated = False
                    if replace_in_name and find_val in r["name"]:
                        r["name"] = r["name"].replace(find_val, replace_val)
                        target_updated = True
                    if replace_in_draft and find_val in r["name_draft"]:
                        r["name_draft"] = r["name_draft"].replace(find_val, replace_val)
                        target_updated = True
                    if target_updated:
                        match_count += 1
                        
                if match_count > 0:
                    update_outputs_from_records(st.session_state.records)
                    st.success(f"Successfully replaced '{find_val}' with '{replace_val}' in {match_count} item(s)!")
                    st.rerun()
                else:
                    st.warning(f"No matches found for '{find_val}' within the targeted fields.")
            else:
                st.error("Please enter a string to find.")

        st.markdown("---")
        st.markdown("### 📊 Live Conversion Preview")
        st.dataframe(st.session_state.df_preview, use_container_width=True)

        st.markdown("### 📥 Export Final Deliverables")
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download Excel Tracker Spreadsheet",
                data=st.session_state.excel_bytes,
                file_name="processed_ad_tracker.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            st.download_button(
                label="🌐 Download Live HTML Preview Directory",
                data=st.session_state.html_bytes,
                file_name="ad_scripts_preview_dashboard.html",
                mime="text/html"
            )


if __name__ == "__main__":
    main()
