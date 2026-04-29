import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from bs4 import BeautifulSoup
import requests
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, folder_path)
    default_excel_name.set(os.path.basename(folder_path) + ".xlsx")
    default_html_name.set(os.path.splitext(default_excel_name.get())[0] + ".html")

def extract_comment(content):
    match = re.search(r'<!--(.*?)-->', content, re.DOTALL)
    if match:
        return match.group(1).strip()
    return ""

def extract_size(name_draft):
    match = re.search(r'(\d+x\d+)', name_draft)
    if match:
        return match.group(1)
    return ""

def process_content(content):
    return content.replace('/redir="', '/redir=%%c1;cpdir="')

def extract_name(name_draft):
    blocks = name_draft.split("/")
    last_block = blocks[-1]
    last_block = last_block.replace("crimtan_", "").replace("crimtan ", "").replace("Crimtan_", "").replace("Crimtan ", "")
    return last_block.strip()

def extract_links(html_path):
    try:
        with open(html_path, 'r') as file:
            content = file.read()
        
        pattern = r'document\.write\(\'<scr\'\+\'ipt src="([^"]+)"'
        matches = re.findall(pattern, content)
        
        links = []
        for match in matches:
            url_part = match.split("url=")
            if len(url_part) > 1:
                extracted_url = url_part[1]
                links.append(extracted_url)
            else:
                links.append(match)
        
        landing_pages = []
        for link in links:
            try:
                response = requests.head(link, allow_redirects=True, timeout=5)
                if response.status_code == 200:
                    landing_pages.append(response.url)
                else:
                    logging.error(f"Failed to follow redirect for {link}: Status Code {response.status_code}")
                    landing_pages.append(link)
            except Exception as e:
                logging.error(f"Failed to follow redirect for {link}: {e}")
                landing_pages.append(link)
        
        return landing_pages
    except Exception as e:
        logging.error(f"An error occurred while extracting links: {e}")
        return None
        
        
def create_excel_and_html():
    folder_path = folder_path_entry.get()
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_excel_name.get(), filetypes=[("Excel files", "*.xlsx")])
    html_path = filedialog.asksaveasfilename(defaultextension=".html", initialfile=default_html_name.get(), filetypes=[("HTML files", "*.html")])

    if folder_path and excel_path and html_path:
        workbook = Workbook()
        sheet = workbook.active
        # Modified header with new column
        sheet.append(["name", "Click URL", "Content", "size", "name_draft", 
                     "Original script without adform clickmacro", 
                     "script with dv360 macros",  # New column
                     "Landing Page"])

        with open(html_path, "w") as html_file:
            for foldername, subfolders, filenames in os.walk(folder_path):
                for filename in filenames:
                    if filename.endswith('.txt'):
                        txt_file = os.path.join(foldername, filename)
                        with open(txt_file, "r") as file:
                            content = file.read()
                        
                        comment = extract_comment(content)
                        name_draft = comment.split(',')[0] if ',' in comment else comment
                        size = extract_size(name_draft)
                        content = content.replace("<!--" + comment + "-->", "").strip()
                        processed_content = process_content(content)
                        name = extract_name(name_draft)
                        
                        # Create modified script
                        modified_script = (
                            content
                            .replace('gdpr=0', 'gdpr=${GDPR}')
                            .replace('gdpr_consent=', 'gdpr_consent=${GDPR_CONSENT_328}')
                            .replace('redir="', 'redir=${CLICK_URL}"')
                        )
                        
                        # Append row with all columns
                        sheet.append([
                            name, 
                            "", 
                            processed_content, 
                            size, 
                            name_draft, 
                            content,  # Original script
                            modified_script,  # New column
                            ""  # Landing page placeholder
                        ])
                        html_file.write(content + "\n" + "_" * 30 + "\n\n")

        # Extract landing pages from the generated HTML
        links = extract_links(html_path)
        
        if links:
            # Update landing pages in column H (8th column)
            link_index = 0
            for row in range(2, sheet.max_row + 1):
                if link_index < len(links):
                    sheet.cell(row=row, column=8).value = links[link_index]
                    link_index += 1

        workbook.save(excel_path)
        status_label.config(text="Excel and HTML files created successfully!")

# Create the main window
root = tk.Tk()
root.title("Text to Excel and HTML Converter with Link Extractor")
root.geometry("600x600")

default_excel_name = tk.StringVar()
default_html_name = tk.StringVar()

# Create and place widgets
folder_path_label = tk.Label(root, text="Select Folder:")
folder_path_label.pack()

folder_path_entry = tk.Entry(root)
folder_path_entry.pack()

browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.pack()

excel_name_label = tk.Label(root, text="Excel File Name:")
excel_name_label.pack()

excel_name_entry = tk.Entry(root, textvariable=default_excel_name)
excel_name_entry.pack()

html_name_label = tk.Label(root, text="HTML File Name:")
html_name_label.pack()

html_name_entry = tk.Entry(root, textvariable=default_html_name)
html_name_entry.pack()

create_button = tk.Button(root, text="Create Excel and HTML", command=create_excel_and_html)
create_button.pack()

status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()