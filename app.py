import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
from docx import Document
import time
import os
import io
import re

# Setup the page configuration
st.set_page_config(page_title="PDF Read Tool", page_icon="📄", layout="centered")

st.title("📄 PDF Read Tool - Smart Analyzer")
st.write("Apni PDF yahan upload karein aur AI uski summary nikal dega.")

# Sidebar for configuration
with st.sidebar:
    st.header("Settings")
    api_key = st.text_input("Google Gemini API Key", type="password")
    start_page = st.number_input("Start Page", min_value=1, value=1)
    end_page = st.number_input("End Page", min_value=1, value=1000)

def add_markdown_to_doc(doc, text):
    lines = text.split('\n')
    in_table = False
    table_data = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.startswith('|') and line.endswith('|'):
            if re.match(r'^\|[\s\-\|]+\|$', line):
                continue
            row = [cell.strip() for cell in line.split('|')[1:-1]]
            table_data.append(row)
            in_table = True
        else:
            if in_table:
                if table_data:
                    table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                    table.style = 'Table Grid'
                    for i, row_data in enumerate(table_data):
                        row_cells = table.rows[i].cells
                        for j, cell_text in enumerate(row_data):
                            if j < len(row_cells):
                                row_cells[j].text = cell_text
                table_data = []
                in_table = False
                
            if line.startswith('# '):
                doc.add_heading(line[2:].strip(), level=1)
            elif line.startswith('## '):
                doc.add_heading(line[3:].strip(), level=2)
            elif line.startswith('### '):
                doc.add_heading(line[4:].strip(), level=3)
            elif line.startswith('- ') or line.startswith('* '):
                doc.add_paragraph(line[2:].strip(), style='List Bullet')
            else:
                clean_line = line.replace('**', '') 
                doc.add_paragraph(clean_line.strip())
                
    if in_table and table_data:
        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
        table.style = 'Table Grid'
        for i, row_data in enumerate(table_data):
            row_cells = table.rows[i].cells
            for j, cell_text in enumerate(row_data):
                if j < len(row_cells):
                    row_cells[j].text = cell_text

def analyze_page(model, text, page_num):
    prompt = f"""You are an Expert Financial Analyst specializing in the Pakistan Stock Exchange (PSX).

Your objective is to read this single page of a financial report and extract ONLY new, critical information. 
DO NOT output boilerplate company introductions if they are already obvious. DO NOT invent information.
If the page contains only empty space, signatures, or irrelevant filler text, reply with exactly "EMPTY_PAGE".

WHAT TO EXTRACT (If present on this page):
1. New Developments: Any new projects, partnerships, or strategic changes.
2. Financial Numbers: Extract revenue, profit, margins, EPS, etc. YOU MUST FORMAT THESE NUMBERS AS A MARKDOWN TABLE.
3. Management Commentary: Real outlook, future plans, risks, or challenges.

STRICT RULES:
1. Language: Roman Urdu (easy words). Use English for financial terms (Revenue, EPS).
2. No Repetition: Do not say "Here is the summary". Just give the facts.
3. Tables: Whenever you see financial figures (e.g. 2024 vs 2025), put them in a Markdown Table format like this:
| Indicator | 2024 | 2025 |
|---|---|---|
| Revenue | 100 | 120 |
4. Headings: Use ## for topics. Do not use # (H1).

PAGE TEXT:
{text}"""
    try:
        response = model.generate_content(prompt)
        if "EMPTY_PAGE" in response.text:
            return ""
        return response.text
    except Exception as e:
        return f"Error on page {page_num}: {str(e)}"

uploaded_file = st.file_uploader("PDF File upload karein", type=["pdf"])

if st.button("Start Analysis"):
    if not uploaded_file:
        st.error("Bhai pehlay PDF file toh upload karein!")
    elif not api_key:
        st.error("API Key zaroori hai!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-flash-lite-latest")
            
            # Read PDF from uploaded file
            pdf_bytes = uploaded_file.read()
            pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            doc = Document()
            total_pages = len(pdf_document)
            end_page_to_process = min(end_page, total_pages)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total_pages_to_process = end_page_to_process - (start_page - 1)
            pages_done = 0
            
            for page_num in range(start_page - 1, end_page_to_process):
                status_text.text(f"Page {page_num + 1} analyze ho rahi hai... Please wait.")
                
                page = pdf_document.load_page(page_num)
                text = page.get_text("text")
                
                if text.strip():
                    summary = analyze_page(model, text, page_num + 1)
                    if summary and summary.strip():
                        doc.add_paragraph(f"--- Page {page_num + 1} ---")
                        add_markdown_to_doc(doc, summary)
                        doc.add_paragraph('\n')
                
                pages_done += 1
                progress_bar.progress(pages_done / total_pages_to_process)
                
                # Sleep to prevent API rate limits (unless it's the last page)
                if page_num < end_page_to_process - 1:
                    time.sleep(5)
            
            status_text.text("Analysis Complete! Generating File...")
            
            # Save the docx to in-memory bytes
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            st.success("✅ Zabardast! File tayar hai.")
            
            file_name = f"{os.path.splitext(uploaded_file.name)[0]} - Summary.docx"
            st.download_button(
                label="📥 Download Summary (.docx)",
                data=doc_io,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Koi masla aagaya: {str(e)}")
