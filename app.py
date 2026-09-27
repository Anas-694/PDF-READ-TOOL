import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import time
import os
import io
import re
from datetime import datetime

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

def add_formatted_paragraph(doc, line, style=None):
    p = doc.add_paragraph(style=style)
    # Split by bold tags
    chunks = re.split(r'(\*\*.*?\*\*)', line)
    for chunk in chunks:
        if chunk.startswith('**') and chunk.endswith('**'):
            run = p.add_run(chunk[2:-2])
            run.bold = True
        else:
            p.add_run(chunk)
    return p

def add_markdown_to_doc(doc, text):
    lines = text.split('\n')
    in_table = False
    table_data = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Detect table row
        if line.startswith('|') and line.endswith('|'):
            if re.match(r'^\|[\s\-\|]+\|$', line):
                continue
            row = [cell.strip() for cell in line.split('|')[1:-1]]
            table_data.append(row)
            in_table = True
        else:
            if in_table:
                if table_data:
                    # Attempt to use a built-in nice style, fallback to Table Grid if not found
                    table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                    try:
                        table.style = 'Light Shading Accent 1'
                    except:
                        table.style = 'Table Grid'
                        
                    for i, row_data in enumerate(table_data):
                        row_cells = table.rows[i].cells
                        for j, cell_text in enumerate(row_data):
                            if j < len(row_cells):
                                cell_text_clean = cell_text.replace('**', '') # Tables don't need bold markers inside
                                run = row_cells[j].paragraphs[0].add_run(cell_text_clean)
                                
                                # Color coding logic for financial data
                                if '%' in cell_text_clean or 'YoY' in cell_text_clean or 'Growth' in cell_text_clean:
                                    if '-' in cell_text_clean or '(' in cell_text_clean or 'Decline' in cell_text_clean:
                                        run.font.color.rgb = RGBColor(204, 0, 0) # Red
                                        run.bold = True
                                    elif '+' in cell_text_clean or 'Growth' in cell_text_clean or 'Increase' in cell_text_clean:
                                        run.font.color.rgb = RGBColor(0, 153, 51) # Green
                                        run.bold = True
                table_data = []
                in_table = False
                
            if line.startswith('# '):
                doc.add_heading(line[2:].strip(), level=1)
            elif line.startswith('## '):
                doc.add_heading(line[3:].strip(), level=2)
            elif line.startswith('### '):
                doc.add_heading(line[4:].strip(), level=3)
            elif line.startswith('- ') or line.startswith('* '):
                # Handle bold in list items
                clean_line = line[2:].strip()
                add_formatted_paragraph(doc, clean_line, style='List Bullet')
            else:
                add_formatted_paragraph(doc, line)
                
    if in_table and table_data:
        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
        try:
            table.style = 'Light Shading Accent 1'
        except:
            table.style = 'Table Grid'
            
        for i, row_data in enumerate(table_data):
            row_cells = table.rows[i].cells
            for j, cell_text in enumerate(row_data):
                if j < len(row_cells):
                    cell_text_clean = cell_text.replace('**', '')
                    run = row_cells[j].paragraphs[0].add_run(cell_text_clean)
                    
                    if '%' in cell_text_clean or 'YoY' in cell_text_clean or 'Growth' in cell_text_clean:
                        if '-' in cell_text_clean or '(' in cell_text_clean or 'Decline' in cell_text_clean:
                            run.font.color.rgb = RGBColor(204, 0, 0)
                            run.bold = True
                        elif '+' in cell_text_clean or 'Growth' in cell_text_clean or 'Increase' in cell_text_clean:
                            run.font.color.rgb = RGBColor(0, 153, 51)
                            run.bold = True

def analyze_page(model, text, page_num):
    prompt = f"""You are an Expert Financial Analyst specializing in the Pakistan Stock Exchange (PSX).

Your objective is to read this single page of a financial report and extract ONLY new, critical information. 
DO NOT output boilerplate company introductions if they are already obvious. DO NOT invent information.
If the page contains only empty space, signatures, or irrelevant filler text, reply with exactly "EMPTY_PAGE".

WHAT TO EXTRACT (If present on this page):
1. New Developments: Any new projects, partnerships, or strategic changes.
2. Financial Numbers: Extract revenue, profit, margins, EPS, etc. YOU MUST FORMAT THESE NUMBERS AS A MARKDOWN TABLE. Add a column for YoY Growth % if possible. Use + for positive growth and - for negative growth.
3. Management Commentary: Real outlook, future plans, risks, or challenges.

STRICT RULES:
1. Language: Roman Urdu (easy words). Use English for financial terms (Revenue, EPS).
2. No Repetition: Do not say "Here is the summary". Just give the facts.
3. Tables: Whenever you see financial figures (e.g. 2024 vs 2025), put them in a Markdown Table format like this:
| Indicator | 2024 | 2025 | Growth % |
|---|---|---|---|
| Revenue | 100 | 120 | +20% |
4. Headings: Use ## for topics. Do not use # (H1).
5. Bold Text: Use **bold** for key metrics, names, and important highlights.

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
            
            pdf_bytes = uploaded_file.read()
            pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            doc = Document()
            
            # --- PROFESSIONAL COVER PAGE ---
            doc.add_heading("Financial Analysis Report", 0)
            doc.add_paragraph(f"Generated on: {datetime.now().strftime('%d %B %Y, %H:%M')}")
            doc.add_paragraph(f"Source File: {uploaded_file.name}")
            doc.add_page_break()
            # -------------------------------
            
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
                        # Professional Subtle Page Divider
                        p = doc.add_paragraph(f"Source: Page {page_num + 1}")
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        p.runs[0].font.size = Pt(9)
                        p.runs[0].font.color.rgb = RGBColor(128, 128, 128)
                        
                        add_markdown_to_doc(doc, summary)
                        doc.add_paragraph('\n')
                
                pages_done += 1
                progress_bar.progress(pages_done / total_pages_to_process)
                
                if page_num < end_page_to_process - 1:
                    time.sleep(5)
            
            status_text.text("Analysis Complete! Generating File...")
            
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            st.success("✅ Zabardast! File tayar hai.")
            
            file_name = f"{os.path.splitext(uploaded_file.name)[0]} - Summary.docx"
            st.download_button(
                label="📥 Download Professional Report (.docx)",
                data=doc_io,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Koi masla aagaya: {str(e)}")
