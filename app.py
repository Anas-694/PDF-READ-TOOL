import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
from docx import Document
import time
import os
import io

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
    for line in lines:
        if line.startswith('# '):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith('## '):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith('### '):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith('- ') or line.startswith('* '):
            doc.add_paragraph(line[2:].strip(), style='List Bullet')
        elif line.strip() != "":
            clean_line = line.replace('**', '') 
            doc.add_paragraph(clean_line.strip())

def analyze_page(model, text, page_num):
    prompt = f"""You are an Expert Investor, a First-Principles Business Operator, and a Financial Analyst specializing in the Pakistan Stock Exchange (PSX). You have built, scaled, and exited successful companies, so you know how to look past the marketing fluff and identify the real core of a business.

Your objective is to read the provided text from a company's financial statement / report carefully and extract the absolute most critical information that an investor MUST know.

WHAT TO EXTRACT:
1. Company Introduction: What the company does, its core business, sector, and any background context provided.
2. New Developments: Any new advancements, expansions, projects, partnerships, product launches, or strategic changes mentioned by the company.
3. Financial Results: Revenue, profit/loss, margins, EPS, growth or decline compared to previous period, and any other key financial figures reported.
4. Management Commentary: Everything the company wants to communicate to its investors — outlook, future plans, challenges, risks, and management's own remarks or explanations.
5. Any other material fact an investor should know before making a decision.

STRICT RULES FOR YOUR RESPONSE:
1. Language: Your entire output MUST be in Roman Urdu. Use simple, easy-to-understand words (asan alfaz) so a common retail investor can easily grasp complex financial concepts. Do NOT write sentences in English. You may only use English for unavoidable financial terms (like 'EBITDA', 'IPO', 'Revenue', 'EPS', etc.), but explain their context in Roman Urdu.
2. No Fluff & High Precision: Be extremely precise, authentic, and to the point. Do not add filler words or generic advice. Every single sentence must carry maximum weight.
3. Structure and Formatting: Provide neat and clean formatting. Use H1 for main section, H2 for sub-topics. Use bullet points instead of long paragraphs.
4. Focus: Extract the core business model, new developments, financial performance, risks, and real truths — everything the company is communicating to its investors.

PAGE TEXT (Page {page_num}):
{text}"""
    try:
        response = model.generate_content(prompt)
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
                    if summary:
                        doc.add_heading(f'Page {page_num + 1} Summary', level=1)
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
