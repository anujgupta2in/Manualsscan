"""
Streamlit Manual/Drawing Scanner App
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import os
import io
import time
import logging
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from pypdf import PdfReader
try:
    from docx import Document
except ImportError:
    Document = None

# Import extraction utilities
from extraction_utils import (
    normalize_text,
    identify_manual_name,
    classify_doc_type,
    METADATA_PATTERNS,
    extract_with_regex
)

# --- Configuration & State ---
st.set_page_config(page_title="Manual & Drawing Scanner", layout="wide")

if "results" not in st.session_state:
    st.session_state.results = []
if "scanning" not in st.session_state:
    st.session_state.scanning = False
if "stop_requested" not in st.session_state:
    st.session_state.stop_requested = False

# --- UI Helper ---
def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder_path = filedialog.askdirectory(master=root)
    root.destroy()
    return folder_path

# --- Extraction Logic ---
def extract_pdf_content(file_path):
    """Extract text and metadata from the first two pages of a PDF."""
    try:
        reader = PdfReader(file_path)
        pages = []
        for i in range(min(2, len(reader.pages))):
            text = reader.pages[i].extract_text() or ""
            pages.append(text)
        return "\n".join(pages), reader.metadata or {}, "Success"
    except Exception as e:
        return "", {}, f"Error: {str(e)}"

def extract_docx_content(file_path):
    """Extract text from the first two pages of a DOCX (approx 50 paras)."""
    if Document is None:
        return "", "Skipped: python-docx not installed"
    try:
        doc = Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs[:50]])
        return text, "Success"
    except Exception as e:
        return "", f"Error: {str(e)}"

# --- Main App UI ---
st.title("🚢 Manual & Drawing Scanner")
st.markdown("Scan folders for ship manuals, drawings, and certificates. Extracts titles and classifies document types.")

with st.sidebar:
    st.header("Scan Settings")
    input_folder = st.text_input("Folder Path", value=st.session_state.get("last_folder", ""))
    
    if st.button("Browse Folder (Windows)"):
        selected = select_folder()
        if selected:
            input_folder = selected
            st.session_state.last_folder = selected
            st.rerun()

    include_subfolders = st.checkbox("Include subfolders", value=True)
    scan_docx = st.checkbox("Scan DOCX", value=True)
    enable_debug = st.checkbox("Enable debug logs", value=False)
    
    col1, col2 = st.columns(2)
    with col1:
        start_btn = st.button("Start Scan", use_container_width=True, type="primary", disabled=st.session_state.scanning)
    with col2:
        stop_btn = st.button("Stop Scan", use_container_width=True, disabled=not st.session_state.scanning)

if stop_btn:
    st.session_state.stop_requested = True
    st.warning("Stop requested. Finishing current file...")

if start_btn:
    if not input_folder or not os.path.exists(input_folder):
        st.error("Please provide a valid folder path.")
    else:
        st.session_state.scanning = True
        st.session_state.stop_requested = False
        st.session_state.results = []
        
        target_path = Path(input_folder)
        if include_subfolders:
            all_files = list(target_path.rglob("*"))
        else:
            all_files = list(target_path.glob("*"))
            
        files_to_scan = [f for f in all_files if f.is_file()]
        total_files = len(files_to_scan)
        
        if total_files == 0:
            st.warning("No files found in the selected folder.")
            st.session_state.scanning = False
        else:
            # LIVE COUNTERS
            stats_cols = st.columns(6)
            total_c = stats_cols[0].metric("Total", total_files)
            processed_c = stats_cols[1].empty()
            success_c = stats_cols[2].empty()
            skipped_u = stats_cols[3].empty()
            skipped_ocr = stats_cols[4].empty()
            errors_c = stats_cols[5].empty()
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            counts = {"processed": 0, "success": 0, "unsupported": 0, "ocr_missing": 0, "error": 0}
            type_counts = {}

            for idx, file_path in enumerate(files_to_scan):
                if st.session_state.stop_requested:
                    st.info("Scan stopped by user.")
                    break
                
                counts["processed"] += 1
                status_text.text(f"Processing ({counts['processed']}/{total_files}): {file_path.name}")
                progress_bar.progress(counts["processed"] / total_files)
                
                ext = file_path.suffix.lower()
                content = ""
                status = "Unknown"
                metadata = {}
                
                if ext == ".pdf":
                    content, metadata, status = extract_pdf_content(file_path)
                elif ext == ".docx" and scan_docx:
                    content, status = extract_docx_content(file_path)
                elif ext == ".doc":
                    content, status = "", "Skipped: .doc requires LibreOffice"
                else:
                    content, status = "", "Skipped: Unsupported format"

                # Post-processing
                if status == "Success" and not content.strip():
                    status = "Skipped: Scanned/No Text (OCR missing)"
                    counts["ocr_missing"] += 1
                elif "Skipped" in status:
                    counts["unsupported"] += 1
                elif "Error" in status:
                    counts["error"] += 1
                else:
                    counts["success"] += 1

                manual_name = identify_manual_name(content, file_path.name, str(file_path.parent), metadata)
                doc_type = classify_doc_type(content, file_path.name, str(file_path.parent))
                type_counts[doc_type] = type_counts.get(doc_type, 0) + 1
                
                # Confidence Logic
                confidence = "Low"
                clues = []
                if content.strip():
                    confidence = "Med"
                    clues.append("Text content")
                if metadata and metadata.get('/Title'):
                    meta_title = str(metadata['/Title']).strip().upper()
                    if meta_title and meta_title in manual_name.upper():
                        confidence = "High"
                        clues.append("Metadata match")
                if "manual" in file_path.name.lower() or "manual" in str(file_path.parent).lower():
                    clues.append("Keyword clue")

                res = {
                    "File Name": file_path.name,
                    "Relative Path": os.path.relpath(file_path, input_folder),
                    "File Type": doc_type,
                    "Extracted Manual/Equipment/System Name": manual_name,
                    "Confidence": confidence,
                    "Clues": ", ".join(clues),
                    "Notes": status
                }
                st.session_state.results.append(res)
                
                # Update UI Counters
                processed_c.metric("Processed", counts["processed"])
                success_c.metric("Success", counts["success"])
                skipped_u.metric("Unsupported", counts["unsupported"])
                skipped_ocr.metric("No Text", counts["ocr_missing"])
                errors_c.metric("Errors", counts["error"])
                
                if enable_debug:
                    st.write(f"DEBUG: Scanned {file_path.name} -> {manual_name} ({doc_type})")

            st.session_state.scanning = False
            st.success("Scan complete!")
            st.balloons()

# --- Results Display ---
if st.session_state.results:
    st.divider()
    df = pd.DataFrame(st.session_state.results)
    st.subheader(f"Scan Results ({len(df)} files)")
    st.dataframe(df, use_container_width=True)

    # --- Excel Export ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Results')
        
        # Summary Sheet
        summary_data = {
            "Metric": ["Total Files Found", "Successfully Scanned", "Total Skipped", "Errors"],
            "Value": [
                len(df),
                sum(1 for x in st.session_state.results if not any(s in x["Notes"] for s in ["Skipped", "Error"])),
                sum(1 for x in st.session_state.results if "Skipped" in x["Notes"]),
                sum(1 for x in st.session_state.results if "Error" in x["Notes"])
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, index=False, sheet_name='Summary', startrow=0)
        
        # Breakdown by Type
        type_summary = df["File Type"].value_counts().reset_index()
        type_summary.columns = ["File Type", "Count"]
        type_summary.to_excel(writer, index=False, sheet_name='Summary', startrow=len(summary_df)+2)

    st.download_button(
        label="📥 Download Results (Excel)",
        data=buffer.getvalue(),
        file_name=f"scan_results_{time.strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
elif not st.session_state.scanning:
    st.info("Enter a folder path and click 'Start Scan' to begin.")
