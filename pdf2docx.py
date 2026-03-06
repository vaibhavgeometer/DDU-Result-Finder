import os
import sys

# IMPORTANT: Remove the current directory from the path so we don't accidentally import this very script instead of the actual `pdf2docx` package!
current_dir = os.path.dirname(os.path.abspath(__file__))
if '' in sys.path:
    sys.path.remove('')
if current_dir in sys.path:
    sys.path.remove(current_dir)

from pdf2docx import Converter

pdf_files = [
    os.path.join("Information", "Saved Results", "1.pdf"),
    os.path.join("Information", "Saved Results", "2.pdf"),
    os.path.join("Information", "Saved Results", "3.pdf")
]

def main():
    out_dir = os.path.join("Information", "pdf2docx")
    os.makedirs(out_dir, exist_ok=True)
    
    for pdf_file in pdf_files:
        if not os.path.exists(pdf_file):
            print(f"File not found: {pdf_file}")
            continue
            
        docx_file = os.path.join(out_dir, os.path.basename(pdf_file).replace('.pdf', '.docx'))
        
        # Skip if already converted
        if os.path.exists(docx_file):
            print(f"Already converted: {docx_file}")
            continue
            
        print(f"Converting {pdf_file} to {docx_file} using pdf2docx...")
        try:
            cv = Converter(pdf_file)
            cv.convert(docx_file, start=0, end=None)
            cv.close()
            print(f"✅ Successfully converted: {docx_file}\n")
        except Exception as e:
            print(f"❌ Failed to convert {pdf_file} to DOCX. Error: {e}\n")

if __name__ == "__main__":
    main()
