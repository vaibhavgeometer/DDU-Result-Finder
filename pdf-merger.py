import os
from pypdf import PdfReader, PdfWriter

# Folder containing your 397 PDF files
pdf_folder = "C:/Users/Vaibhav/Documents/Projects/ddu-result-code/Saved_PDFs"  # Change this to your folder path

# Output file
output_pdf_path = "1st Sem Result.pdf"

# Initialize PDF writer
writer = PdfWriter()

# Get all PDF files in the folder
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
pdf_files.sort()  # Optional: sort alphabetically

# Loop through each file and add the first page
for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    try:
        reader = PdfReader(pdf_path)
        if len(reader.pages) > 0:
            writer.add_page(reader.pages[0])
    except Exception as e:
        print(f"Error processing {pdf_file}: {e}")

# Write the combined PDF
with open(output_pdf_path, "wb") as out_file:
    writer.write(out_file)

print(f"Merged PDF saved to {output_pdf_path}")
