from pypdf import PdfReader, PdfWriter
import os

# PDF path input
pdf_path = input("PDF file path enter karo: ").strip()

if not os.path.exists(pdf_path):
    print("PDF file nahi mili.")
    exit()

reader = PdfReader(pdf_path)
total_pages = len(reader.pages)

print(f"\nIs PDF me total {total_pages} pages hain.")

pages_input = input(
    "Kaun kaun se pages chahiye? (Example: 1,2,5,10): "
).strip()

try:
    selected_pages = [
        int(page.strip()) for page in pages_input.split(",")
    ]
except ValueError:
    print("Invalid page numbers.")
    exit()

writer = PdfWriter()

for page_num in selected_pages:
    if 1 <= page_num <= total_pages:
        writer.add_page(reader.pages[page_num - 1])
    else:
        print(f"Page {page_num} exist nahi karta.")

output_file = "selected_pages.pdf"

with open(output_file, "wb") as f:
    writer.write(f)

print(f"\nDone! Output PDF saved as: {output_file}")