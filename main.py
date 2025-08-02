import os
import tempfile
import fitz  # PyMuPDF
import win32com.client
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import re

def natural_key(text):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r'(\d+)', text)]

# === הגדרות נתיבים ===
input_folder = r"C:\Users\einav\Downloads\word_file"
output_pdf_path = r"C:\Users\einav\Downloads\output_with_toc.pdf"
hebrew_font_path = r"C:\Users\einav\Downloads\NotoSansHebrew-Regular.ttf"

def rtl(text):
    return text[::-1]

# רישום פונט עברי
pdfmetrics.registerFont(TTFont("HebrewFont", hebrew_font_path))

# אתחול Word COM
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# פתיחת מסמך ממוזג
merged_pdf = fitz.open()
toc_entries = []
current_page_number = 0

# המרת Word ל־PDF ומיזוג
with tempfile.TemporaryDirectory() as temp_dir:
    for filename in sorted(os.listdir(input_folder), key=natural_key):
        full_path = os.path.join(input_folder, filename)
        name_without_ext = os.path.splitext(filename)[0]
        ext = filename.lower().split(".")[-1]

        if ext not in ("docx", "pdf"):
            continue  # נתעלם מקבצים לא רלוונטיים

        try:
            if ext == "docx":
                temp_pdf = os.path.join(temp_dir, f"{name_without_ext}.pdf")
                doc = word.Documents.Open(full_path)
                doc.SaveAs(temp_pdf, FileFormat=17)  # PDF
                doc.Close()
                part_pdf = fitz.open(temp_pdf)
            elif ext == "pdf":
                part_pdf = fitz.open(full_path)

            num_pages = len(part_pdf)
            toc_entries.append((name_without_ext, current_page_number))
            merged_pdf.insert_pdf(part_pdf)
            current_page_number += num_pages
            part_pdf.close()

        except Exception as e:
            print(f"שגיאה בקובץ {filename}: {e}")

word.Quit()

# יצירת עמוד תוכן עניינים
packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=A4)
can.setFont("HebrewFont", 14)
can.drawRightString(550, 800, rtl("תוכן עניינים"))
can.setFont("HebrewFont", 11)

y_positions = []
y = 770

for title, page_number in toc_entries:
    title_text = rtl(title)
    page_text = str(page_number + 2)
    display_text = f"{title_text} .......... {page_text}"
    text_width = pdfmetrics.stringWidth(display_text, "HebrewFont", 11)
    x_start = 550 - text_width
    can.drawString(x_start, y, display_text)
    y_positions.append((x_start, y, text_width, page_number + 1))
    y -= 20

can.save()
packet.seek(0)

# מיזוג תוכן עניינים עם המסמך
toc_pdf = fitz.open("pdf", packet.read())
final_doc = fitz.open()
final_doc.insert_pdf(toc_pdf)
final_doc.insert_pdf(merged_pdf)

# הוספת קישורים לתוכן עניינים
toc_page = final_doc[0]
for x, y, width, page_target in y_positions:
    rect = fitz.Rect(x, A4[1] - y - 12, x + width, A4[1] - y + 2)
    toc_page.insert_link({
        "kind": fitz.LINK_GOTO,
        "from": rect,
        "page": page_target
    })

# שמירה
final_doc.save(output_pdf_path)
final_doc.close()

print("נוצר קובץ עם תוכן עניינים מבוסס על שמות קבצי Word ו־PDF, כולל קישורים.")
