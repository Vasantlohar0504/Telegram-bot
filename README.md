📄 Telegram File Converter Bot
A Python-based Telegram Bot that converts files between multiple formats such as PDF, Word, PPTX, Images, and also supports compression — all directly inside Telegram.
This bot is built using python-telegram-bot, pure Python converters, and avoids Windows COM dependencies.

🚀 Features
📄 PDF Conversions
PDF → Images
PDF → Word
PDF → Excel
PDF Compression

🖼 Image Conversions
Images → PDF
Image → Word (OCR-ready structure)
Image Compression

📄 Office Conversions
Word → PDF
PPTX → PDF (Python-only, no COM, no PowerPoint required)

🛠 Tech Stack
Python 3.10+
python-telegram-bot (v20+)
python-pptx
python-docx
Pillow (PIL)
ReportLab
PyMuPDF (for PDFs)

📲 How to Use
Open Telegram
Send a file (PDF / DOCX / PPTX / Image)
Choose conversion from inline buttons
Get the converted file instantly 🎉

⚠️ Limitations
PPTX → PDF uses image-based rendering
Animations, transitions, and complex charts may not render perfectly
Full PowerPoint fidelity requires LibreOffice or COM (not used here)
🔒 No Windows COM Required

✔ Works on:
Windows

✔ Does NOT require:
Microsoft Office
PowerPoint
win32com

📌 Future Improvements
LibreOffice-based perfect PPTX → PDF
OCR for scanned PDFs
ZIP bulk downloads
Cloud deployment
User history & analytics


👨‍💻 Author
Vasant Lohar
Python | Automation | Telegram Bots

⭐ Support
If you like this project:
⭐ Star the repository
🍴 Fork it
🧑‍💻 Contribute improvements
