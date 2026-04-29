# 📄 Telegram File Converter Bot

A powerful **Python-based Telegram Bot** that enables seamless file conversion across multiple formats — including PDFs, Word documents, PowerPoint presentations, and images — directly within Telegram.

Built with a focus on **portability and simplicity**, this bot uses pure Python libraries and **does not rely on Windows COM or Microsoft Office**.

---

## 🚀 Features

### 📄 PDF Conversions

* Convert **PDF → Images**
* Convert **PDF → Word (DOCX)**
* Convert **PDF → Excel**
* **Compress PDF** files

### 🖼 Image Conversions

* Convert **Images → PDF**
* Convert **Image → Word (OCR-ready structure)**
* **Compress Images**

### 📊 Office Conversions

* Convert **Word → PDF**
* Convert **PPTX → PDF** *(pure Python, no PowerPoint required)*

---

## 🛠 Tech Stack

* **Python 3.10+**
* `python-telegram-bot` (v20+)
* `python-pptx`
* `python-docx`
* `Pillow (PIL)`
* `ReportLab`
* `PyMuPDF`

---

## 📲 How It Works

1. Open Telegram
2. Upload a file *(PDF / DOCX / PPTX / Image)*
3. Select a conversion option via inline buttons
4. Receive the converted file instantly 🎉

---

## 🔒 Platform Compatibility

### ✔ Supported

* Windows
* Linux *(if dependencies are installed)*
* Cloud environments *(deployment-ready)*

### ❌ Not Required

* Microsoft Office
* PowerPoint
* `win32com` / COM automation

---

## ⚠️ Limitations

* **PPTX → PDF conversion is image-based**

  * Animations and transitions are not preserved
  * Complex charts may lose fidelity

> ⚡ For perfect PPTX rendering, tools like LibreOffice or COM automation are required (not used in this project).

---

## 📌 Future Enhancements

* 🔄 LibreOffice-based high-fidelity PPTX → PDF
* 🔍 OCR support for scanned PDFs
* 📦 Bulk file download (ZIP support)
* ☁️ Cloud deployment (AWS / Render / Railway)
* 📊 User analytics & conversion history

---

## 🧑‍💻 Author

**Vasant Lohar**
Python Developer 
