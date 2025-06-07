# PDF→Word Converter

A simple Gradio‐based web app to convert one or more PDF files into DOCX,  
format all tables as **Table Grid**, extract any URLs and append them at the end,  
and—if you upload multiple PDFs—bundle all outputs into a ZIP.

---

## 🚀 Features

- **Batch or single** PDF → DOCX conversion  
- Automatically style all tables as **Table Grid**  
- Extracts both annotation and inline URLs into a “Daftar Link” section  
- If multiple PDFs are uploaded, outputs are packaged into a ZIP  

---

## 📦 Installation

1. **Clone this repo**  
   ```bash
   git clone https://github.com/salmanalfarisi11/pdf-to-docx.git
   cd pdf-to-docx
   ```

2. Create and activate a virtual environment:

   ```bash
    python -m venv .venv
    source .venv/bin/activate   # Linux/macOS
    .venv\Scripts\activate      # Windows
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Running Locally

Launch the app on your machine:
   ```bash
   python app.py
   ```
By default, it will start on http://127.0.0.1:7860/. Open that URL in your browser to access the interface.

## 🎯 Usage

1. Open the localhost URL shown in your terminal.
2. Drag & drop one or more PDF files into the Upload PDF(s) panel.
3. Click Convert.
4. When ready, click ⬇️ Download Output to download either a single DOCX or a ZIP of all DOCXs.

## 🛠️ Dependencies

- Python 3.8+
- Gradio ≥ 5.33.0
- PyMuPDF
- pdf2docx
- python-docx


## 📄 License

This project is licensed under the [MIT License](LICENSE).

---


## 🖋️ Author & Credits

Developed by **[Salman Alfarisi](https://github.com/salmanalfarisi11)** © 2025  
- GitHub: [salmanalfarisi11](https://github.com/salmanalfarisi11)  
- LinkedIn: [salmanalfarisi11](https://linkedin.com/in/salmanalfarisi11)  
- Instagram: [faris.salman111](https://instagram.com/faris.salman111)  
