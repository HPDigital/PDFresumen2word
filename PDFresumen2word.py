"""
PDFresumen2word
"""

#!/usr/bin/env python
# coding: utf-8

# In[1]:


import fitz  # PyMuPDF
from docx import Document

def pdf_to_word(pdf_path, word_path, max_length=300):
    # Abrir el documento PDF
    pdf_document = fitz.open(pdf_path)
    text = ""

    # Leer cada página del documento PDF
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()

    # Aquí puedes implementar tu lógica para resumir el texto.
    # Para este ejemplo, simplemente truncaremos el texto a `max_length` caracteres.
    summary = text[:max_length] + '...' if len(text) > max_length else text

    # Crear un nuevo documento Word
    doc = Document()
    doc.add_paragraph(summary)

    # Guardar el documento Word
    doc.save(word_path)

    # Cerrar el documento PDF
    pdf_document.close()

# Uso del programa
pdf_path = "C:\\Users\\HP\\Downloads\\23 ADJTO Modelo Institucional - DIGITAL. (2).pdf"
word_path = "C:\\Users\\HP\\Downloads\\23 ADJTO Modelo Institucional - DIGITAL. (2).docx"
pdf_to_word(pdf_path, word_path)


# In[ ]:






if __name__ == "__main__":
    pass
