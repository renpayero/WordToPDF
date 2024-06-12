import os
import comtypes.client

def docx_to_pdf(docx_directory, pdf_directory):
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    
    if not os.path.exists(pdf_directory):
        os.makedirs(pdf_directory)

    for docx_file in os.listdir(docx_directory):
        if docx_file.endswith(".docx"):
            docx_directory = os.path.abspath(docx_directory)
            pdf_directory = os.path.abspath(pdf_directory)

            docx_path = os.path.join(docx_directory, docx_file)
            pdf_path = os.path.join(pdf_directory, docx_file.replace(".docx", ".pdf"))

            print("Intentando abrir:", docx_path)
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 es el código para PDF en Word
            doc.Close()
    print("Proceso finalizado.")
    word.Quit()

#RUTA ABSOLUTA DEL DIRECTORIO DONDE SE ENCUENTRAN LOS ARCHIVOS WORD
input_directory = './WORDs'
#RUTA ABSOLUTA DEL DIRECTORIO DONDE SE GUARDARÁN LOS ARCHIVOS PDF
output_directory = './PDFs'

docx_to_pdf(input_directory, output_directory)
