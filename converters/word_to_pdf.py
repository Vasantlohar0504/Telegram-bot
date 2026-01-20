from docx2pdf import convert

def word_to_pdf(word_path, output_path):
    convert(word_path, output_path)
    return output_path
