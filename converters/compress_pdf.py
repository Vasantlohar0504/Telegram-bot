import fitz  # PyMuPDF

def compress_pdf(input_pdf, output_pdf):
    """
    Compress PDF safely by cleaning and deflating streams.
    No image re-encoding (stable & cloud-safe).
    """
    doc = fitz.open(input_pdf)

    doc.save(
        output_pdf,
        garbage=4,      # remove unused objects
        deflate=True,   # compress streams
        clean=True      # optimize structure
    )

    doc.close()
