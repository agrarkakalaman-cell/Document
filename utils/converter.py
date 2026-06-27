from pathlib import Path
from docx2pdf import convert


def convert_docx_to_pdf(docx_path, output_dir):
    docx_path = Path(docx_path)
    output_dir = Path(output_dir)

    output_dir.mkdir(exist_ok=True)

    pdf_path = output_dir / f"{docx_path.stem}.pdf"

    try:
        convert(str(docx_path), str(pdf_path))
    except Exception as error:
        print("Ошибка PDF-конвертации:", error)
        return None

    if pdf_path.exists():
        return pdf_path

    return None