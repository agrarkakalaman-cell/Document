import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
from html import escape
from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK

from utils.formatter import format_dissertation

# Экранды барынша кеңейту
st.set_page_config(page_title="Киберқауіпсіздік кафедрасы", layout="wide")

TEXTS = {
    "kk": {
        "page_title": "Құжат сапасын бақылау",
        "description": "Word құжатын жүктеңіз. Қосымша оны таңдалған рәсімдеу талаптарына сәйкес өңдейді.",
        "settings_title": "Талаптарды енгізіңіз:",
        "title_settings": "Титулка деректері",
        "document_type": "Құжат түрі",
        "topic": "Тақырыбы",
        "author": "Орындаған",
        "head": "Кафедра меңгерушісі",
        "advisor": "Ғылыми жетекші",
        "consultant": "Консультант (бөлім)",
        "control": "Нормоконтроль",
        "reviewer": "Рецензент",
        "code": "Шифр",
        "year": "Жыл",
        "font_size": "Қаріп өлшемі",
        "line_spacing": "Жоларалық интервал",
        "left_margin": "Сол жақ өріс, мм",
        "right_margin": "Оң жақ өріс, мм",
        "top_margin": "Жоғарғы өріс, мм",
        "bottom_margin": "Төменгі өріс, мм",
        "first_line_indent": "Абзац шегінісі, мм",
        "upload_button": "Жүктеу",
        "process_button": "Өңдеу",
        "clear_button": "Тазалау",
        "download_word": "Word жүктеп алу",
        "download_pdf": "PDF жүктеп алу",
        "file_uploader_label": "Word құжатын жүктеу",
        "changed_warning": "Талаптар өзгерді. Қайтадан «Өңдеу» батырмасын басыңыз.",
        "processing_error": "Құжатты өңдеу қатесі:",
        "template_missing": "Титулка шаблоны табылмады:",
        "language_label": "Тіл",
        "logo_missing": "logo.png файлы табылмады",
        "preview_title": "Алдын ала қарау",
        "preview_empty": "DOCX",
        "fill_required": "Титулка үшін Тақырыбы, Орындаған, Шифр, Жетекші және Жыл өрістерін толтырыңыз.",
        "doc_types": [
            "Дипломдық жұмыс",
            "Өндірістік практика есебі",
            "Оқу практикасы есебі",
            "Зертханалық жұмыс",
            "Курстық жұмыс",
            "СӨЖ",
            "СОӨЖ",
        ],
    },
    "ru": {
        "page_title": "Нормоконтроль документов",
        "description": "Загрузите Word-документ. Приложение приведет его к выбранным стандартам оформления.",
        "settings_title": "Введите требования:",
        "title_settings": "Данные титульного листа",
        "document_type": "Тип документа",
        "topic": "Тема",
        "author": "Выполнил(а)",
        "head": "Заведующий кафедрой",
        "advisor": "Руководитель",
        "consultant": "Консультанты (раздел)",
        "control": "Нормоконтроль",
        "reviewer": "Рецензент",
        "code": "Шифр",
        "year": "Год",
        "font_size": "Размер шрифта",
        "line_spacing": "Межстрочный интервал",
        "left_margin": "Левое поле, мм",
        "right_margin": "Правое поле, мм",
        "top_margin": "Верхнее поле, мм",
        "bottom_margin": "Нижнее поле, мм",
        "first_line_indent": "Абзацный отступ, мм",
        "upload_button": "Загрузить",
        "process_button": "Обработать",
        "clear_button": "Очистить",
        "download_word": "Скачать Word",
        "download_pdf": "Скачать PDF",
        "file_uploader_label": "Загрузить Word-документ",
        "changed_warning": "Требования изменились. Нажмите «Обработать» заново.",
        "processing_error": "Ошибка обработки документа:",
        "template_missing": "Шаблон титульного листа не найден:",
        "language_label": "Язык",
        "logo_missing": "Файл logo.png не найден",
        "preview_title": "Предварительный просмотр",
        "preview_empty": "DOCX",
        "fill_required": "Заполните поля: Тема, Выполнил(а), Шифр, Руководитель и Год.",
        "doc_types": [
            "Дипломная работа",
            "Отчет по производственной практике",
            "Отчет по учебной практике",
            "Лабораторная работа",
            "Курсовая работа",
            "СРО",
            "СРСП",
        ],
    },
}

TEMPLATE_MAP = {
    "Дипломдық жұмыс": "templates/Diploma_kaz.docx",
    "Дипломная работа": "templates/Diploma_rus.docx",
    "Өндірістік практика есебі": "templates/Practic_kaz.docx",
    "Отчет по производственной практике": "templates/Practic_rus.docx",
    "Оқу практикасы есебі": "templates/Study_practic_kaz.docx",
    "Отчет по учебной практике": "templates/Study_practic_rus.docx",
    "Зертханалық жұмыс": "templates/Lab_kaz.docx",
    "Лабораторная работа": "templates/Lab_rus.docx",
    "Курстық жұмыс": "templates/Course_kaz.docx",
    "Курсовая работа": "templates/Course_rus.docx",
    "СӨЖ": "templates/SRO_kaz.docx",
    "СРО": "templates/SRO_rus.docx",
    "СОӨЖ": "templates/SRSP_kaz.docx",
    "СРСП": "templates/SRSP_rus.docx",
}

DEFAULT_SETTINGS = {
    "font_size": 14,
    "line_spacing": 1.5,
    "left_margin": 30,
    "right_margin": 20,
    "top_margin": 20,
    "bottom_margin": 20,
    "first_line_indent": 12.5,
}

if "lang" not in st.session_state:
    st.session_state["lang"] = "kk"

current_lang = st.session_state["lang"]
t = TEXTS[current_lang]

SESSION_DEFAULTS = {
    "docx_path": None,
    "error_message": None,
    "uploader_key": 0,
    "last_settings": None,
    "processed_settings": None,
    "preview_html": None,
    "document_type": t["doc_types"][0],
    "title_topic": "",
    "title_author": "",
    "title_head": "",
    "title_advisor": "",
    "title_consultants": "",
    "title_control": "",
    "title_reviewer": "",
    "title_code": "",
    "title_year": str(datetime.now().year),
}

for key, value in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value

for key, value in DEFAULT_SETTINGS.items():
    if key not in st.session_state:
        st.session_state[key] = value


def clear_all():
    st.session_state["docx_path"] = None
    st.session_state["error_message"] = None
    st.session_state["processed_settings"] = None
    st.session_state["last_settings"] = None
    st.session_state["preview_html"] = None
    st.session_state["uploader_key"] += 1

    for key, value in DEFAULT_SETTINGS.items():
        st.session_state[key] = value

    st.session_state["document_type"] = t["doc_types"][0]
    st.session_state["title_topic"] = ""
    st.session_state["title_author"] = ""
    st.session_state["title_head"] = ""
    st.session_state["title_advisor"] = ""
    st.session_state["title_consultants"] = ""
    st.session_state["title_control"] = ""
    st.session_state["title_reviewer"] = ""
    st.session_state["title_code"] = ""
    st.session_state["title_year"] = str(datetime.now().year)

    for folder in [Path("uploads"), Path("outputs")]:
        if folder.exists():
            for file_path in folder.glob("*"):
                if file_path.is_file():
                    try:
                        file_path.unlink()
                    except OSError:
                        pass


def replace_placeholders_in_docx(doc: Document, replacements: dict):
    def replace_in_paragraph(paragraph):
        text = paragraph.text
        if not any(key in text for key in replacements):
            return

        for key, value in replacements.items():
            text = text.replace(key, str(value))

        for run in paragraph.runs:
            run.text = ""

        if paragraph.runs:
            paragraph.runs[0].text = text
        else:
            paragraph.add_run(text)

    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph)


def append_docx_body(target_doc: Document, source_doc: Document):
    target_body = target_doc.element.body
    source_body = source_doc.element.body

    target_sect_pr = None
    if len(target_body) and target_body[-1].tag.endswith("sectPr"):
        target_sect_pr = target_body[-1]
        target_body.remove(target_sect_pr)

    p = target_doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)

    for element in source_body:
        if element.tag.endswith("sectPr"):
            continue
        target_body.append(element)

    if target_sect_pr is not None:
        target_body.append(target_sect_pr)


def create_document_with_title_page(template_path, formatted_docx_path, output_docx_path, data):
    template_path = Path(template_path)
    formatted_docx_path = Path(formatted_docx_path)
    output_docx_path = Path(output_docx_path)

    if not template_path.exists():
        raise FileNotFoundError(f"{t['template_missing']} {template_path}")

    title_doc = Document(template_path)
    replace_placeholders_in_docx(title_doc, data)

    main_doc = Document(formatted_docx_path)
    append_docx_body(title_doc, main_doc)

    output_docx_path.parent.mkdir(exist_ok=True)
    title_doc.save(output_docx_path)


def get_docx_preview_html(docx_path, max_paragraphs=220):
    document = Document(docx_path)
    html_parts = []

    for paragraph in document.paragraphs:
        text = paragraph.text

        if not text.strip():
            html_parts.append("<p>&nbsp;</p>")
        else:
            alignment = "left"
            if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                alignment = "center"
            elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                alignment = "right"
            elif paragraph.
