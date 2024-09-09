import PyPDF2
from docx import Document
from openpyxl import load_workbook
import pytesseract
from PIL import Image
import io
import re
import os
import logging
import sys
import olefile

def extract_text_from_pdf(pdf_path):
    """
    Извлекает текст из PDF-файла.

    Args:
        pdf_path: Путь к PDF-файлу.

    Returns:
        Строка с извлеченным текстом или None, если текст извлечь не удалось.
    """
    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            num_pages = len(pdf_reader.pages)

            text = ""
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                
                # Проверяем, был ли извлечен текст
                if page_text:
                    text += page_text
            
            # Проверяем, есть ли извлеченный текст
            if text.strip():  # Если текст не пустой
                return clean_text(text)
            else:  # Если текст пустой, вызываем функцию для извлечения текста из изображений
                print("Текст не найден, пытаемся извлечь текст из изображений.")
                text = clean_text(extract_text_from_pdf_images(pdf_path))
                print(text)
                return(text)
    except Exception as e:
        print(f"Ошибка при чтении PDF как текста: {e}")
        try:
            return clean_text(extract_text_from_pdf_images(pdf_path))
        except Exception as e:
            print(f"Ошибка при извлечении текста из изображений: {e}")
            return None

def extract_text_from_pdf_images(pdf_path):
    """
    Извлекает текст из PDF-файла, используя PIL и pytesseract, если 
    не удалось прочитать как текст.

    Args:
        pdf_path: Путь к PDF-файлу.

    Returns:
        Строка с извлеченным текстом или None, если текст извлечь не удалось.
    """
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path)
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image, lang='rus+eng')
        return text
    except Exception as e:
        print(f"Ошибка при конвертации PDF в изображения: {e}")
        return None

def extract_text_from_doc(file_path):
    """Извлекает текст из .doc файла."""
    if not olefile.isOleFile(file_path):
        raise ValueError("Файл не является OLE-файлом.")
    
    ole = olefile.OleFileIO(file_path)
    text_stream = ole.openstream('WordDocument')
    
    # Читаем содержимое потока
    content = text_stream.read()
    
    # Пытаемся декодировать содержимое как текст
    try:
        return content.decode('utf-16').strip()
    except UnicodeDecodeError:
        return content.decode('latin1').strip()  # Попробуем альтернативное кодирование

def extract_text_from_word(file_path: str) -> str:
    """Извлекает текст из .doc или .docx файла."""
    _, ext = os.path.splitext(file_path)

    if ext == '.docx':
        # Обработка .docx файлов
        doc = Document(file_path)
        return '\n'.join([para.text for para in doc.paragraphs])

    elif ext == '.doc':
        # Обработка .doc файлов
        text = extract_text_from_doc(file_path)
        print(text)
        return text

    elif ext == '.txt':
        # Обработка .txt файлов
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()

    else:
        raise ValueError(f"Unsupported file extension: {ext}")

def extract_text_from_xlsx(file_path:str)->str:
    wb = load_workbook(filename=file_path)
    sheet = wb.active
    text = []
    for row in sheet.iter_rows(values_only=True):
        row_text = " ".join([str(cell) if cell is not None else "" for cell in row])
        text.append(row_text)
    return ' '.join(text)

def text_preparation(text:str)->str:
    text = text.replace('\n',' ')
    while '  ' in text:
        text = text.replace('  ',' ')
    return text

def find_files_by_extensions(directory, extensions):
    found_files = {"word": [],
                  "excel":[],
                  "pdf": []}


    for root, _, files in os.walk(directory):
        for file in files:
          filename, extension = os.path.splitext(file)
          for key in extensions:
              
              if extension[1:].lower() in [ext.lower() for ext in extensions[key]]:
                found_files[key].append(os.path.join(root, file))
    return found_files

def clean_text(text):
    # Убираем знаки табуляции
    text = text.replace('\t', ' ').replace('\n', " ")
    
    # Убираем лишние пробелы (более одного пробела)
    text = re.sub(r'\s+', ' ', text)
    
    # Убираем пробелы в начале и конце строки
    text = text.strip()
    
    return text

if __name__ == "__main__":
    extensions = {"word": ["txt", "doc", "docx"],
                  "excel":["xls", "xlsx", "csv"],
                  "pdf": ["pdf"]}
    path = input("Введите путь: ")
    files = find_files_by_extensions(path, extensions)
    #print(files)
    for file in files['word']:
        print(file)
        text = extract_text_from_word(file)
        #print(text)
    
