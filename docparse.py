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
from email import policy
from email.parser import BytesParser

def extract_text_from_pdf(pdf_path):

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
    """Извлекает текст из .docx файла."""
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

def extract_text_from_excel(file_path:str)->str:
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
                  "pdf": [],
                   "email": []}


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



def sensitive_data_finder(text:str)->str:
    email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    phone_pattern = r'((\+7\d{10})|([^0-9](7|8)\d{10}[^0-9])|((7|8)\(\d{3}\)\d{7})|( \d{3} \d{2} \d{2} )|( \d{3}-\d{2}-\d{2} )|((7|8) \(\d{3}\) \d{3}-\d{2}-\d{2})|((7|8)-\d{3}-\d{3}-\d{2}-\d{2})|([^0-9]\d{7}[^0-9])|((7|8) \d{3} \d{3} \d{2} \d{2})|(\(\d{3}\) \d{3}-\d{2}-\d{2})|((7|8) \(\d{3}\) \d{3} \d{2} \d{2}))'
    company_name_pattern = re.compile(r'((ООО|ИП|АО|ПАО|НКО|ОП|ТСЖ|НАО|ЗАО|НПАО)( ?)(\"|\«| )[а-яА-Я0-9-_]+(\"|\»| ))', re.IGNORECASE)
    ul_inn_pattern = r'[^0-9]\d{10}[^0-9]'
    fl_inn_pattern = r'[^0-9]\d{12}[^0-9]'
    kpp_pattern = r'[^0-9]\d{9}[^0-9]'
    bik_pattern = r'[^0-9]04\d{7}[^0-9]'
    snils_pattern = r'\d{3}-\d{3}-\d{3} \d{2}'
    Full_FIO_pattern = r'[А-Я][а-я]+ [А-Я][а-я]+ [А-Я][а-я]+'
    Abr_FIO_patterns = r'([А-Я](\.|\. | )[А-Я](\.|\. | )[А-Я][а-я]+)'

    emails = re.findall(email_pattern,text)
    phones = re.findall(phone_pattern,text)
    company_names = re.findall(company_name_pattern,text)
    ul_inn = re.findall(ul_inn_pattern,text)
    fl_inn = re.findall(fl_inn_pattern,text)
    kpp = re.findall(kpp_pattern,text)
    bik = re.findall(bik_pattern,text)
    snilses = re.findall(snils_pattern,text)
    Full_FIOS = re.findall(Full_FIO_pattern,text)
    Abr_FIOS = re.findall(Abr_FIO_patterns,text)
    password_sentenses = extract_password_sentences(text)

    phones = [phone[0].strip() for phone in phones if phone[0] != '']
    company_names = [company_name[0].strip() for company_name in company_names if company_name[0] != '']
    ul_inn = [inn[1:-1].strip() for inn in ul_inn if inn != '']
    fl_inn = [inn[1:-1].strip() for inn in fl_inn if inn != '']
    kpp = [inn[1:-1].strip() for inn in kpp if inn != '' and not inn[1:-1].startswith('04')]
    bik = [inn[1:-1].strip() for inn in bik if inn != '']
    snilses = [snils.strip() for snils in snilses if snils != '']
    Full_FIOS = [FIO.strip() for FIO in Full_FIOS if FIO != '']
    Abr_FIOS = [FIO[0].strip() for FIO in Abr_FIOS if FIO != '']

    sensitive_data = {
            'phones': list(set(phones)),
            'emails': list(set(emails)),
            'company_names': list(set(company_names)),
            'ul_inn' : list(set(ul_inn)),
            'fl_inn' : list(set(fl_inn)),
            'kpp' : list(set(kpp)),
            'bik' : list(set(bik)),
            'snilses' : list(set(snilses)),
            'Full_FIOS' : list(set(Full_FIOS)),
            'Abr_FIOS' : list(set(Abr_FIOS)),
            'password_sentenses' : list(set(password_sentenses))
            }
    return sensitive_data

def extract_password_sentences(text):
    pattern = r'(?<![A-Z])([A-Z][^.!?]*?\b(пароль|password|пароли|passwords|п|пасс|pass)\b[^.!?]*?[.!?])'
    matches = re.findall(pattern, text, re.IGNORECASE)
    
    sentences = [match[0].strip() for match in matches]
    
    return sentences






def extract_email_address(header):
    # Регулярное выражение для извлечения адреса электронной почты
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    match = re.search(email_pattern, header)
    return match.group(0) if match else None

def eml_parse(file):
    with open(file, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)

        return {
            'to': extract_email_address(msg['to']),
            'from': extract_email_address(msg['from'])
        }
#Можно еще Subject и прочее
def txt_parse(file):
    with open(file, 'r', encoding='utf-8') as file:
        content = file.read()

    # Регулярки для поиска адресов
    to_pattern = r'To:\s*([\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}(?:,\s*[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,})*)'
    from_pattern = r'From:\s*([\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,})'
    to_matches = re.findall(to_pattern, content)
    from_matches = re.findall(from_pattern, content)

    # Обработка найденных адресов
    to_addresses = []
    from_addresses = []

    for match in to_matches:
        # Извлечение адресов из строки
        addresses = [email.strip() for email in re.split(r',\s*|\s*;\s*', match) if email]
        to_addresses.extend(addresses)

    for match in from_matches:
        addresses = [email.strip() for email in re.split(r',\s*|\s*;\s*', match) if email]
        from_addresses.extend(addresses)

    return {
        'to': to_addresses,
        'from': from_addresses
    }

def email_parse(file):
    # Получение расширения файла
    _, file_extension = os.path.splitext(file)

    if file_extension.lower() == '.txt':
        result = txt_parse(file)
    elif file_extension.lower() == '.eml':
        result = eml_parse(file)
    else:
        result = "Неизвестное расширение файла!"

    return result


def extract_domains_from_email_lists(email_data):
    domains = {}
    to_domain_list = []

    if 'from' in email_data and email_data['from'] != '':
        from_domain = ''.join(email_data['from']).split('@')[-1]
    else:
        from_domain = ""

    if 'to' in email_data:
        for email in email_data['to']:
            domain = email.split('@')[-1]
            to_domain_list.append(domain)
    else:
        to_domain_list = ""

    domains = {'from': from_domain, 'to': to_domain_list}
    return domains

def split_and_deduplicate_domains(json_data_list):

    unique_data = set()

    for data in json_data_list:
        if isinstance(data['to'], list):
            for to_domain in data['to']:
                new_data = {'from': data['from'], 'to': to_domain}
                unique_data.add(tuple(sorted(new_data.items())))
        else:
            unique_data.add(tuple(sorted(data.items())))

    result = []
    for from_to_pair in unique_data:
        result.append(dict(from_to_pair))

    return result


def categorize_domains(json_data_list):

    public_email_services = {'mail.ru', 'gmail.com', 'yandex.ru', 'rambler.ru', 'inbox.ru', 'bk.ru', 'ya.ru', 'list.ru'}
    public_public = []
    public_private = []
    private_private = []

    for data in json_data_list:
        from_domain = data['from']
        to_domain = data['to']

        is_from_public = from_domain in public_email_services
        is_to_public = to_domain in public_email_services

        if is_from_public and is_to_public:
            public_public.append(data)
        elif is_from_public or is_to_public:
            public_private.append(data)
        else:
            private_private.append(data)

    return public_public, public_private, private_private


def extract_text_from_email(file_path):
    pairs = []
    addresses = email_parse(file_path)
    domains = extract_domains_from_email_lists(addresses)
    pairs.append(domains)

    clear_list = split_and_deduplicate_domains(pairs)

    public_public, public_private, private_private = categorize_domains(clear_list)

    print (public_public)
    print (public_private)
    print (private_private)

if __name__ == "__main__":
    extensions = {"word": ["txt", "docx"],
                  "excel":[ "xlsx"],
                  "pdf": ["pdf"],
                  "email": ["eml", "txt"]}

    path = input("Введите путь: ")
    files = find_files_by_extensions(path, extensions)
    for file in files['email']:
        extract_text_from_email(file)
    for file in files['word']:
        text = extract_text_from_word(file)
        sensitive_data = sensitive_data_finder(text)
        for data in sensitive_data:
            print (sensitive_data[data])
 
    for file in files['excel']:
        text = extract_text_from_excel(file)
        sensitive_data = sensitive_data_finder(text)
        for data in sensitive_data:
            print (sensitive_data[data])
 
    for file in files['pdf']:
        text = extract_text_from_pdf(file)
        sensitive_data = sensitive_data_finder(text)
        for data in sensitive_data:
            print (sensitive_data[data])



    
