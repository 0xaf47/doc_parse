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
import argparse
import csv 
import json
import pandas as pd
import pypff
from bs4 import BeautifulSoup
import subprocess

def extract_text_from_pdf(pdf_path):

    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            num_pages = len(pdf_reader.pages)

            text = ""
            for page_num in range(num_pages):
                try:
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    
                    # Проверяем, был ли извлечен текст
                    if page_text:
                        text += page_text
                except:
                    text = ''
                    pass
            
            # Проверяем, есть ли извлеченный текст
            if text.strip():  # Если текст не пустой
                return clean_text(text)
            else:  # Если текст пустой, вызываем функцию для извлечения текста из изображений
                print("Текст не найден, пытаемся извлечь текст из изображений.")
                text = clean_text(extract_text_from_pdf_images(pdf_path))
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

    if not os.path.exists("temp"):
        os.makedirs("temp")

    command = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        "temp",
        file_path,
    ]

    try:
        subprocess.run(command, check=True) 
    except subprocess.CalledProcessError as e:
        print(f"Error converting file: {e}")
        return ""  # Или другое действие при ошибке

    output_docx_path = os.path.join("temp", os.path.splitext(os.path.basename(file_path))[0] + ".docx")

    try:
        doc = Document(output_docx_path)
        text = '\n'.join([para.text for para in doc.paragraphs])
    except:
        text = ""

    os.remove(output_docx_path)

    return text 



def extract_text_from_word(file_path: str) -> str:
    """Извлекает текст из .docx файла."""
    _, ext = os.path.splitext(file_path)

    if ext == '.docx':
        # Обработка .docx файлов
        try:
            doc = Document(file_path)
            return '\n'.join([para.text for para in doc.paragraphs])
        except:
            return ''

    elif ext == '.doc':
        # Обработка .doc файлов
        try:
            return extract_text_from_doc(file_path)
        except:
            return ''

    elif ext == '.txt':
        # Обработка .txt файлов
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read().strip()
        except:
            return ''

    else:
        raise ValueError(f"Unsupported file extension: {ext}")

def extract_text_from_excel(file_path:str)->str:
    _, file_extension = os.path.splitext(file)
    if file_extension.lower() == '.xls':
        if not os.path.exists("temp"):
            os.makedirs("temp")

        command = [
            "libreoffice",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            "temp",
            file_path,
        ]

        try:
            subprocess.run(command, check=True) 
        except subprocess.CalledProcessError as e:
            print(f"Error converting file: {e}")
            return ""  # Или другое действие при ошибке
        file_path = os.path.join("temp", os.path.splitext(os.path.basename(file_path))[0] + ".xlsx")
        print("converting to " + file_path)

    try:
        print ("working on xlsx")
        wb = load_workbook(filename=file_path)
        sheet = wb.active
        text = []
        for row in sheet.iter_rows(values_only=True):
            row_text = " ".join([str(cell) if cell is not None else "" for cell in row])
            text.append(row_text)
            print ("added")
        return ' '.join(text)
    except:
        return ''

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
    text = re.sub(r'\s+', ' ', text) 
    text = text.strip()
    return text


def sensitive_data_finder(text):
    email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    phone_pattern = r"(?:89|\+7|88)\d{9,10}"
    company_name_pattern = re.compile(r'((ООО|ИП|АО|ПАО|ОАО|МОУ|МБОУ|ЗАО|НПАО)( ?)(\"|\«| )[а-яА-Я0-9-_]+(\"|\»| ))', re.IGNORECASE)
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
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    match = re.search(email_pattern, header)
    return match.group(0) if match else None

def eml_parse(file):
    try:
        with open(file, 'rb') as fp:
            msg = BytesParser(policy=policy.default).parse(fp)

            return {
                'to': extract_email_address(msg['to']),
                'from': extract_email_address(msg['from'])
            }
    except:
        return {
            'to': [],
            'from': []
        }

#Можно еще Subject и прочее
def txt_email_parse(file):
    try:
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
    except:
        return {
            'to': [],
            'from': []
            }


def extract_domains_from_email_lists(email_data):
    domains = {}

    if 'from' in email_data and email_data['from'] != '':
        try:
            from_domain = ''.join(email_data['from']).split('@')[1]
        except:
            from_domain = ""
    else:
        from_domain = ""

    if 'to' in email_data:
        try:
            to_domain = email_data['to'].split('@')[1]
        except:
            to_domain = ''
    else:
        to_domain = ''

    domains = {'from': from_domain, 'to': to_domain}
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

    public_email_services = {'mail.ru', 'gmail.com', 'yandex.ru', 'rambler.ru', 'inbox.ru', 'bk.ru', 'ya.ru', 'list.ru', "yahoo.com"}
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

    domains = {'public_public': public_public,
               'public_private': public_private,
               'private_private': private_private
               }
    return domains


def extract_from_emails(files):
    pairs = []
    for file in files:
        _, file_extension = os.path.splitext(file)
        if file_extension.lower() == '.txt':
            addresses = txt_email_parse(file)
            domains = extract_domains_from_email_lists(addresses)
            pairs.append(domains)

        elif file_extension.lower() == '.eml':
            addresses = eml_parse(file)
            domains = extract_domains_from_email_lists(addresses)
            pairs.append(domains)

        elif file_extension.lower() == '.pst':
            addresses_pairs = pst_parse(file)
            for pair in addresses_pairs:
                domains = extract_domains_from_email_lists(pair)
                pairs.append(domains)
        else:
            continue
    clear_list = split_and_deduplicate_domains(pairs)
    clear_domains = categorize_domains(clear_list)
    return clear_domains

def export_json_to_csv(mode, json_data, csv_file_path):
    if mode == 'sensitive_data': 
        # Получаем заголовки из первого элемента JSON
        headers = ['file'] + list(next(iter(json_data.values())).keys())

        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            
            # Записываем заголовки
            writer.writerow(headers)
            
            # Записываем данные
            for file_path, data in json_data.items():
                row = [file_path] + [data[key] for key in headers[1:]]
                writer.writerow(row)
    if mode == 'emails':
        rows = []
        for key in json_data:
            for entry in json_data[key]:
                row = {'type': key, 'from': entry['from'], 'to': entry['to']}
                rows.append(row)
        df = pd.DataFrame(rows)
        df.to_csv(csv_file_path, index=False)



def pst_parse(pst_file_path):
    addresses = []
    pst_file = pypff.file()
    pst_file.open(pst_file_path)
    
    root = pst_file.get_root_folder()

    for folder in root.sub_folders:
        for sub in folder.sub_folders:
            for message in sub.sub_messages:

                headers = message.transport_headers
                raw_body = message.get_html_body()
                if raw_body != None:
                    #print("stststs" + str(headers))
                    #print(raw_body)
                    try:
                        raw_body = raw_body.decode('Windows-1251')
                        #print(raw_body)
                    except:
                        raw_body = raw_body.decode('utf-8')
                            #print(raw_body)
                    soup = BeautifulSoup(raw_body, "lxml")
                    plain_text = soup.get_text()
                    addresses.extend(extract_addresses(plain_text))

    return addresses



def extract_addresses(text):

    from_match = re.search(r"From: ([^<]+<([^>]*)>)", text)
    to_match = re.findall(r"To: ([^<]+<([^>]*)>)", text)

    addresses = []

    if from_match:
        from_address = from_match.group(2)

        for to_address in [match[1] for match in to_match]:
            addresses.append({
                "from": from_address,
                "to": to_address
            })

    return addresses


if __name__ == "__main__":
    extensions = {"word": ["txt", "docx", "doc"],
                  "excel":[ "xlsx", "xls"],
                  "pdf": ["pdf"],
                  "email": ["pst", "txt", "eml"]}

    parser = argparse.ArgumentParser(description="Extract emails or sensitive data from a directory.")
    parser.add_argument('--input_path', type=str, required=True, help='Path to the input directory')
    parser.add_argument('--mode', type=str, choices=['emails', 'sensitive_data'], required=True,
                        help='Mode of operation: "emails" (search for From-To emails and domains in addresses) or "sensitive_data" (search for phones, names etc)')
    args = parser.parse_args()
    path = args.input_path

    files = find_files_by_extensions(path, extensions)
    results = {}

    if args.mode == "emails":
        results = extract_from_emails(files['email'])




    if args.mode == "sensitive_data":

                 
        for file in files['excel']:
            text = extract_text_from_excel(file)
            sensitive_data = sensitive_data_finder(text)
            results[file] = sensitive_data



    csv_file_path = 'output2.csv'
    export_json_to_csv(args.mode, results, csv_file_path)   
    with open('output1.json', 'w', encoding='utf-8') as json_file:
        json.dump(results, json_file, ensure_ascii=False, indent=4)


                
'''
        for file in files['word']:
            text = extract_text_from_word(file)
            sensitive_data = sensitive_data_finder(text)
            results[file] = sensitive_data

        for file in files['pdf']:
            text = extract_text_from_pdf(file)
            sensitive_data = sensitive_data_finder(text)
            results[file] = sensitive_data
'''        
    
