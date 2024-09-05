from email import policy
from email.parser import BytesParser
import os
import re
eml_file = "/home/x/base/Securelist - Аналитика и отчеты о киберугрозах от «Лаборатории Касперского».eml"


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
    to_pattern = r'(?:^|\n)(?:To|to|FROM|from):?\s*([^\n]*)'
    from_pattern = r'(?:^|\n)(?:From|from|TO|to):?\s*([^\n]*)'
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

def find_files_by_extensions(directory, extensions):
    found_files = []
    for root, _, files in os.walk(directory):
        for file in files:
          filename, extension = os.path.splitext(file)
          if extension[1:].lower() in [ext.lower() for ext in extensions]:
            found_files.append(os.path.join(root, file))
    return found_files

def extract_domains_from_email_lists(email_data):
    domain_list = []
    for field in ['to', 'from']:  # Проходим по полям 'to' и 'from'
        if field in email_data:
            for email in email_data[field]:
                domain = email.split('@')[-1]
                domain_list.append(domain)

    return domain_list


if __name__ == '__main__':
    path = input("Введите путь к файлам: ")
    emails =  find_files_by_extensions(path, ["eml", "txt"])
    for email in emails:
        addresses = email_parse(email)
        domains = extract_domains_from_email_lists(addresses)
        print (domains)

 
