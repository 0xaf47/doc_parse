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
    domains = {}
    to_domain_list = []
    print(email_data['from'])
    if 'from' in email_data and email_data['from'] != '':
        from_domain = email_data['from'][0].split('@')[-1]
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

if __name__ == '__main__':
    path = input("Введите путь к файлам: ")
    emails =  find_files_by_extensions(path, ["eml", "txt"])
    pairs = []
    for email in emails:
        addresses = email_parse(email)
        domains = extract_domains_from_email_lists(addresses)
        pairs.append(domains)

    clear_list = split_and_deduplicate_domains(pairs)

    public_public, public_private, private_private = categorize_domains(clear_list)

    print("Public-Public:")
    for item in public_public:
        print(item)

    print("\nPublic-Private:")
    for item in public_private:
        print(item)

    print("\nPrivate-Private:")
    for item in private_private:
        print(item) 

 
