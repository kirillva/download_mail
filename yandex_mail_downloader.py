import os
import email
import imaplib
import argparse
from mailbox import mbox
from datetime import datetime, timedelta
import email
# import base64
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup

def merge_html_files_with_separators(input_dir, input_files, output_file, 
                                    start_separator="начало письма", 
                                    end_separator="конец письма",
                                    encoding='utf-8'):
    """
    Объединяет несколько HTML файлов в один с разделителями.
    
    Args:
        input_files (list): Список путей к входным HTML файлам
        output_file (str): Путь к выходному файлу
        start_separator (str): Текст для разделителя начала
        end_separator (str): Текст для разделителя конца
        encoding (str): Кодировка файлов
    """
    
    with open(output_file, 'w', encoding=encoding) as out_f:
        # Создаем основную структуру HTML для объединенного файла
        out_f.write('<!DOCTYPE html>\n<html>\n<head>\n')
        out_f.write('<meta charset="UTF-8">\n')
        out_f.write('<title>Объединенные письма</title>\n')
        out_f.write('</head>\n<body>\n')
        
        for i, input_file in enumerate(input_files, 1):
            try:
                with open(os.path.join(input_dir, input_file), 'r', encoding=encoding) as in_f:
                    html_content = in_f.read()
                
                # Извлекаем только тело письма (без тегов html, head, body)
                soup = BeautifulSoup(html_content, 'html.parser')
                
                # Удаляем лишние теги, если они есть
                for tag in ['html', 'head', 'body']:
                    if soup.find(tag):
                        tag_content = soup.find(tag)
                        if tag_content:
                            tag_content.unwrap()
                
                # Добавляем разделитель начала
                out_f.write(f'<div>\n')
                out_f.write(f'<h3>= {start_separator} {i} =</h3>\n')
                out_f.write('<hr>\n')
                
                # Добавляем содержимое письма
                out_f.write(str(soup))
                
                # Добавляем разделитель конца
                out_f.write(f'<hr>\n')
                out_f.write(f'<h3>= {end_separator} {i} =</h3>\n')
                out_f.write('</div>\n\n')
                
                print(f"Файл {input_file} успешно добавлен")
                
            except Exception as e:
                print(f"Ошибка при обработке файла {input_file}: {e}")
                # Добавляем сообщение об ошибке в выходной файл
                out_f.write(f'<div style="border: 2px solid red; padding: 10px; margin: 20px 0; background-color: #ffe6e6;">\n')
                out_f.write(f'<h3 style="color: red;">ОШИБКА: {start_separator} {i}</h3>\n')
                out_f.write(f'<p>Не удалось загрузить файл: {input_file}</p>\n')
                out_f.write(f'<p>Ошибка: {str(e)}</p>\n')
                out_f.write(f'<h3 style="color: red;">{end_separator} {i}</h3>\n')
                out_f.write('</div>\n\n')
        
        # Закрываем HTML структуру
        out_f.write('</body>\n</html>')
    
    print(f"\nВсе файлы объединены в: {output_file}")

def remove_styles_from_html(html_content):
    """
    Удаляет стили из HTML-контента.
    
    Args:
        html_content (str): Исходный HTML-контент
    
    Returns:
        str: Оптимизированный HTML без стилей
    """
    # Создаем объект BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Удаляем все теги <style>
    for style_tag in soup.find_all('style'):
        style_tag.decompose()
    
    # Удаляем атрибуты style у всех элементов
    for tag in soup.find_all(attrs={'style': True}):
        del tag['style']
    
    # Удаляем атрибуты class у всех элементов (если нужно)
    for tag in soup.find_all(attrs={'class': True}):
        del tag['class']
    
    # Возвращаем оптимизированный HTML
    return str(soup)

# Function for converting downloaded mailboxes from EML to Mbox format
def convert_to_mbox(mailbox_folder):
    for item in os.listdir(mailbox_folder):
        item_path = os.path.join(mailbox_folder, item)
        if os.path.isdir(item_path):
            # Recursively convert nested mailboxes
            convert_to_mbox(item_path)
        elif item.endswith('.eml'):
            mbox_path = os.path.join(mailbox_folder, f'{os.path.basename(os.path.normpath(mailbox_folder))}.mbox')
            mbox_file = mbox(mbox_path)
            with open(item_path, 'rb') as f:
                message = f.read()
            mbox_file.add(message)
            mbox_file.flush()

def process_eml_file(file_name, input_dir, output_dir='output'):
    """Обработка EML файла с сохранением вложений"""
    eml_file = os.path.join(input_dir, file_name)

    # Создаем выходную директорию
    os.makedirs(output_dir, exist_ok=True)
    
    with open(eml_file, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    
    # Информация о письме
    print("Subject:", msg['subject'])
    print("From:", msg['from'])
    print("Date:", msg['date'])

    head = 'Subject:' + msg['subject'] + '\nFrom:' + msg['from'] + '\nDate:' + msg['date'] + '\n'
    # Сохраняем текст письма
    text_body = ''
    html_body = ''
    
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            
            if content_type == 'text/plain':
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or 'utf-8'
                text_body = payload.decode(charset, errors='ignore')
                
            elif content_type == 'text/html':
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or 'utf-8'
                html_body = payload.decode(charset, errors='ignore')
                html_body = remove_styles_from_html(html_body)
            # elif part.get_filename():  # Вложение
            #     filename = part.get_filename()
            #     content = part.get_payload(decode=True)
                
            #     # Сохраняем вложение
            #     filepath = os.path.join(output_dir, filename)
            #     with open(filepath, 'wb') as f:
            #         f.write(content)
            #     print(f"Сохранено вложение: {filename}")
    else:
        # Не multipart письмо
        payload = msg.get_payload(decode=True)
        charset = msg.get_content_charset() or 'utf-8'
        text_body = payload.decode(charset, errors='ignore')
    
    if head:
        with open(os.path.join(output_dir, f'{file_name}.txt'), 'w', encoding='utf-8') as f:
            f.write(head)
    
    # Сохраняем текст письма
    if text_body:
        with open(os.path.join(output_dir, f'{file_name}.txt'), 'w', encoding='utf-8') as f:
            f.write(head + text_body)
    
    if html_body:
        with open(os.path.join(output_dir, f'{file_name}.html'), 'w', encoding='utf-8') as f:
            f.write(head + html_body)
    
    return {
        'subject': msg['subject'],
        'from': msg['from'],
        'text_body': text_body,
        'html_body': html_body
    }


if __name__ == '__main__':
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Download all mailboxes and their contents from a Yandex email account')
    parser.add_argument('username', type=str, help='Yandex email account username')
    parser.add_argument('password', type=str, help='Yandex email account password')
    parser.add_argument('-m', '--mbox', action='store_true', help='Convert downloaded mailboxes to Mbox format')
    parser.add_argument('-s', '--sync', action='store_true', help='Delete local email files that are not on the server')
    parser.add_argument('-a', '--max-age', type=int, default=-1, help='Only download emails newer than (since) X days')
    parser.add_argument('-e', '--exclude', type=str, nargs='+', help='List mailboxes to exclude from downloading')
    parser.add_argument('-i', '--include', type=str, nargs='+', help='List mailboxes to include (only those specified will be downloaded)')

    args = parser.parse_args()

    # Connect to the Yandex IMAP server over SSL
    imap_server = 'imap.yandex.com'
    imap_port = 993
    print(f'Connecting to {imap_server}:{imap_port}...')
    connection = imaplib.IMAP4_SSL(imap_server, imap_port)

    # Login to the Yandex email account
    print(f'Logging in as {args.username}..')
    try:
        connection.login(args.username, args.password)
    except Exception as e:
        print('Error: Failed to login to the Yandex email account')
        print(str(e))
        exit()

    # Get the list of mailboxes
    print('Listing account mailboxes..\n')
    try:
        connection.select()
        typ, data = connection.list()
    except Exception as e:
        print('Error: Failed to get the list of mailboxes')
        print(str(e))
        exit()

    # Create local account folder
    local_folder_name = args.username
    os.makedirs(local_folder_name, exist_ok=True)

    # Download all mailboxes and their contents locally
    for mailbox in data:
        # Mailbox name
        mailbox_name = mailbox.decode('utf-8').split(' "|" ')[-1].replace('"', '').replace('/', '_')
        mailbox_name_canonical = mailbox_name.replace('|', '/')

        if args.include is not None and mailbox_name_canonical not in args.include:
            continue
        if args.exclude is not None and mailbox_name_canonical in args.exclude:
            continue

        # Mailbox server path
        mailbox_path = mailbox_name.split('|')

        # Determine local mailbox folder path
        mailbox_folder_path = local_folder_name
        for folder in mailbox_path:
            mailbox_folder_path = os.path.join(mailbox_folder_path, folder)

        # Create necessary directories recursively
        os.makedirs(mailbox_folder_path, exist_ok=True)

        # Select mailbox
        try:
            connection.select('"' + mailbox_name + '"' if (' ' in mailbox_name or not mailbox_name.isascii()) else mailbox_name, readonly=True)

            if(args.max_age > 0):
                cutoff_date = (datetime.today() - timedelta(days=args.max_age)).strftime('%d-%b-%Y')
                typ, data = connection.uid('SEARCH', None, f'SINCE {cutoff_date}')
            else:
                typ, data = connection.uid('SEARCH', None, 'ALL')
        except Exception as e:
            print(f'Error: Failed to select mailbox {mailbox_name_canonical}')
            print(str(e))
            continue

        # Initialize counters
        saved = 0
        skipped = 0
        failed = 0
        removed = 0
        total = 0

        # Download mailbox contents
        print(f'Downloading contents of mailbox {mailbox_name_canonical}..')

        email_uids = data[0].split()

        for email_uid in email_uids:
            total += 1

            try:
                # Decode email UID
                email_uid = email_uid.decode()

                email_file_name = f'{email_uid}.eml'
                email_file_path = os.path.join(mailbox_folder_path, email_file_name)
                email_file_size = os.stat(email_file_path).st_size if os.path.isfile(email_file_path) else -1

                # Check if the email has already been downloaded
                if email_file_name in os.listdir(mailbox_folder_path) and email_file_size > 0:
                    skipped += 1
                    continue

                typ, data = connection.uid('FETCH', email_uid, '(RFC822)')
                email_content = data[0][1]

                # Parse the email message
                msg = email.message_from_bytes(email_content)

                # Save the email message in EML format
                encoding = msg.get_content_charset() or 'utf-8'
                with open(email_file_path, 'wb') as f:
                    f.write(email_content)
                    saved += 1
            except Exception as e:
                print(f'Error: Failed to download email with UID {email_uid} from mailbox {mailbox_name_canonical}')
                print(str(e))
                failed += 1
                continue

        if args.sync:
            for file in os.listdir(mailbox_folder_path):
                if file.endswith(".eml"):
                    email_uid = os.path.splitext(file)[0]
                    email_file_path = os.path.join(mailbox_folder_path, file)
                    if email_uid.encode() not in email_uids:
                        os.remove(email_file_path)
                        removed += 1

        print(f'  Saved: {saved}\n  Skipped: {skipped}\n  Failed: {failed}\n  Removed: {removed}\n  Total: {total}')

        # Convert to MBOX format if specified
        if args.mbox:
            print('Creating Mbox file..\n')
            convert_to_mbox(mailbox_folder_path)
        else:
            print('')

    # Close the connection to the Yandex email account
    print('Closing the connection..')
    try:
        connection.close()
        connection.logout()
    except Exception as e:
        print('Error: Failed to close the connection to the Yandex email account')
        print(str(e))
        exit()

    print('All mailboxes and their contents have been downloaded successfully!')

    for filepath in os.listdir(mailbox_folder_path):
        # Чтение EML файла
        process_eml_file(filepath, input_dir=mailbox_folder_path, output_dir=os.path.join('txt', mailbox_folder_path))
    

    merge_html_files_with_separators(os.path.join('txt', mailbox_folder_path),
                                    os.listdir(os.path.join('txt', mailbox_folder_path)), 
                                    os.path.join('txt', mailbox_folder_path, 
                                    'result.html'))
        
    print('All mailboxes and their contents have been decoded successfully!')