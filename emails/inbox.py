import imaplib
import email

host = 'imap.gmail.com'
username = 'cute.ashwani@gmail.com'
password = 'nehatyagi'
path = r'D:\Process Improvement Project\python_programming\Downloads'
already_done = []

def get_inbox():
    mail = imaplib.IMAP4_SSL(host)
    mail.login(username, password)
    mail.select("inbox")
    _, search_data = mail.search(None, 'SUBJECT HILMS')
    my_message = None
    with open('already.txt','rb') as f:
        already_done = f.read()
    print(len(search_data[0].split()))
    for num in search_data[0].split():
    # num = search_data[0].split()[-1]
    if num is not in already_done:
        file_list = []
        email_data = {}
        _, data = mail.fetch(num, '(RFC822)')
        _, b = data[0]
        email_message = email.message_from_bytes(b)
        # for header in ['subject', 'to', 'from', 'date']:
        #     email_data[header] = email_message[header]
        for part in email_message.walk():
            # if part.get_content_type() == "text/plain":
            #     body = part.get_payload(decode=True)
            #     email_data['body'] = body.decode()
            # elif part.get_content_type() == "text/html":
            #     html_body = part.get_payload(decode=True)
            #     email_data['html_body'] = html_body.decode()
            fileName = part.get_filename()
            if bool(fileName):
                file_list.append(fileName)
                with open(path+'\\'+fileName,'wb') as f:
                    f.write(part.get_payload(decode=True))
        # email_data['attachments'] = file_list
        # my_message.append(email_data)
    return my_message


if __name__ == "__main__":
    my_inbox = get_inbox()