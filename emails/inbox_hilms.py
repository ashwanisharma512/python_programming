import imaplib
import email
import pickle

host = 'imap.gmail.com'
username = 'cute.ashwani@gmail.com'
password = 'nehatyagi'
path = r'D:\Process Improvement Project\python_programming'
already_done = []

mail = imaplib.IMAP4_SSL(host)
mail.login(username, password)
x = mail.list()
print(x)
mail.select("inbox")
_, search_data = mail.search(None, 'FROM ashwanisharma512@gmail.com')
my_message = None
with open('already.dat','rb') as f:
    already_done = pickle.load(f)
print(len(search_data[0].split()))
for num in search_data[0].split():
    if num not in already_done:
        _, data = mail.fetch(num, '(RFC822)')
        _, b = data[0]
        email_message = email.message_from_bytes(b)
        try:
            for part in email_message.walk():
                fileName = part.get_filename()
                if bool(fileName):
                    with open(path+'\\'+fileName,'wb') as f:
                        f.write(part.get_payload(decode=True))
        except:
            pass
    already_done.append(num)
with open('already.dat','wb') as f:
    pickle.dump(already_done,f)