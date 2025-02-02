import os
import smtplib
import json
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_mail_intern():
    global smtp_server, smtp_port, smtp_user, smtp_password, email
    # Pfade und Einstellungen
    json_path = './files/zustaendigkeiten.json'
    folder_path = './files/bewertungsBoegen'  # Pfad zu dem Ordner mit den Dateien
    smtp_server = 'smtp.web.de'
    smtp_port = 587
    smtp_user = 'leanderniehoff@web.de'
    smtp_password = #Password einfügen
    # Lade die Zuständigkeiten aus der JSON-Datei
    with open(json_path, 'r') as file:
        data = json.load(file)

    # Funktion zum Versenden einer E-Mail
    def send_email(to_email, to_name, attachment_path):
        # Erstelle eine MIMEMultipart Nachricht
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Subject'] = 'Ihre Datei'

        # Füge den Textteil hinzu
        body = f'Hallo {to_name},\n\nAnbei finden Sie die gewünschte Datei.\n\nBeste Grüße\nIhr Team'
        msg.attach(MIMEText(body, 'plain'))

        # Füge den Anhang hinzu
        filename = os.path.basename(attachment_path)
        attachment = open(attachment_path, 'rb')

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {filename}')

        msg.attach(part)

        # Verbinde dich mit dem SMTP-Server und sende die E-Mail
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        text = msg.as_string()
        server.sendmail(smtp_user, to_email, text)
        server.quit()

    # Sende E-Mails für alle Dateien im Ordner
    for item in data:
        zustaendigkeit = item['Zustaendigkeit']
        name = item['Name']
        email = item['Email']
        # Hier gehen wir davon aus, dass die Dateien im Ordner nach dem Namen der Person benannt sind
        file_path = os.path.join(folder_path, f'{zustaendigkeit}.xlsx')

        if os.path.exists(file_path):
            send_email(email, name, file_path)
        else:
            print(f'Datei für {zustaendigkeit} nicht gefunden.')
    print('E-Mails wurden gesendet.')

