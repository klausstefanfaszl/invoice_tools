from exchangelib import Credentials, Account, DELEGATE

# Zugangsdaten anpassen
EMAIL = 'dein.email@deinedomain.de'
PASSWORD = 'deinpasswort'

# Verbindung aufbauen
creds = Credentials(username=EMAIL, password=PASSWORD)
account = Account(primary_smtp_address=EMAIL, credentials=creds, autodiscover=True, access_type=DELEGATE)

# Die letzten 10 E-Mails aus dem Posteingang ausgeben
for item in account.inbox.all().order_by('-datetime_received')[:10]:
    print(f'Subject: {item.subject}')
    print(f'From: {item.sender.email_address if item.sender else "Unknown"}')
    print(f'Received: {item.datetime_received}')
    print(f'Body: {item.text_body}\n')

# Hinweis: Die Mails bleiben auf dem Server, werden nicht gelöscht
