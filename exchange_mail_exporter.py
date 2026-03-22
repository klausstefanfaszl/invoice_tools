import os
import argparse
import urllib3
from exchangelib import Credentials, Account, DELEGATE, Configuration
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

def sanitize_filename(s):
    return "".join(c if c.isalnum() or c in (' ', '.', '_', '-') else '_' for c in s)

def main():
    parser = argparse.ArgumentParser(description="Exchange Mail Exporter")
    parser.add_argument('--email', required=True, help="Email-Adresse für Exchange Login")
    parser.add_argument('--password', required=True, help="Passwort für Exchange Login")
    parser.add_argument('--output-dir', default='./exported-mails', help="Ordner zum Speichern der Mails und Anhänge")
    parser.add_argument('--limit', type=int, default=10, help="Anzahl der zuletzt abzurufenden Mails")
    parser.add_argument('--server', default=None, help="Exchange Server Adresse, falls Autodiscover nicht genutzt wird")
    parser.add_argument('--test-modus', action='store_true', help="Simuliert den Export ohne Dateien zu schreiben")
    parser.add_argument('--debug', action='store_true', help="Aktiviert die Protokollierung der Aktionen")

    args = parser.parse_args()

    # Selbstsignierte Zertifikate erlauben
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

    debug = args.debug or args.test_modus
    if debug:
        print(f"Starte mit Parametern: {args}")

    if not args.test_modus:
        os.makedirs(args.output_dir, exist_ok=True)

    creds = Credentials(username=args.email, password=args.password)

    if args.server:
        config = Configuration(server=args.server, credentials=creds)
        account = Account(primary_smtp_address=args.email,
                          credentials=creds,
                          config=config,
                          autodiscover=False,
                          access_type=DELEGATE)
    else:
        account = Account(primary_smtp_address=args.email,
                          credentials=creds,
                          autodiscover=True,
                          access_type=DELEGATE)

    for item in account.inbox.all().order_by('-datetime_received')[:args.limit]:
        timestamp = item.datetime_received.strftime('%Y%m%d%H%M%S') if item.datetime_received else 'unknown'
        base_filename = os.path.join(args.output_dir, f"{timestamp}")

        html_body = str(item.body) if item.body else ''

        mail_path = base_filename + '.html'

        if debug:
            print(f"Verarbeite Mail '{item.subject}' empfangen: {item.datetime_received}")
            print(f"Zielpfad für Mail: {mail_path}")

        if not args.test_modus:
            with open(mail_path, 'w', encoding='utf-8') as f:
                f.write(html_body)

        if debug:
            print(f"Mail '{item.subject}' gespeichert als {mail_path}")

        for attachment in item.attachments:
            if hasattr(attachment, 'content') and hasattr(attachment, 'name') and attachment.name:
                attachment_name = sanitize_filename(attachment.name)
                attachment_path = f"{base_filename}_{attachment_name}"

                if debug:
                    print(f"Verarbeite Anhang '{attachment_name}' für Mail '{item.subject}'")
                    print(f"Zielpfad für Anhang: {attachment_path}")

                if not args.test_modus:
                    with open(attachment_path, 'wb') as f:
                        f.write(attachment.content)

                if debug:
                    print(f"Anhang '{attachment_name}' gespeichert als {attachment_path}")

if __name__ == "__main__":
    main()
