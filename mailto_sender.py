#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mailto_sender.py — Versendet PDF-Ausgangsrechnungen per E-Mail.

Empfängeradressen werden auf vier Wegen ermittelt:
  0 (Standard): Feste Test-Mailadresse aus der Konfigurationsdatei (<TestAdresse>)
  1: Nur aus PDF-Text (sucht nach "email:" im Rechnungstext)
  2: Nur aus MySQL-Datenbank (via konfigurierbaren SQL-Abfragen)
  3: PDF-Text versuchen, danach Datenbank (Kombination aus 1 und 2)

Konfiguration: invoice_tools_mailto_config.xml (oder per -c angeben)

Aufruf:
  python3 mailto_sender.py
  python3 mailto_sender.py -r 12345
  python3 mailto_sender.py --min-r-nr 100 --max-r-nr 200
  python3 mailto_sender.py --dry-run
  python3 mailto_sender.py --noconfirm
"""

import argparse
import os
import re
import smtplib
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime
from email import encoders as _email_encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from typing import List, Optional, Tuple

import fitz  # PyMuPDF

try:
    import mysql.connector
    _MYSQL_OK = True
except ImportError:
    _MYSQL_OK = False


# ── Exit-Codes ────────────────────────────────────────────────────────────────
_EXIT_OK          = 0   # Erfolgreich (inkl. Dry-Run / keine PDFs)
_EXIT_CONFIG      = 1   # Konfigurationsfehler
_EXIT_VERBINDUNG  = 2   # SMTP-Verbindungsfehler
_EXIT_TEILERFOLG  = 4   # Mind. 1 gesendet, mind. 1 fehlgeschlagen
_EXIT_ALLE_FEHL   = 5   # Keine einzige Mail gesendet


# ══════════════════════════════════════════════════════════════════════════════
# Konfigurationsdatenstrukturen
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class _SmtpKfg:
    host:      str
    port:      int
    ssl:       bool     # True  = SMTPS (Port 465)
    tls:       bool     # True  = STARTTLS (Port 587); False = plain SMTP (Port 25)
    user:      str
    password:  str
    absender:  str      # From-Adresse
    antwort_an: str = ''  # Reply-To (optional)


@dataclass
class _MailVorlageKfg:
    betreff_template: str
    html_body:        str
    plaintext_body:   str = ''


@dataclass
class _DbKfg:
    host:     str
    port:     int
    db_name:  str
    user:     str
    password: str
    sql1:     str
    sql2:     str = ''
    auftrags_nr_stellen: int = 0  # Anzahl Stellen der Auftragsnummer für PDF-Suche


@dataclass
class _MailtoKfg:
    smtp:             _SmtpKfg
    vorlage:          _MailVorlageKfg
    pdf_verzeichnis:  str
    rg_nr_regex:      str
    db:               Optional[_DbKfg]
    test_adresse:     str = ''      # Modus 0: feste Empfängeradresse aus Config
    email_text_marker: str = 'email:'
    auftrags_verzeichnis: str = ''  # Optionales Verzeichnis für Auftrags-PDFs


# ══════════════════════════════════════════════════════════════════════════════
# Hilfsfunktionen
# ══════════════════════════════════════════════════════════════════════════════

def _cfg_text(node: ET.Element, tag: str, default: str = '') -> str:
    """Liest den Textinhalt eines XML-Kindelements; gibt default zurück wenn nicht vorhanden."""
    el = node.find(tag)
    return (el.text or '').strip() if el is not None else default


def _normalize_rgnr(s: str) -> int:
    """Wandelt Rechnungsnummer-String in int um (entfernt alle Nicht-Ziffern)."""
    digits = re.sub(r'\D', '', s)
    return int(digits) if digits else -1


def _platzhalter(template: str, rgnr: int) -> str:
    """Ersetzt {rgnr}/{RechnungsNr} Platzhalter im Template."""
    return template.replace('{rgnr}', str(rgnr)).replace('{RechnungsNr}', str(rgnr))


def _parse_rg_datum(s: str) -> date:
    """Parst ein Datum in den Formaten TT.MM, TT.MM.JJ oder TT.MM.JJJJ.

    Bei fehlendem Jahr wird das aktuelle Jahr verwendet.
    Bei zweistelligem Jahr wird 2000 addiert.
    Gibt ValueError bei ungültigem Format oder Datum.
    """
    parts = s.strip().split('.')
    if len(parts) == 2:
        day, month = int(parts[0]), int(parts[1])
        year = date.today().year
    elif len(parts) == 3:
        day, month = int(parts[0]), int(parts[1])
        year = 2000 + int(parts[2]) if len(parts[2]) <= 2 else int(parts[2])
    else:
        raise ValueError(f'Ungültiges Datumsformat: "{s}" – erwartet TT.MM, TT.MM.JJ oder TT.MM.JJJJ')
    return date(year, month, day)


# ══════════════════════════════════════════════════════════════════════════════
# Konfiguration laden
# ══════════════════════════════════════════════════════════════════════════════

def _lade_kfg(config_path: str) -> _MailtoKfg:
    if not os.path.isfile(config_path):
        print(f'Fehler: Konfigurationsdatei nicht gefunden: {config_path}', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)
    try:
        root = ET.parse(config_path).getroot()
    except ET.ParseError as exc:
        print(f'Fehler beim Parsen der Konfigurationsdatei: {exc}', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)

    # ── SMTP ──────────────────────────────────────────────────────────────────
    smtp_node = root.find('SMTP')
    if smtp_node is None:
        print('Fehler: <SMTP>-Block fehlt in der Konfigurationsdatei.', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)
    _ssl  = _cfg_text(smtp_node, 'SSL', 'false').lower() == 'true'
    _port = int(_cfg_text(smtp_node, 'Port', '465' if _ssl else '587'))
    # TLS-Standard: true wenn nicht SSL und Port != 25, sonst false
    _tls_default = 'false' if (_ssl or _port == 25) else 'true'
    smtp = _SmtpKfg(
        host       = _cfg_text(smtp_node, 'Host'),
        port       = _port,
        ssl        = _ssl,
        tls        = _cfg_text(smtp_node, 'TLS', _tls_default).lower() == 'true',
        user       = _cfg_text(smtp_node, 'User'),
        password   = _cfg_text(smtp_node, 'Password'),
        absender   = _cfg_text(smtp_node, 'Absender'),
        antwort_an = _cfg_text(smtp_node, 'AntwortAn', ''),
    )
    if not smtp.host or not smtp.absender:
        print('Fehler: <SMTP> benötigt mindestens Host und Absender.', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)

    # ── PDF-Verzeichnis ───────────────────────────────────────────────────────
    pdf_verzeichnis = _cfg_text(root, 'PdfVerzeichnis')
    if not pdf_verzeichnis:
        print('Fehler: <PdfVerzeichnis> fehlt in der Konfigurationsdatei.', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)

    # ── Dateiname-Regex ───────────────────────────────────────────────────────
    rg_nr_regex = _cfg_text(root, 'RgNrRegex', r'(\d+)')

    # ── Mail-Vorlage ──────────────────────────────────────────────────────────
    vorlage_node = root.find('MailVorlage')
    if vorlage_node is None:
        print('Fehler: <MailVorlage>-Block fehlt in der Konfigurationsdatei.', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)
    betreff = _cfg_text(vorlage_node, 'Betreff', 'Rechnung Nr. {rgnr}')
    html_el = vorlage_node.find('HtmlBody')
    html_body = (html_el.text or '').strip() if html_el is not None else ''
    plaintext_body = _cfg_text(vorlage_node, 'PlaintextBody', '')
    if not html_body:
        print('Warnung: <HtmlBody> ist leer – Mails werden ohne Inhalt versendet.', file=sys.stderr)
    vorlage = _MailVorlageKfg(
        betreff_template = betreff,
        html_body        = html_body,
        plaintext_body   = plaintext_body,
    )

    # ── Test-Adresse (Modus 0) ────────────────────────────────────────────────
    test_adresse = _cfg_text(root, 'TestAdresse', '')

    # ── E-Mail-Marker im PDF-Text ─────────────────────────────────────────────
    email_text_marker = _cfg_text(root, 'EmailTextMarker', 'email:')

    # ── Datenbank (optional) ──────────────────────────────────────────────────
    db: Optional[_DbKfg] = None
    db_node = root.find('Datenbank')
    if db_node is not None:
        sql1 = _cfg_text(db_node, 'SQL1')
        sql2 = _cfg_text(db_node, 'SQL2', '')
        if sql1:
            db = _DbKfg(
                host                = _cfg_text(db_node, 'Host', 'localhost'),
                port                = int(_cfg_text(db_node, 'Port', '3306')),
                db_name             = _cfg_text(db_node, 'DbName'),
                user                = _cfg_text(db_node, 'User'),
                password            = _cfg_text(db_node, 'Password'),
                sql1                = sql1,
                sql2                = sql2,
                auftrags_nr_stellen = int(_cfg_text(db_node, 'AuftragNrStellen', '0')),
            )

    # ── Optionales Verzeichnis für Auftrags-PDFs ──────────────────────────────
    auftrags_verzeichnis = _cfg_text(root, 'PdfVerzeichnis_Auftraege', '')

    return _MailtoKfg(
        smtp                 = smtp,
        vorlage              = vorlage,
        pdf_verzeichnis      = pdf_verzeichnis,
        rg_nr_regex          = rg_nr_regex,
        db                   = db,
        test_adresse         = test_adresse,
        email_text_marker    = email_text_marker,
        auftrags_verzeichnis = auftrags_verzeichnis,
    )


# ══════════════════════════════════════════════════════════════════════════════
# PDF-Verzeichnis durchsuchen
# ══════════════════════════════════════════════════════════════════════════════

def _finde_pdfs(
    verzeichnis: str,
    rg_nr_regex: str,
    min_nr: Optional[int],
    max_nr: Optional[int],
    einzelne_nr: Optional[int],
    rg_datum: Optional[date],
    debug: int,
) -> List[Tuple[int, str]]:
    """Gibt Liste von (rgnr_int, dateipfad) zurück, nach Parametern gefiltert."""
    if not os.path.isdir(verzeichnis):
        print(f'Fehler: PDF-Verzeichnis nicht gefunden: {verzeichnis}', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)

    try:
        pattern = re.compile(rg_nr_regex, re.IGNORECASE)
    except re.error as exc:
        print(f'Fehler: Ungültiger Regex in <RgNrRegex>: {exc}', file=sys.stderr)
        sys.exit(_EXIT_CONFIG)

    ergebnis: List[Tuple[int, str]] = []

    for fname in sorted(os.listdir(verzeichnis)):
        if not fname.lower().endswith('.pdf'):
            continue
        m = pattern.search(fname)
        if not m:
            if debug >= 2:
                print(f'  Übersprungen (kein Regex-Treffer im Dateinamen): {fname}')
            continue

        # Rechnungsnummer: named group "rgnr" bevorzugen, sonst erste Capture-Group
        try:
            rgnr_str = m.group('rgnr')
        except IndexError:
            groups = m.groups()
            rgnr_str = groups[0] if groups else m.group(0)

        rgnr = _normalize_rgnr(rgnr_str)
        if rgnr < 0:
            if debug >= 1:
                print(f'  Übersprungen (keine gültige Rechnungsnummer in "{fname}")')
            continue

        # Filter anwenden
        if einzelne_nr is not None and rgnr != einzelne_nr:
            continue
        if min_nr is not None and rgnr < min_nr:
            continue
        if max_nr is not None and rgnr > max_nr:
            continue

        pfad = os.path.join(verzeichnis, fname)

        if rg_datum is not None:
            mtime = datetime.fromtimestamp(os.path.getmtime(pfad)).date()
            if mtime != rg_datum:
                if debug >= 2:
                    print(f'  Übersprungen (Dateidatum {mtime} != {rg_datum}): {fname}')
                continue

        ergebnis.append((rgnr, pfad))
        if debug >= 2:
            print(f'  Gefunden: {fname}  (RE-Nr {rgnr})')

    return ergebnis


# ══════════════════════════════════════════════════════════════════════════════
# Adressermittlung
# ══════════════════════════════════════════════════════════════════════════════

def _email_aus_pdf(pdf_pfad: str, marker: str, debug: int) -> Optional[str]:
    """Sucht nach dem konfigurierten Marker im PDF-Text und gibt die E-Mail zurück."""
    try:
        doc = fitz.open(pdf_pfad)
        text = ''.join(page.get_text() for page in doc)
        doc.close()
    except Exception as exc:
        if debug >= 1:
            print(f'  PDF-Lesefehler: {exc}')
        return None

    # Marker case-insensitiv suchen, danach E-Mail-Adresse
    escaped = re.escape(marker)
    m = re.search(escaped + r'\s*([^\s,;<>()\[\]]+@[^\s,;<>()\[\]]+)', text, re.IGNORECASE)
    if m:
        addr = m.group(1).strip().rstrip('.')
        if debug >= 1:
            print(f'  E-Mail aus PDF-Text: {addr}')
        return addr
    if debug >= 1:
        print(f'  Kein "{marker}" im PDF-Text gefunden.')
    return None


def _email_aus_db(
    rgnr: int, db_kfg: _DbKfg, debug: int
) -> Tuple[Optional[str], Optional[str]]:
    """Ermittelt E-Mail-Adresse (und ggf. Auftragsnummer) aus MySQL-Datenbank.

    Rückgabe: (email, auftragsnummer) — auftragsnummer ist None wenn keine 2. Spalte.
    Spalte 1 = E-Mail-Adresse, Spalte 2 (optional) = Auftragsnummer.
    """
    if not _MYSQL_OK:
        print('  Warnung: mysql-connector-python nicht installiert – '
              'Datenbankabfrage nicht möglich.', file=sys.stderr)
        return None, None

    try:
        conn = mysql.connector.connect(
            host     = db_kfg.host,
            port     = db_kfg.port,
            database = db_kfg.db_name,
            user     = db_kfg.user,
            password = db_kfg.password,
            use_pure = True,   # Reine Python-Implementierung, keine nativen DLLs nötig
        )
    except Exception as exc:
        print(f'  Datenbankverbindungsfehler: {exc}', file=sys.stderr)
        return None, None

    cursor = conn.cursor()
    try:
        sqls = [s for s in (db_kfg.sql1, db_kfg.sql2) if s]
        for i, sql_template in enumerate(sqls, start=1):
            sql = sql_template.replace('{rgnr}', str(rgnr))
            if debug >= 1:
                print(f'  DB-SQL{i}: {sql}')
            try:
                cursor.execute(sql)
                rows = cursor.fetchall()
            except Exception as exc:
                print(f'  DB-SQL{i}-Fehler: {exc}', file=sys.stderr)
                continue

            if not rows:
                if debug >= 1:
                    print(f'  DB-SQL{i}: kein Ergebnis.')
                continue

            row = rows[0]
            # Spalte 1 = E-Mail-Adresse
            email_val = str(row[0]).strip()
            # Spalte 2 (optional) = Auftragsnummer
            auftrag_val = str(row[1]).strip() if len(row) > 1 else None

            if email_val and '@' in email_val:
                if debug >= 1:
                    print(f'  E-Mail aus DB (SQL{i}): {email_val}')
                    if auftrag_val:
                        print(f'  Auftragsnummer aus DB (SQL{i}): {auftrag_val}')
                return email_val, auftrag_val
            if debug >= 1:
                print(f'  DB-SQL{i}: ungültige Adresse: {email_val!r}')
    finally:
        cursor.close()
        conn.close()

    return None, None


def _ermittle_adresse(
    rgnr: int,
    pdf_pfad: str,
    adress_modus: int,
    kfg: _MailtoKfg,
    debug: int,
) -> Tuple[Optional[str], Optional[str]]:
    """Ermittelt Empfängeradresse gemäß gewähltem Modus.

    0 = feste Test-Adresse aus Config (<TestAdresse>)
    1 = nur PDF-Text
    2 = nur Datenbank
    3 = PDF-Text, danach Datenbank als Fallback

    Rückgabe: (email, auftragsnummer) — auftragsnummer nur bei DB-Abfrage mit 2. Spalte.
    """
    if adress_modus == 0:
        if kfg.test_adresse:
            if debug >= 1:
                print(f'  Test-Adresse (Modus 0): {kfg.test_adresse}')
            return kfg.test_adresse, None
        print('  Fehler: Modus 0 gewählt, aber <TestAdresse> fehlt in der Konfiguration.',
              file=sys.stderr)
        return None, None

    if adress_modus in (1, 3):
        addr = _email_aus_pdf(pdf_pfad, kfg.email_text_marker, debug)
        if addr:
            return addr, None
        if adress_modus == 1:
            return None, None

    # Modus 2 oder Modus 3 (PDF erfolglos)
    if kfg.db is None:
        if debug >= 1:
            print('  Kein <Datenbank>-Block in der Konfiguration.')
        return None, None
    return _email_aus_db(rgnr, kfg.db, debug)


def _finde_auftrags_pdfs(
    verzeichnis: str,
    auftragsnummer: str,
    stellen: int,
    exclude_pfad: str,
    debug: int,
) -> List[str]:
    """Sucht PDFs im Verzeichnis deren Dateiname die letzten `stellen` Ziffern
    der Auftragsnummer enthält.

    Gibt eine Liste von Dateipfaden zurück (ohne exclude_pfad).
    """
    if stellen <= 0 or not auftragsnummer:
        return []

    # Nur Ziffern aus der Auftragsnummer extrahieren, dann letzten X nehmen
    digits = re.sub(r'\D', '', auftragsnummer)
    if not digits:
        return []
    suffix = digits[-stellen:]
    if debug >= 1:
        print(f'  Auftrags-PDFs: suche nach "{suffix}" '
              f'(letzte {stellen} Stellen von Auftrag {auftragsnummer})')

    ergebnis: List[str] = []
    exclude_norm = os.path.normcase(os.path.abspath(exclude_pfad))
    try:
        eintraege = sorted(os.listdir(verzeichnis))
    except OSError:
        return []

    for fname in eintraege:
        if not fname.lower().endswith('.pdf'):
            continue
        pfad = os.path.join(verzeichnis, fname)
        if os.path.normcase(os.path.abspath(pfad)) == exclude_norm:
            continue
        if suffix in fname:
            ergebnis.append(pfad)
            if debug >= 1:
                print(f'  Auftrags-PDF gefunden: {fname}')

    return ergebnis


# ══════════════════════════════════════════════════════════════════════════════
# Mail-Erstellung und Versand
# ══════════════════════════════════════════════════════════════════════════════

def _erstelle_mail(
    empfaenger: str,
    rgnr: int,
    pdf_pfade: List[str],
    kfg: _MailtoKfg,
) -> MIMEMultipart:
    """Baut die MIME-Mail auf. pdf_pfade[0] ist die Hauptrechnung; weitere sind Anlagen."""
    betreff   = _platzhalter(kfg.vorlage.betreff_template, rgnr)
    html_body = _platzhalter(kfg.vorlage.html_body, rgnr)
    plain     = _platzhalter(kfg.vorlage.plaintext_body, rgnr)

    if plain:
        # multipart/alternative: erst Plaintext, dann HTML
        alternative = MIMEMultipart('alternative')
        alternative.attach(MIMEText(plain, 'plain', 'utf-8'))
        alternative.attach(MIMEText(html_body, 'html', 'utf-8'))
        msg = MIMEMultipart('mixed')
        msg.attach(alternative)
    else:
        msg = MIMEMultipart('mixed')
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

    msg['From']    = kfg.smtp.absender
    msg['To']      = empfaenger
    msg['Subject'] = betreff
    msg['Date']    = formatdate(localtime=True)
    if kfg.smtp.antwort_an:
        msg['Reply-To'] = kfg.smtp.antwort_an

    # PDF-Anhänge
    for pdf_pfad in pdf_pfade:
        with open(pdf_pfad, 'rb') as f:
            pdf_data = f.read()
        fname = os.path.basename(pdf_pfad)
        part = MIMEBase('application', 'pdf')
        part.set_payload(pdf_data)
        _email_encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=fname)
        part.add_header('Content-Type', 'application/pdf', name=fname)
        msg.attach(part)

    return msg


def _sende_mail(
    msg: MIMEMultipart,
    empfaenger: str,
    smtp_kfg: _SmtpKfg,
    debug: int,
) -> bool:
    """Sendet die Mail über SMTP. Gibt True bei Erfolg, False bei Fehler zurück."""
    modus = 'SSL' if smtp_kfg.ssl else ('STARTTLS' if smtp_kfg.tls else 'plain')
    if debug >= 1:
        print(f'  SMTP: {smtp_kfg.host}:{smtp_kfg.port}  {modus}')
    try:
        if smtp_kfg.ssl:
            server = smtplib.SMTP_SSL(smtp_kfg.host, smtp_kfg.port)
        else:
            server = smtplib.SMTP(smtp_kfg.host, smtp_kfg.port)
            if smtp_kfg.tls:
                server.starttls()
        if smtp_kfg.user:
            server.login(smtp_kfg.user, smtp_kfg.password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as exc:
        print(f'  SMTP-Fehler: {exc}', file=sys.stderr)
        return False


# ══════════════════════════════════════════════════════════════════════════════
# Hauptfunktion
# ══════════════════════════════════════════════════════════════════════════════

def run(argv=None):
    parser = argparse.ArgumentParser(
        prog='invoice_tools mailto',
        description='Versendet PDF-Ausgangsrechnungen per E-Mail.',
    )
    parser.add_argument(
        '-c', '--config',
        default='invoice_tools_mailto_config.xml',
        help='Konfigurationsdatei (Standard: invoice_tools_mailto_config.xml)',
    )

    rgnr_group = parser.add_mutually_exclusive_group(required=True)
    rgnr_group.add_argument(
        '-r', '--rg-nr', type=int, default=None,
        help='Nur diese einzelne Rechnungsnummer versenden',
    )
    rgnr_group.add_argument(
        '--min-rg-nr', type=int, default=None,
        help='Kleinste zu versendende Rechnungsnummer',
    )

    parser.add_argument(
        '--max-rg-nr', type=int, default=None,
        help='Größte zu versendende Rechnungsnummer (nur zusammen mit --min-rg-nr)',
    )
    parser.add_argument(
        '-a', '--adress-modus', type=int, choices=[0, 1, 2, 3], default=0,
        help='Adressermittlung: 0=Test-Adresse aus Config (Standard), '
             '1=nur PDF-Text, 2=nur DB, 3=PDF+DB',
    )
    parser.add_argument(
        '--dry-run', action='store_true',
        help='Simulation: zeigt was gesendet würde, ohne tatsächlich zu senden',
    )
    parser.add_argument(
        '-d', '--debug', type=int, default=0, metavar='LEVEL',
        help='Debug-Level: 0=aus, 1=Überblick, 2=Details (Standard: 0)',
    )
    parser.add_argument(
        '--noconfirm', action='store_true',
        help='Mails sofort versenden ohne Bestätigungsdialog',
    )
    parser.add_argument(
        '--rg-datum', default=None, metavar='DATUM',
        help='Nur PDFs mit diesem Dateidatum versenden. '
             'Format: TT.MM  oder  TT.MM.JJ  oder  TT.MM.JJJJ',
    )

    args = parser.parse_args(argv)

    # --max-rg-nr ist nur zusammen mit --min-rg-nr sinnvoll
    if args.max_rg_nr is not None and args.rg_nr is not None:
        parser.error('--max-rg-nr kann nicht zusammen mit --rg-nr verwendet werden.')

    # --rg-datum parsen
    rg_datum: Optional[date] = None
    if args.rg_datum:
        try:
            rg_datum = _parse_rg_datum(args.rg_datum)
        except (ValueError, IndexError):
            parser.error(f'--rg-datum: ungültiges Datum "{args.rg_datum}". '
                         'Erwartet: TT.MM, TT.MM.JJ oder TT.MM.JJJJ')

    # Konfiguration laden
    config_path = args.config
    if not os.path.isabs(config_path):
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_path)
    kfg = _lade_kfg(config_path)

    # PDFs suchen
    pdfs = _finde_pdfs(
        verzeichnis  = kfg.pdf_verzeichnis,
        rg_nr_regex  = kfg.rg_nr_regex,
        min_nr       = args.min_rg_nr,
        max_nr       = args.max_rg_nr,
        einzelne_nr  = args.rg_nr,
        rg_datum     = rg_datum,
        debug        = args.debug,
    )

    if not pdfs:
        datum_info = f' (Datum: {rg_datum.strftime("%d.%m.%Y")})' if rg_datum else ''
        print(f'Keine passenden PDF-Dateien gefunden{datum_info}.')
        sys.exit(_EXIT_OK)

    print(f'{len(pdfs)} PDF-Datei(en) gefunden.'
          + (f'  [Datum: {rg_datum.strftime("%d.%m.%Y")}]' if rg_datum else ''))
    if args.dry_run:
        print('[DRY-RUN] Es werden keine Mails tatsächlich versendet.\n')

    # Für jede Rechnung E-Mail-Adresse ermitteln
    # to_send: (rgnr, [pdf_pfad, ...], empfaenger)
    to_send: List[Tuple[int, List[str], str]] = []
    keine_adresse: List[Tuple[int, str]] = []

    for rgnr, pdf_pfad in pdfs:
        fname = os.path.basename(pdf_pfad)
        print(f'\nRE-Nr {rgnr}: {fname}')
        addr, auftragsnr = _ermittle_adresse(rgnr, pdf_pfad, args.adress_modus, kfg, args.debug)
        if not addr:
            print('  WARNUNG: Keine E-Mail-Adresse ermittelbar – übersprungen.')
            keine_adresse.append((rgnr, fname))
            continue

        # Zusätzliche Auftrags-PDFs suchen wenn DB eine Auftragsnummer geliefert hat
        pdf_pfade = [pdf_pfad]
        if auftragsnr and kfg.db and kfg.db.auftrags_nr_stellen > 0:
            such_verzeichnis = kfg.auftrags_verzeichnis or kfg.pdf_verzeichnis
            extra = _finde_auftrags_pdfs(
                verzeichnis    = such_verzeichnis,
                auftragsnummer = auftragsnr,
                stellen        = kfg.db.auftrags_nr_stellen,
                exclude_pfad   = pdf_pfad,
                debug          = args.debug,
            )
            pdf_pfade.extend(extra)

        to_send.append((rgnr, pdf_pfade, addr))
        print(f'  -> Empfänger: {addr}')
        if len(pdf_pfade) > 1:
            print(f'  -> Anhänge:   {len(pdf_pfade)} PDFs '
                  f'({", ".join(os.path.basename(p) for p in pdf_pfade)})')

    if keine_adresse:
        print(f'\n{len(keine_adresse)} Rechnung(en) ohne Empfängeradresse übersprungen:')
        for rgnr, fname in keine_adresse:
            print(f'  RE-Nr {rgnr}: {fname}')

    if not to_send:
        print('\nKeine Mails zu versenden.')
        sys.exit(_EXIT_OK)

    # Übersicht der geplanten Sendungen
    print(f'\n{"─"*60}')
    print(f'Geplante Sendungen ({len(to_send)}):')
    for rgnr, pdf_pfade, addr in to_send:
        betreff = _platzhalter(kfg.vorlage.betreff_template, rgnr)
        anhang_info = (f'{len(pdf_pfade)} Anhänge' if len(pdf_pfade) > 1
                       else os.path.basename(pdf_pfade[0]))
        print(f'  RE-Nr {rgnr:>6}  ->  {addr:<40}  [{anhang_info}]')
        if args.debug >= 1:
            print(f'             Betreff: {betreff}')
            if len(pdf_pfade) > 1:
                for p in pdf_pfade:
                    print(f'             Anhang:  {os.path.basename(p)}')
    print(f'{"─"*60}')

    # Bestätigung einholen (außer bei --noconfirm oder --dry-run)
    if not args.dry_run and not args.noconfirm:
        print(f'\n{len(to_send)} Mail(s) versenden? [j/N] ', end='', flush=True)
        try:
            antwort = input().strip().lower()
        except (EOFError, KeyboardInterrupt):
            print('\nAbgebrochen.')
            sys.exit(_EXIT_OK)
        if antwort not in ('j', 'ja', 'y', 'yes'):
            print('Abgebrochen.')
            sys.exit(_EXIT_OK)

    # Mails senden
    ok = 0
    fehler = 0
    for rgnr, pdf_pfade, addr in to_send:
        fname = os.path.basename(pdf_pfade[0])
        print(f'\nSende RE-Nr {rgnr} ({fname}) -> {addr} ...')

        if args.dry_run:
            betreff = _platzhalter(kfg.vorlage.betreff_template, rgnr)
            print(f'  [DRY-RUN] Betreff: {betreff}')
            for p in pdf_pfade:
                print(f'  [DRY-RUN] Anhang:  {os.path.basename(p)}')
            ok += 1
            continue

        try:
            msg = _erstelle_mail(addr, rgnr, pdf_pfade, kfg)
        except OSError as exc:
            print(f'  Fehler beim Lesen der PDF-Datei: {exc}', file=sys.stderr)
            fehler += 1
            continue

        if _sende_mail(msg, addr, kfg.smtp, args.debug):
            print('  OK')
            ok += 1
        else:
            fehler += 1

    # Abschlusszusammenfassung
    print(f'\n{"═"*60}')
    if args.dry_run:
        print(f'[DRY-RUN] {ok} Mail(s) würden versendet.')
        if keine_adresse:
            print(f'          {len(keine_adresse)} Rechnung(en) ohne Empfängeradresse.')
    else:
        print(f'Ergebnis: {ok} gesendet, {fehler} Fehler'
              + (f', {len(keine_adresse)} ohne Adresse' if keine_adresse else '') + '.')

    if fehler > 0 and ok == 0:
        sys.exit(_EXIT_ALLE_FEHL)
    if fehler > 0:
        sys.exit(_EXIT_TEILERFOLG)
    sys.exit(_EXIT_OK)


if __name__ == '__main__':
    run()
