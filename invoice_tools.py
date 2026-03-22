#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
invoice_tools — Kombiniertes Einstiegs-Skript für Rechnungsverarbeitung.

Tools:
  extractor   Extrahiert Rechnungsfelder aus PDF-Dateien
  inbox       Verarbeitet PDF-Rechnungsanhänge aus Exchange-/IMAP-Postfächern

Verwendung:
  invoice_tools extractor rechnung.pdf
  invoice_tools extractor -f csv "/pfad/*.pdf"
  invoice_tools inbox -m unread
  invoice_tools inbox -m dry -c andere_config.xml

Hilfe zum jeweiligen Tool:
  invoice_tools extractor --help
  invoice_tools inbox --help
"""

import sys


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ('-h', '--help'):
        print(__doc__)
        sys.exit(0)

    tool = sys.argv[1]
    argv = sys.argv[2:]

    if tool == 'extractor':
        from invoice_extractor import run
        run(argv)
    elif tool == 'inbox':
        from inbox_processor import run
        run(argv)
    else:
        print(f'Fehler: Unbekanntes Tool "{tool}". Erlaubt: extractor, inbox\n',
              file=sys.stderr)
        print(__doc__)
        sys.exit(1)


if __name__ == '__main__':
    main()
