# invoice_tools

Python-Toolkit zur automatischen Extraktion strukturierter Daten aus PDF-Rechnungen und zum Abruf von Rechnungen aus E-Mail-Postfächern.

## Features

- **Extraktion** von Rechnungsfeldern (Lieferant, Rechnungsnummer, Beträge, Datum, IBAN, Zahlungsart) aus PDF-Dateien
- **Drei Extraktionswege** in Prioritätsreihenfolge: ZUGFeRD/Factur-X XML → KI-Analyse (Gemini/OpenAI/Claude) → Regex-Fallback
- **Automatischer KI-Provider-Test** beim Start — bei mehreren konfigurierten Providern wird der erste funktionierende verwendet
- **Postfach-Integration** für Exchange (EWS) und IMAP — lädt Rechnungsanhänge, extrahiert Felder und legt PDFs automatisch ab
- **BankingZV-Export** — übergibt Zahlungen, Lastschriften und Gutschriften als Erwartete Zahlungen an Banking4Windows
- **Ausgabeformate**: stdout, txt, pdf, csv, json, xml
- **Windows-Executable** über PyInstaller

## Verwendung

```
invoice_tools.exe extractor rechnung.pdf
invoice_tools.exe extractor -c invoice_extractor_config_RE.xml -f pdf -o bericht.pdf "*.pdf"
invoice_tools.exe inbox -m unread
invoice_tools.exe inbox -m dry
invoice_tools.exe inbox -m all -b export
```

## Batch-Skripte

| Skript | Funktion |
|--------|---------|
| `rechnungseingang_in.bat [dry\|unread\|all]` | Eingangsrechnungen aus Postfach laden |
| `rechnungseingang_ex.bat JJJJ/MM` | Eingangsrechnungen eines Monats als PDF-Bericht |
| `rechnungsausgang_ex.bat JJJJ/MM` | Ausgangsrechnungen eines Monats als PDF-Bericht |

## Konfiguration

| Datei | Funktion |
|-------|---------|
| `invoice_extractor_config_RE.xml` | Extractor-Config für Eingangsrechnungen |
| `invoice_extractor_config_RA.xml` | Extractor-Config für Ausgangsrechnungen |
| `invoice_inbox_config.xml` | Postfach, Ablagestruktur, BankingZV-Einstellungen |
| `invoice_tools_api_config.xml` | Zentrale KI-API-Konfiguration (Keys, Provider, Modell) |

Beispiel-Configs (ohne echte Keys) liegen als `*.example.xml` bei.

## KI-Konfiguration

Mehrere Provider können parallel konfiguriert werden — beim Start wird automatisch der erste funktionierende per Testanfrage ausgewählt:

```xml
<AI>
  <Provider>gemini</Provider>
  <Model>gemini-2.5-flash-lite</Model>
  <APIKey>AIza...</APIKey>
</AI>
<AI>
  <Provider>openai</Provider>
  <Model>gpt-4o-mini</Model>
  <APIKey>sk-proj-...</APIKey>
</AI>
```

## Entwicklung

```bash
# Executable bauen
build.bat

# Ins Produktionsverzeichnis veröffentlichen
publish.bat

# In GitHub speichern
github.bat

# Von GitHub wiederherstellen
github.bat restore

# Dokumentation neu generieren
py make_doku.py
```

### Abhängigkeiten

```
pip install pymupdf pillow reportlab pyinstaller
pip install exchangelib          # für Exchange-Postfächer
pip install anthropic            # für Claude
pip install openai               # für OpenAI
pip install google-genai         # für Gemini
```

## Dokumentation

Vollständige Dokumentation: [`invoice_tools_doku.pdf`](invoice_tools_doku.pdf)
