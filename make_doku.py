"""
make_doku.py — Generiert invoice_tools_doku.pdf mit reportlab Platypus.
Aufruf: py make_doku.py
"""

import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import Flowable

# ---------------------------------------------------------------------------
# Farben
# ---------------------------------------------------------------------------
DARK_BLUE   = colors.HexColor("#1a3a5c")
LIGHT_BLUE  = colors.HexColor("#f0f4f8")
CODE_BG     = colors.HexColor("#f5f5f5")
CODE_BORDER = colors.HexColor("#cccccc")
GREY_TEXT   = colors.HexColor("#666666")
WHITE       = colors.white
ROW_ALT     = colors.HexColor("#f0f4f8")

# ---------------------------------------------------------------------------
# Ausgabepfad
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PDF  = os.path.join(SCRIPT_DIR, "invoice_tools_doku.pdf")

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
def make_styles():
    styles = {}

    styles["title"] = ParagraphStyle(
        "title",
        fontName="Helvetica-Bold",
        fontSize=28,
        leading=34,
        alignment=TA_CENTER,
        textColor=DARK_BLUE,
        spaceAfter=8,
    )
    styles["subtitle"] = ParagraphStyle(
        "subtitle",
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        alignment=TA_CENTER,
        textColor=GREY_TEXT,
        spaceAfter=18,
    )
    styles["h1"] = ParagraphStyle(
        "h1",
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=17,
        textColor=DARK_BLUE,
        spaceBefore=14,
        spaceAfter=5,
    )
    styles["h2"] = ParagraphStyle(
        "h2",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=15,
        textColor=DARK_BLUE,
        spaceBefore=10,
        spaceAfter=4,
    )
    styles["body"] = ParagraphStyle(
        "body",
        fontName="Helvetica",
        fontSize=9,
        leading=13,
        spaceAfter=4,
    )
    styles["note"] = ParagraphStyle(
        "note",
        fontName="Helvetica-Oblique",
        fontSize=9,
        leading=13,
        textColor=GREY_TEXT,
        spaceAfter=4,
    )
    styles["code"] = ParagraphStyle(
        "code",
        fontName="Courier",
        fontSize=8,
        leading=11,
        leftIndent=6,
        rightIndent=6,
    )
    styles["th"] = ParagraphStyle(
        "th",
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        textColor=WHITE,
        alignment=TA_LEFT,
    )
    styles["td"] = ParagraphStyle(
        "td",
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
    )
    styles["td_code"] = ParagraphStyle(
        "td_code",
        fontName="Courier",
        fontSize=8,
        leading=11,
        alignment=TA_LEFT,
    )
    return styles


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def h1(text, styles):
    return KeepTogether([
        Paragraph(text, styles["h1"]),
        HRFlowable(width="100%", thickness=1, color=DARK_BLUE, spaceAfter=4),
    ])


def h2(text, styles):
    return Paragraph(text, styles["h2"])


def body(text, styles):
    return Paragraph(text, styles["body"])


def note(text, styles):
    return Paragraph(text, styles["note"])


def sp(height=4):
    return Spacer(1, height)


def make_table(header_row, data_rows, styles, col_widths=None, td_style="td"):
    """Erzeugt eine Platypus-Table mit farbigem Header und abwechselnden Zeilen."""
    th_s = styles["th"]
    td_s = styles[td_style]

    all_rows = [[Paragraph(str(c), th_s) for c in header_row]]
    for row in data_rows:
        all_rows.append([Paragraph(str(c), td_s) for c in row])

    table = Table(all_rows, colWidths=col_widths, hAlign="LEFT",
                  repeatRows=1)

    ts = TableStyle([
        # Header
        ("BACKGROUND",  (0, 0), (-1, 0), DARK_BLUE),
        ("TEXTCOLOR",   (0, 0), (-1, 0), WHITE),
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0), 9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING",    (0, 0), (-1, 0), 5),
        # Datenzeilen
        ("FONTNAME",    (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 1), (-1, -1), 9),
        ("TOPPADDING",  (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        # Rahmen
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
    ])
    # Abwechselnde Zeilen einfärben
    for i in range(1, len(all_rows)):
        if i % 2 == 0:
            ts.add("BACKGROUND", (0, i), (-1, i), ROW_ALT)
        else:
            ts.add("BACKGROUND", (0, i), (-1, i), WHITE)

    table.setStyle(ts)
    return table


class CodeBlock(Flowable):
    """Ein grau hinterlegter, umrandeter Code-Block."""

    def __init__(self, lines, width, styles):
        super().__init__()
        self._lines = lines
        self._width = width
        self._style = styles["code"]
        self._pad = 6

    def wrap(self, availWidth, availHeight):
        self._used_width = min(self._width, availWidth)
        line_h = self._style.leading
        self._height = line_h * len(self._lines) + 2 * self._pad
        return self._used_width, self._height

    def draw(self):
        c = self.canv
        w, h = self._used_width, self._height
        # Hintergrund
        c.setFillColor(CODE_BG)
        c.rect(0, 0, w, h, fill=1, stroke=0)
        # Rahmen
        c.setStrokeColor(CODE_BORDER)
        c.setLineWidth(0.5)
        c.rect(0, 0, w, h, fill=0, stroke=1)
        # Text
        c.setFillColor(colors.black)
        c.setFont(self._style.fontName, self._style.fontSize)
        line_h = self._style.leading
        y = h - self._pad - self._style.fontSize
        for line in self._lines:
            c.drawString(self._pad, y, line)
            y -= line_h


def code_block(text, width, styles):
    lines = text.split("\n")
    return CodeBlock(lines, width, styles)


# ---------------------------------------------------------------------------
# Seitennummern
# ---------------------------------------------------------------------------
def add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(GREY_TEXT)
    canvas.drawRightString(
        A4[0] - 2 * cm,
        1.2 * cm,
        f"Seite {doc.page}",
    )
    canvas.restoreState()


# ---------------------------------------------------------------------------
# Inhalt aufbauen
# ---------------------------------------------------------------------------
def build_content(styles, usable_width):
    story = []
    W = usable_width  # Gesamtbreite für Tabellen / Code

    # ------------------------------------------------------------------
    # Titelseite / Header-Block
    # ------------------------------------------------------------------
    story.append(Paragraph("invoice_tools", styles["title"]))
    story.append(Paragraph(
        "Technische Dokumentation · Version 1.0 · UHDE Datentechnik",
        styles["subtitle"],
    ))
    story.append(HRFlowable(width="100%", thickness=1.5, color=DARK_BLUE, spaceAfter=12))

    # ------------------------------------------------------------------
    # 1. Überblick
    # ------------------------------------------------------------------
    story.append(h1("1. Überblick", styles))
    story.append(body(
        "<b>invoice_tools.exe</b> ist ein kombiniertes Kommandozeilenprogramm mit zwei integrierten Tools:",
        styles,
    ))

    tbl = make_table(
        ["Tool", "Funktion"],
        [
            ["extractor", "Felder aus PDFs extrahieren (Lieferant, Beträge, Datum, IBAN …)"],
            ["inbox",     "Rechnungsanhänge automatisch aus E-Mail-Postfach laden und ablegen"],
        ],
        styles,
        col_widths=[W * 0.18, W * 0.82],
    )
    story.append(tbl)
    story.append(sp(6))

    story.append(body("<b>Grundsyntax:</b>", styles))
    story.append(code_block(
        "invoice_tools.exe extractor [Optionen] PDF [PDF ...]\n"
        "invoice_tools.exe inbox -m MODUS [Optionen]",
        W, styles,
    ))
    story.append(sp(4))
    story.append(body(
        "Beide Tools teilen das <i>_internal</i>-Verzeichnis. "
        "Für den Betrieb sind keine weiteren Installationen erforderlich.",
        styles,
    ))

    # ------------------------------------------------------------------
    # 2. Systemvoraussetzungen
    # ------------------------------------------------------------------
    story.append(h1("2. Systemvoraussetzungen", styles))

    tbl = make_table(
        ["Komponente", "Erforderlich", "Hinweis"],
        [
            ["invoice_tools.exe + _internal/", "Ja", "Im ZIP-Paket enthalten"],
            ["Config-XML-Dateien",              "Ja", "Im gleichen Verzeichnis wie die EXE ablegen"],
            ["Internetzugang / API-Key",        "Nein", "Nur für KI-Modus (Claude / OpenAI / Gemini)"],
            ["Tesseract OCR",                   "Nein", "Nur für gescannte PDFs ohne KI-Modus"],
            ["Exchange / IMAP-Zugangsdaten",    "Für inbox", "In invoice_inbox_config.xml konfigurieren"],
        ],
        styles,
        col_widths=[W * 0.38, W * 0.16, W * 0.46],
    )
    story.append(tbl)

    # ------------------------------------------------------------------
    # 3. Tool: extractor
    # ------------------------------------------------------------------
    story.append(h1("3. Tool: extractor", styles))
    story.append(body(
        "Extrahiert strukturierte Daten (Lieferant, Rechnungsnummer, Beträge, Datum, IBAN, Zahlungsart) "
        "aus PDF-Rechnungen. Unterstützt ZUGFeRD/Factur-X, KI-Analyse und Regex-Fallback.",
        styles,
    ))

    # 3.1 Extraktionsprioritäten
    story.append(h2("3.1 Extraktionsprioritäten", styles))
    tbl = make_table(
        ["Priorität", "Methode", "Voraussetzung"],
        [
            ["1 (höchste)", "ZUGFeRD/Factur-X XML",
             "Eingebettetes XML im PDF vorhanden"],
            ["2",           "KI-Analyse",
             "API-Key in der Config konfiguriert"],
            ["3 (Fallback)", "Regex-Extraktion",
             "Durchsuchbarer Text im PDF"],
            ["–",           "OCR (Tesseract)",
             "Tesseract installiert + gescanntes PDF"],
        ],
        styles,
        col_widths=[W * 0.15, W * 0.28, W * 0.57],
    )
    story.append(tbl)

    # 3.2 Parameter
    story.append(h2("3.2 Parameter", styles))
    tbl = make_table(
        ["Parameter", "Kurz", "Standard", "Beschreibung"],
        [
            ["--config DATEI", "-c", "invoice_extractor_config.xml",
             "XML-Konfigurationsdatei"],
            ["--format FORMAT", "-f", "stdout",
             "stdout · txt · pdf · csv · json · xml"],
            ["--output DATEI", "-o", "(automatisch)",
             'Ausgabedatei. "-o STDOUT" leitet auf die Konsole um.'],
            ["--logfile DATEI", "-l", "(kein Log)",
             "Verarbeitungsprotokoll in Datei schreiben"],
            ["--debug LEVEL", "-d", "0",
             "0=aus · 1=Pfad · 2=Details · 3=Vollausgabe"],
            ["--api DATEI", "-a", "(automatisch)",
             "Zentrale KI-API-Konfigurationsdatei"],
        ],
        styles,
        col_widths=[W * 0.22, W * 0.08, W * 0.25, W * 0.45],
    )
    story.append(tbl)

    # 3.3 Ausgabeformate
    story.append(h2("3.3 Ausgabeformate", styles))
    tbl = make_table(
        ["Format", "Beschreibung"],
        [
            ["stdout", "Tabellarische Ausgabe direkt auf die Konsole (Standard)"],
            ["txt",    "Identisch mit stdout, in Datei gespeichert"],
            ["pdf",    "Formatierte PDF-Tabelle mit Spaltenbezeichnungen und optionaler Summenzeile"],
            ["csv",    "Kommagetrennte Werte, geeignet für Excel-Import"],
            ["json",   "Maschinenlesbare JSON-Ausgabe"],
            ["xml",    "XML-Ausgabe"],
        ],
        styles,
        col_widths=[W * 0.15, W * 0.85],
    )
    story.append(tbl)

    # 3.4 Beispielaufrufe
    story.append(h2("3.4 Beispielaufrufe", styles))

    story.append(body("Einzelne Rechnung auf Konsole:", styles))
    story.append(code_block("invoice_tools.exe extractor rechnung.pdf", W, styles))
    story.append(sp(4))

    story.append(body("Alle PDFs eines Monats als PDF-Bericht mit Log:", styles))
    story.append(code_block(
        'invoice_tools.exe extractor -c invoice_extractor_config_RA.xml \\\n'
        '    -f pdf -o Rechnungsausgang_01.pdf -l log.log \\\n'
        '    "C:\\Buchhaltung\\2026\\01\\*.pdf"',
        W, styles,
    ))
    story.append(sp(4))

    story.append(body("CSV-Ausgabe für Excel:", styles))
    story.append(code_block(
        'invoice_tools.exe extractor -f csv -o ausgabe.csv "*.pdf"',
        W, styles,
    ))

    # ------------------------------------------------------------------
    # 4. Tool: inbox
    # ------------------------------------------------------------------
    story.append(h1("4. Tool: inbox", styles))
    story.append(body(
        "Verbindet sich mit einem Exchange- oder IMAP-Postfach, lädt ungelesene E-Mails "
        "mit PDF-Anhängen, extrahiert Rechnungsfelder und speichert die PDFs automatisch "
        "in der konfigurierten Verzeichnisstruktur. Nicht-Rechnungen werden übersprungen.",
        styles,
    ))

    # 4.1 Parameter
    story.append(h2("4.1 Parameter", styles))
    tbl = make_table(
        ["Parameter", "Kurz", "Standard", "Beschreibung"],
        [
            ["--modus MODUS", "-m", "(Pflicht)",
             "dry=Simulation · unread=nur ungelesene Mails · all=alle Mails"],
            ["--config DATEI", "-c", "invoice_inbox_config.xml",
             "XML-Konfigurationsdatei"],
            ["--debug LEVEL", "-d", "0",
             "0=aus · 1=Pfad · 2=Details · 3=Vollausgabe"],
            ["--dry-run", "–", "–",
             "Simulation kombinierbar mit -m all oder -m unread"],
            ["--bzv MODUS", "-b", "–",
             "BankingZV-Export: dry=Anzeige · json=+JSON-Datei · export=+BankingZV-Aufruf"],
            ["--api DATEI", "-a", "(automatisch)",
             "Zentrale KI-API-Konfigurationsdatei"],
        ],
        styles,
        col_widths=[W * 0.22, W * 0.08, W * 0.20, W * 0.50],
    )
    story.append(tbl)

    # 4.2 Modi
    story.append(h2("4.2 Modi", styles))
    tbl = make_table(
        ["Modus", "Dateien speichern", "Als gelesen markieren", "Welche Mails"],
        [
            ["dry",    "Nein", "Nein", "Nur ungelesene — zeigt was gespeichert würde"],
            ["unread", "Ja",   "Ja",   "Nur ungelesene Mails"],
            ["all",    "Ja",   "Ja",   "Alle Mails (gelesen + ungelesen)"],
        ],
        styles,
        col_widths=[W * 0.13, W * 0.20, W * 0.23, W * 0.44],
    )
    story.append(tbl)

    # 4.3 Unterstützte Postfach-Typen
    story.append(h2("4.3 Unterstützte Postfach-Typen", styles))
    tbl = make_table(
        ["Typ", "Protokoll", "Typische Anbieter", "Pflichtfelder"],
        [
            ["exchange", "EWS (HTTPS)", "Microsoft Exchange, Office 365, On-Premise",
             "Email, Password"],
            ["imap", "IMAP4 SSL", "Gmail, GMX, T-Online, Outlook.com",
             "Email, Password, Server"],
        ],
        styles,
        col_widths=[W * 0.14, W * 0.16, W * 0.44, W * 0.26],
    )
    story.append(tbl)
    story.append(sp(4))
    story.append(note(
        "<b>Hinweis:</b> Gmail: Bei aktivierter 2-Faktor-Authentifizierung ist "
        "ein App-Passwort erforderlich.",
        styles,
    ))

    # 4.4 Dateiablage
    story.append(h2("4.4 Dateiablage", styles))
    story.append(body(
        "PDFs werden automatisch in der folgenden Struktur abgelegt:",
        styles,
    ))
    story.append(code_block(
        "<BaseDir>/<YYYY>/<MM>/ER_MMTT_<Lieferant>_<RechnungsNr>.pdf",
        W, styles,
    ))
    story.append(sp(4))
    story.append(body(
        "Rechnungen ohne erkennbares Datum landen in <i>&lt;BaseDir&gt;/_unbekannt/</i>. "
        "Vorhandene Dateien werden nicht überschrieben — bei Namenskollision wird ein "
        "Suffix angehängt (_2, _3, …).",
        styles,
    ))

    # 4.5 Filter
    story.append(h2("4.5 Filter", styles))
    tbl = make_table(
        ["Filter", "Config-Element", "Funktion"],
        [
            ["Dateinamen-Filter",
             "<AttachmentFilter><SkipPattern>",
             'Anhänge überspringen, deren Dateiname das Muster enthält (z.B. "summary", "lastschrift")'],
            ["Pflichtfeld-Filter",
             "<InvoiceFilter><RequiredField>",
             "PDF wird nur gespeichert, wenn alle Pflichtfelder extrahiert wurden "
             "(z.B. InvoiceNumber, GrossAmount)"],
        ],
        styles,
        col_widths=[W * 0.22, W * 0.30, W * 0.48],
    )
    story.append(tbl)

    # 4.6 Beispielaufrufe
    story.append(h2("4.6 Beispielaufrufe", styles))
    story.append(code_block(
        "invoice_tools.exe inbox -m dry\n"
        "invoice_tools.exe inbox -m unread -c invoice_inbox_config.xml\n"
        "invoice_tools.exe inbox -m all",
        W, styles,
    ))

    # ------------------------------------------------------------------
    # 5. Konfigurationsdateien
    # ------------------------------------------------------------------
    story.append(h1("5. Konfigurationsdateien", styles))
    tbl = make_table(
        ["Datei", "Verwendung"],
        [
            ["invoice_extractor_config.xml",
             "extractor: Vollversion mit allen Feldern (Standard)"],
            ["invoice_extractor_config_RE.xml",
             "extractor: Eingangsrechnungen — Feld SupplierName, berechneter MwSt.-Satz"],
            ["invoice_extractor_config_RA.xml",
             "extractor: Ausgangsrechnungen — Feld RecipientName, Summenzeile aktiv"],
            ["invoice_inbox_config.xml",
             "inbox: Exchange-/IMAP-Zugangsdaten, Ablagestruktur, Filter"],
            ["invoice_tools_api_config.xml",
             "Zentrale KI-API-Konfiguration (alle Provider / Keys)"],
        ],
        styles,
        col_widths=[W * 0.38, W * 0.62],
    )
    story.append(tbl)

    # 5.1 KI-Konfiguration
    story.append(h2("5.1 KI-Konfiguration (invoice_tools_api_config.xml)", styles))
    story.append(body(
        "Zentrale Konfigurationsdatei für alle KI-Provider. Mehrere &lt;AI&gt;-Blöcke möglich — "
        "beim Start wird automatisch der erste funktionierende Provider per Testanfrage ausgewählt.",
        styles,
    ))
    story.append(code_block(
        "<AI>\n"
        "  <Provider>gemini</Provider>  <!-- claude | openai | gemini -->\n"
        "  <Model>gemini-2.5-flash-lite</Model>\n"
        "  <APIKey>AIza...</APIKey>\n"
        "</AI>",
        W, styles,
    ))

    # 5.2 Postfach-Konfiguration
    story.append(h2("5.2 Postfach-Konfiguration (invoice_inbox_config.xml)", styles))
    story.append(body(
        "Der Postfachtyp (exchange oder imap) wird über das Attribut am &lt;Mailbox&gt;-Tag gesetzt:",
        styles,
    ))
    story.append(code_block(
        '<Mailbox type="exchange">\n'
        "  <Email>buchhaltung@firma.de</Email>\n"
        "  <Password>...</Password>\n"
        "  <Server>mail.firma.de</Server>\n"
        "  <MarkAsRead>true</MarkAsRead>\n"
        "  <Limit>100</Limit>\n"
        "</Mailbox>",
        W, styles,
    ))

    # ------------------------------------------------------------------
    # 6. Batch-Skripte
    # ------------------------------------------------------------------
    story.append(h1("6. Batch-Skripte", styles))
    tbl = make_table(
        ["Skript", "Funktion", "Aufruf"],
        [
            ["rechnungseingang_in.bat",
             "Eingangsrechnungen aus Postfach laden und speichern",
             "rechnungseingang_in.bat [modus]"],
            ["rechnungseingang_ex.bat",
             "Eingangsrechnungen eines Monats als PDF-Bericht",
             "rechnungseingang_ex.bat 2026/01"],
            ["rechnungsausgang_ex.bat",
             "Ausgangsrechnungen eines Monats als PDF-Bericht",
             "rechnungsausgang_ex.bat 2026/01"],
        ],
        styles,
        col_widths=[W * 0.30, W * 0.42, W * 0.28],
    )
    story.append(tbl)
    story.append(sp(4))
    story.append(body(
        "Mögliche Modi für <b>rechnungseingang_in.bat</b>: "
        "(Standard: <i>unread</i>) · <i>dry</i> (Simulation) · <i>all</i> (alle Mails)",
        styles,
    ))

    # ------------------------------------------------------------------
    # 7. Exit-Codes
    # ------------------------------------------------------------------
    story.append(h1("7. Exit-Codes", styles))
    tbl = make_table(
        ["Code", "Bedeutung"],
        [
            ["0", "Erfolgreich abgeschlossen"],
            ["1", "Keine verwertbaren Rechnungen / Verbindungsfehler / Konfigurationsfehler"],
            ["2", "KI-API nicht verfügbar (alle konfigurierten Provider fehlgeschlagen)"],
        ],
        styles,
        col_widths=[W * 0.10, W * 0.90],
    )
    story.append(tbl)

    # ------------------------------------------------------------------
    # 8. IBAN-Validierung
    # ------------------------------------------------------------------
    story.append(h1("8. IBAN-Validierung", styles))
    story.append(body(
        "Alle extrahierten IBANs werden automatisch nach ISO 13616 (Modulo-97-Verfahren) validiert. "
        "Ungültige oder maskierte IBANs (mit * oder XXX) werden stillschweigend verworfen.",
        styles,
    ))

    # ------------------------------------------------------------------
    # 9. Weiterentwicklung
    # ------------------------------------------------------------------
    story.append(h1("9. Weiterentwicklung", styles))
    story.append(body(
        "Alle Quelldateien befinden sich im GitHub-Repository: "
        '<a href="https://github.com/klausstefanfaszl/invoice_tools" color="#1a3a5c">'
        "https://github.com/klausstefanfaszl/invoice_tools</a>",
        styles,
    ))
    story.append(sp(4))

    tbl = make_table(
        ["Paket", "Verwendungszweck"],
        [
            ["pymupdf",                          "PDF-Verarbeitung und Textextraktion"],
            ["pillow",                           "Bildverarbeitung für OCR-Vorbereitung"],
            ["exchangelib",                      "Exchange Web Services (EWS) Anbindung"],
            ["anthropic / openai / google-genai","KI-Extraktion (Claude / OpenAI / Gemini)"],
            ["reportlab",                        "PDF-Dokumentenerstellung"],
            ["pyinstaller",                      "Erstellen der Windows-Executable"],
        ],
        styles,
        col_widths=[W * 0.38, W * 0.62],
    )
    story.append(tbl)
    story.append(sp(6))

    story.append(body("Build-Befehl:", styles))
    story.append(code_block("pyinstaller invoice_tools.spec", W, styles))

    return story


# ---------------------------------------------------------------------------
# Hauptfunktion
# ---------------------------------------------------------------------------
def main():
    margin = 2 * cm
    doc = SimpleDocTemplate(
        OUTPUT_PDF,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin + 0.5 * cm,   # etwas mehr für Seitenzahl
        title="invoice_tools – Technische Dokumentation",
        author="UHDE Datentechnik",
    )

    usable_width = A4[0] - 2 * margin
    styles = make_styles()
    story = build_content(styles, usable_width)

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(f"Dokumentation erstellt: {OUTPUT_PDF}")


if __name__ == "__main__":
    main()
