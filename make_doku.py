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
        "Technische Dokumentation · Version 1.1 · UHDE Datentechnik",
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
             "dry · unread · all · archiv  (siehe 4.2)"],
            ["--config DATEI", "-c", "invoice_inbox_config.xml",
             "XML-Konfigurationsdatei"],
            ["--debug LEVEL", "-d", "0",
             "0=aus · 1=Pfad · 2=Details · 3=Vollausgabe"],
            ["--log DATEI", "-l", "(kein Log)",
             "Protokolldatei; ohne -d wird stdout vollständig dorthin umgeleitet "
             "(kein Ausgabe auf stdout — Exit-Code weiterhin über %errorlevel% verfügbar)"],
            ["--dry-run", "–", "–",
             "Simulation kombinierbar mit -m all oder -m unread"],
            ["--bzv MODUS", "-b", "–",
             "BankingZV-Export: dry=Anzeige · json=+JSON-Datei · export=+BankingZV-Aufruf"],
            ["--bdir VERZ.", "-B", "–",
             "Basisverzeichnis für PDF-Ablage; überschreibt &lt;BaseDir&gt; aus der Config. "
             "Nützlich wenn die Batch-Skripte von wechselnden Rechnern mit unterschiedlichen "
             "Pfaden aufgerufen werden."],
            ["--export-excel", "-e", "–",
             "Rechnungsdaten in Excel-Eingangstabelle schreiben (Pfad aus &lt;ExcelExport&gt;"
             " in der Config). Duplikat-Prüfung verhindert doppelte Einträge."],
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
        ["Modus", "Speichern", "Als gelesen", "Archivieren", "Welche Mails"],
        [
            ["dry",    "Nein", "Nein", "Nein",
             "Nur ungelesene — zeigt was gespeichert würde"],
            ["unread", "Ja",   "Ja",   "Nein",
             "Nur ungelesene Mails"],
            ["all",    "Ja",   "Ja",   "Nein",
             "Alle Mails (gelesen + ungelesen)"],
            ["archiv", "Ja",   "Ja",   "Ja",
             "Nur ungelesene — verschiebt erfolgreich verarbeitete Mails "
             "in den konfigurierten Archiv-Ordner (Standard: \"Archiv\")"],
        ],
        styles,
        col_widths=[W * 0.13, W * 0.13, W * 0.13, W * 0.13, W * 0.48],
    )
    story.append(tbl)
    story.append(sp(4))
    story.append(note(
        "<b>Archiv-Ordner konfigurieren:</b> &lt;ArchiveFolder&gt;Archiv&lt;/ArchiveFolder&gt; "
        "im &lt;Mailbox&gt;-Block der invoice_inbox_config.xml. "
        "Fehlt der Eintrag, wird bei Modus <i>archiv</i> automatisch der Ordner \"Archiv\" verwendet.",
        styles,
    ))

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
        "invoice_tools.exe inbox -m archiv -b export\n"
        "invoice_tools.exe inbox -m unread -b export -l inbox.log\n"
        "invoice_tools.exe inbox -m dry -e -d 1              (Vorschau Excel-Zellwerte)\n"
        "invoice_tools.exe inbox -m unread -e                (PDF ablegen + Excel schreiben)\n"
        "invoice_tools.exe inbox -m unread -e -b export      (PDF + Excel + BankingZV)\n"
        "invoice_tools.exe inbox -m dry -B \"F:\\Buchhaltung\"  (BaseDir per Parameter)\n"
        "invoice_tools.exe inbox -m unread -B \"F:\\Buchhaltung\" -e",
        W, styles,
    ))
    story.append(sp(4))
    story.append(note(
        "Mit <b>-l inbox.log</b> und ohne <b>-d</b> wird kein Text auf stdout ausgegeben. "
        "Alle Verarbeitungsmeldungen landen in der Logdatei. "
        "Fehlermeldungen (stderr) bleiben immer sichtbar. "
        "Der Exit-Code steht dem aufrufenden Batch-Skript weiterhin über %errorlevel% zur Verfügung.",
        styles,
    ))

    # 4.7 Excel-Export
    story.append(h2("4.7 Excel-Export (--export-excel / -e)", styles))
    story.append(body(
        "Mit dem Flag <b>-e</b> wird jede erfolgreich verarbeitete Rechnung zusätzlich als "
        "neue Zeile in eine Excel-Datei geschrieben. Die Zieltabelle wird über den "
        "Tabellennamen (Standard: <i>tb_rechnungen</i>) referenziert — "
        "Spaltenreihenfolge und -namen werden dynamisch aus der Tabellendefinition ermittelt.",
        styles,
    ))
    story.append(sp(4))

    story.append(body("<b>Befüllte Spalten (konfigurierbar per &lt;Spaltenmapping&gt;):</b>", styles))
    tbl = make_table(
        ["Excel-Spalte", "Quelle", "Typ"],
        [
            ["Re-Datum",       "InvoiceDate",   "datum"],
            ["Name/Lieferant", "SupplierName",  "Text"],
            ["RE-Nr",          "InvoiceNumber", "Text"],
            ["TYP",            "{StandardTyp}", "Text (aus Config)"],
            ["Netto",          "NetAmount",     "betrag"],
            ["Brutto",         "GrossAmount",   "betrag"],
            ["Fällig_am",      "DueDate",       "datum"],
            ["IBAN",           "IBAN",          "iban (erstes Element)"],
            ["Konto-ID",       "{KontoId}",     "Text (aus BankingZV-Routing)"],
        ],
        styles,
        col_widths=[W * 0.25, W * 0.35, W * 0.40],
    )
    story.append(tbl)
    story.append(sp(4))

    story.append(body(
        "Spalten mit <b>calculatedColumnFormula</b> (z.B. Monat, Jahr, Steuer, St%) "
        "werden automatisch erkannt und aus der Vorgängerzeile kopiert.",
        styles,
    ))
    story.append(sp(4))

    story.append(body("<b>Duplikat-Prüfung:</b> Vor dem Schreiben wird die konfigurierte "
                      "Duplikat-Spalte (Standard: <i>RE-Nr</i>) auf Übereinstimmung geprüft. "
                      "Ein bereits vorhandener Eintrag erzeugt einen Info-Hinweis und "
                      "verhindert den Schreibvorgang — kein Fehler, kein Abbruch.", styles))
    story.append(sp(4))

    story.append(body("<b>Dry-Run-Vorschau:</b> Mit <i>-m dry -e -d 1</i> werden die "
                      "Zellwerte aller zu schreibenden Spalten angezeigt, ohne die Datei "
                      "zu verändern. Die Duplikat-Prüfung läuft auch im Dry-Run durch.", styles))
    story.append(sp(4))

    story.append(body("<b>Config-Block &lt;ExcelExport&gt; in invoice_inbox_config.xml:</b>", styles))
    story.append(code_block(
        "<ExcelExport>\n"
        "  <DateiPfad>C:\\Pfad\\zur\\Rechnungseingang.xlsx</DateiPfad>\n"
        "  <Tabellenblatt>ER</Tabellenblatt>\n"
        "  <Tabellenname>tb_rechnungen</Tabellenname>\n"
        "  <StandardTyp>E</StandardTyp>\n"
        "  <DuplikatSpalte>RE-Nr</DuplikatSpalte>\n"
        "  <Spaltenmapping>\n"
        "    <!-- name=Excel-Spalte  typ=datum|betrag|iban|(leer=Text) -->\n"
        "    <!-- Sonderwerte: {StandardTyp} {KontoId}                 -->\n"
        "    <Spalte name=\"Re-Datum\"       typ=\"datum\" >InvoiceDate</Spalte>\n"
        "    <Spalte name=\"Name/Lieferant\"             >SupplierName</Spalte>\n"
        "    <Spalte name=\"RE-Nr\"                      >InvoiceNumber</Spalte>\n"
        "    <Spalte name=\"TYP\"                        >{StandardTyp}</Spalte>\n"
        "    <Spalte name=\"Netto\"          typ=\"betrag\">NetAmount</Spalte>\n"
        "    <Spalte name=\"Brutto\"         typ=\"betrag\">GrossAmount</Spalte>\n"
        "    <Spalte name=\"Faellig_am\"     typ=\"datum\" >DueDate</Spalte>\n"
        "    <Spalte name=\"IBAN\"           typ=\"iban\"  >IBAN</Spalte>\n"
        "    <Spalte name=\"Konto-ID\"                   >{KontoId}</Spalte>\n"
        "  </Spaltenmapping>\n"
        "</ExcelExport>",
        W, styles,
    ))
    story.append(sp(4))
    story.append(note(
        "<b>Hinweis:</b> Fehlt &lt;Spaltenmapping&gt;, wird das eingebaute Standard-Mapping "
        "verwendet. Die Tabelle wird nach jedem Einfügen automatisch auf die neue Zeile ausgedehnt.",
        styles,
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
             "inbox: Exchange-/IMAP-Zugangsdaten, Ablagestruktur, Filter, "
             "BankingZV-Export, Excel-Export (&lt;ExcelExport&gt;)"],
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
        "(Standard: <i>unread</i>) · <i>dry</i> (Simulation) · <i>all</i> (alle Mails) · "
        "<i>archiv</i> (wie unread + verarbeitete Mails in Archiv-Ordner verschieben)",
        styles,
    ))

    # ------------------------------------------------------------------
    # 7. Exit-Codes
    # ------------------------------------------------------------------
    story.append(h1("7. Exit-Codes", styles))
    story.append(body(
        "Das Tool <b>inbox</b> meldet den Verarbeitungsstatus immer als Exit-Code an das "
        "aufrufende Skript — unabhängig davon, ob die Ausgabe auf stdout oder in eine "
        "Log-Datei umgeleitet wurde.",
        styles,
    ))
    story.append(sp(4))

    story.append(h2("7.1 inbox", styles))
    tbl = make_table(
        ["Code", "Konstante", "Bedeutung"],
        [
            ["0", "OK",         "Erfolgreich — alle Rechnungen gespeichert (inkl. keine Mails gefunden)"],
            ["1", "CONFIG",     "Konfigurationsfehler — Config-Datei fehlt, ungültig oder Pflichtfeld leer"],
            ["2", "VERBINDUNG", "Verbindungsfehler — Postfach nicht erreichbar oder Netzwerkfehler"],
            ["3", "KI-ABBRUCH", "KI-Abbruch — alle konfigurierten KI-Provider nicht verfügbar"],
            ["4", "TEILERFOLG", "Teilerfolg — mindestens 1 Rechnung gespeichert, aber mindestens 1 Fehler"],
            ["5", "ALLE FEHL",  "Alle Anhänge fehlgeschlagen — Mails vorhanden, aber 0 Rechnungen gespeichert"],
        ],
        styles,
        col_widths=[W * 0.08, W * 0.18, W * 0.74],
    )
    story.append(tbl)
    story.append(sp(6))

    story.append(h2("7.2 extractor", styles))
    tbl = make_table(
        ["Code", "Bedeutung"],
        [
            ["0", "Erfolgreich — mind. 1 PDF verarbeitet"],
            ["1", "Fehler — Konfiguration ungültig, keine PDFs gefunden oder Ausgabefehler"],
        ],
        styles,
        col_widths=[W * 0.10, W * 0.90],
    )
    story.append(tbl)
    story.append(sp(4))
    story.append(note(
        "Tipp für Batch-Skripte: <b>if %errorlevel% NEQ 0 ...</b> prüft ob ein Fehler vorliegt. "
        "Für differenzierte Auswertung: <b>if %errorlevel% EQU 4 ...</b> usw.",
        styles,
    ))

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
            ["openpyxl",                         "Excel-Export: Schreiben in .xlsx-Tabellen (--export-excel)"],
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
