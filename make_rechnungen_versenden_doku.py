"""
make_rechnungen_versenden_doku.py — Erzeugt Rechnungen_versenden_doku.pdf mit reportlab.
Aufruf: py make_rechnungen_versenden_doku.py
"""

import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)

# ---------------------------------------------------------------------------
# Farben
# ---------------------------------------------------------------------------
DARK_BLUE   = colors.HexColor("#1a3a5c")
LIGHT_BLUE  = colors.HexColor("#f0f4f8")
CODE_BG     = colors.HexColor("#f5f5f5")
CODE_BORDER = colors.HexColor("#cccccc")
GREY_TEXT   = colors.HexColor("#666666")
WHITE       = colors.white

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PDF  = os.path.join(SCRIPT_DIR, "Rechnungen_versenden_doku.pdf")


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
def make_styles():
    s = {}
    s["title"] = ParagraphStyle("title",
        fontName="Helvetica-Bold", fontSize=24, leading=30,
        alignment=TA_CENTER, textColor=DARK_BLUE, spaceAfter=6)
    s["subtitle"] = ParagraphStyle("subtitle",
        fontName="Helvetica", fontSize=10, leading=14,
        alignment=TA_CENTER, textColor=GREY_TEXT, spaceAfter=16)
    s["h1"] = ParagraphStyle("h1",
        fontName="Helvetica-Bold", fontSize=13, leading=17,
        textColor=DARK_BLUE, spaceBefore=14, spaceAfter=5)
    s["h2"] = ParagraphStyle("h2",
        fontName="Helvetica-Bold", fontSize=11, leading=15,
        textColor=DARK_BLUE, spaceBefore=10, spaceAfter=4)
    s["body"] = ParagraphStyle("body",
        fontName="Helvetica", fontSize=9, leading=13, spaceAfter=4)
    s["note"] = ParagraphStyle("note",
        fontName="Helvetica-Oblique", fontSize=9, leading=13,
        textColor=GREY_TEXT, spaceAfter=4)
    s["code"] = ParagraphStyle("code",
        fontName="Courier", fontSize=8, leading=11,
        leftIndent=6, rightIndent=6)
    s["th"] = ParagraphStyle("th",
        fontName="Helvetica-Bold", fontSize=9, leading=12,
        textColor=WHITE, alignment=TA_LEFT)
    s["td"] = ParagraphStyle("td",
        fontName="Helvetica", fontSize=9, leading=12, alignment=TA_LEFT)
    s["td_code"] = ParagraphStyle("td_code",
        fontName="Courier", fontSize=8, leading=11, alignment=TA_LEFT)
    return s


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------
def h1(text, s):
    return KeepTogether([
        Paragraph(text, s["h1"]),
        HRFlowable(width="100%", thickness=1, color=DARK_BLUE, spaceAfter=4),
    ])

def h2(text, s):
    return Paragraph(text, s["h2"])

def body(text, s):
    return Paragraph(text, s["body"])

def note(text, s):
    return Paragraph(text, s["note"])

def sp(height=4):
    return Spacer(1, height)

def code_block(lines, s):
    text = "<br/>".join(lines)
    p = Paragraph(text, s["code"])
    return Table([[p]], colWidths=["100%"], hAlign="LEFT",
                 style=TableStyle([
                     ("BACKGROUND",  (0, 0), (-1, -1), CODE_BG),
                     ("BOX",         (0, 0), (-1, -1), 0.5, CODE_BORDER),
                     ("TOPPADDING",  (0, 0), (-1, -1), 5),
                     ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                     ("LEFTPADDING",  (0, 0), (-1, -1), 8),
                     ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                 ]))

def make_table(header, rows, s, col_widths=None, td_style="td"):
    all_rows = [[Paragraph(str(c), s["th"]) for c in header]]
    for row in rows:
        all_rows.append([Paragraph(str(c), s[td_style]) for c in row])
    tbl = Table(all_rows, colWidths=col_widths, hAlign="LEFT", repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  DARK_BLUE),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  9),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  5),
        ("TOPPADDING",    (0, 0), (-1, 0),  5),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 9),
        ("TOPPADDING",    (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT_BLUE]),
        ("GRID",          (0, 0), (-1, -1), 0.4, CODE_BORDER),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))
    return tbl


# ---------------------------------------------------------------------------
# Inhalt
# ---------------------------------------------------------------------------
def build_content(s):
    W = 17 * cm   # Nutzbereite Breite (A4 minus Ränder)
    story = []

    # ── Titelseite ────────────────────────────────────────────────────────
    story += [
        sp(20),
        Paragraph("Rechnungen_versenden.bat", s["title"]),
        Paragraph("Ausgangsrechnungen per E-Mail versenden", s["subtitle"]),
        HRFlowable(width="100%", thickness=2, color=DARK_BLUE, spaceAfter=20),
        sp(8),
    ]

    # ── 1. Zweck ──────────────────────────────────────────────────────────
    story += [
        h1("1. Zweck", s),
        body("Das Skript <b>Rechnungen_versenden.bat</b> versendet Ausgangsrechnungen "
             "als PDF-Anhang per E-Mail. Es verwendet die Funktion <b>mailto</b> von "
             "<i>invoice_tools.exe</i> und ermittelt die Empfängeradresse automatisch "
             "aus der konfigurierten MySQL-Datenbank (Adressmodus 2).", s),
        body("Das Verzeichnis mit den Rechnungs-PDFs sowie alle SMTP- und "
             "Datenbankeinstellungen werden aus der Konfigurationsdatei "
             "<b>invoice_tools_mailto_config.xml</b> gelesen.", s),
        sp(6),
    ]

    # ── 2. Aufruf ─────────────────────────────────────────────────────────
    story += [
        h1("2. Aufruf", s),
        body("<b>Syntax:</b>", s),
        sp(2),
        code_block([
            "Rechnungen_versenden.bat  [min-rg-nr]  [max-rg-nr]",
        ], s),
        sp(6),
        make_table(
            ["Parameter", "Bedeutung"],
            [
                ["(kein)", "Interaktive Abfrage: Das Skript fragt nach der kleinsten "
                           "zu versendenden Rechnungsnummer."],
                ["min-rg-nr", "Kleinste zu versendende Rechnungsnummer (ganze Zahl). "
                              "Alle PDFs ab dieser Nummer werden versendet."],
                ["max-rg-nr", "Größte zu versendende Rechnungsnummer (optional). "
                              "Nur zusammen mit min-rg-nr verwendbar."],
                ["--rg-datum DATUM", "Nur PDFs versenden, deren Dateidatum dem angegebenen "
                                    "Datum entspricht. "
                                    "Formate: TT.MM  |  TT.MM.JJ  |  TT.MM.JJJJ. "
                                    "Bei TT.MM wird das aktuelle Jahr verwendet."],
            ],
            s,
            col_widths=[3.5*cm, 13.5*cm],
        ),
        sp(8),
    ]

    # ── 3. Beispiele ──────────────────────────────────────────────────────
    story += [
        h1("3. Beispiele", s),
        make_table(
            ["Aufruf", "Wirkung"],
            [
                ["Rechnungen_versenden.bat",
                 "Fragt interaktiv nach der Startrechnungsnummer. "
                 "Alle PDFs ab dieser Nummer werden versendet."],
                ["Rechnungen_versenden.bat 4711",
                 "Versendet alle Rechnungs-PDFs ab RE-Nr 4711 "
                 "(ohne obere Grenze)."],
                ["Rechnungen_versenden.bat 4700 4750",
                 "Versendet alle Rechnungs-PDFs mit RE-Nr 4700 bis 4750."],
                ["Rechnungen_versenden.bat --rg-datum 15.04",
                 "Versendet alle PDFs vom 15.04. des aktuellen Jahres "
                 "(interaktive Abfrage der Rechnungsnummer)."],
                ["Rechnungen_versenden.bat 4700 --rg-datum 15.04.2026",
                 "Versendet PDFs ab RE-Nr 4700 mit Dateidatum 15.04.2026."],
            ],
            s,
            col_widths=[7*cm, 10*cm],
            td_style="td_code",
        ),
        sp(8),
    ]

    # ── 4. Ablauf ─────────────────────────────────────────────────────────
    story += [
        h1("4. Ablauf", s),
        body("Das Skript führt folgende Schritte aus:", s),
        sp(2),
        make_table(
            ["Schritt", "Beschreibung"],
            [
                ["1", "PDF-Verzeichnis durchsuchen: Alle PDFs werden anhand des "
                      "konfigurierten Dateinamen-Regex gefiltert und die Rechnungsnummer "
                      "extrahiert."],
                ["2", "Adressermittlung (Modus 2 – Datenbank): Für jede Rechnungsnummer "
                      "werden bis zu zwei SQL-Abfragen ausgeführt. "
                      "Spalte 1 des Ergebnisses = E-Mail-Adresse. "
                      "Spalte 2 (optional) = Auftragsnummer für zusätzliche Anhänge."],
                ["3", "Auftrags-PDFs suchen (optional): Wenn die Datenbank eine "
                      "Auftragsnummer liefert und &lt;AuftragNrStellen&gt; &gt; 0 konfiguriert "
                      "ist, werden im Auftragsverzeichnis passende PDFs als weitere "
                      "Anhänge hinzugefügt."],
                ["4", "Bestätigung: Das Skript zeigt eine Übersicht aller geplanten "
                      "Sendungen und fragt zur Bestätigung (j/N)."],
                ["5", "Versand per SMTP: Die Mails werden mit allen Anhängen gesendet."],
            ],
            s,
            col_widths=[1.5*cm, 15.5*cm],
        ),
        sp(8),
    ]

    # ── 5. Konfiguration ──────────────────────────────────────────────────
    story += [
        h1("5. Konfiguration", s),
        body("Alle Einstellungen befinden sich in <b>invoice_tools_mailto_config.xml</b>. "
             "Die wichtigsten Abschnitte:", s),
        sp(4),
        make_table(
            ["XML-Element", "Beschreibung"],
            [
                ["&lt;PdfVerzeichnis&gt;",
                 "Verzeichnis mit den Ausgangsrechnungs-PDFs."],
                ["&lt;PdfVerzeichnis_Auftraege&gt;",
                 "Optionales separates Verzeichnis für Auftrags-PDFs (Anhänge). "
                 "Wenn nicht angegeben, wird im &lt;PdfVerzeichnis&gt; gesucht."],
                ["&lt;RgNrRegex&gt;",
                 "Regulärer Ausdruck zum Extrahieren der Rechnungsnummer "
                 "aus dem Dateinamen. Named Group \"rgnr\" bevorzugt."],
                ["&lt;SMTP&gt;",
                 "SMTP-Server, Port, SSL/TLS, Absender und optionale Zugangsdaten."],
                ["&lt;Datenbank&gt; / &lt;SQL1&gt;",
                 "Primäre SQL-Abfrage. Spalte 1 = E-Mail-Adresse, "
                 "Spalte 2 (optional) = Auftragsnummer."],
                ["&lt;Datenbank&gt; / &lt;SQL2&gt;",
                 "Fallback-SQL, wird nur verwendet wenn SQL1 kein Ergebnis liefert."],
                ["&lt;AuftragNrStellen&gt;",
                 "Anzahl der letzten Ziffern der Auftragsnummer für die PDF-Suche. "
                 "0 = deaktiviert."],
                ["&lt;MailVorlage&gt;",
                 "Betreff-Vorlage und HTML-Body der E-Mail. "
                 "Platzhalter: {rgnr} oder {RechnungsNr}."],
            ],
            s,
            col_widths=[5.5*cm, 11.5*cm],
        ),
        sp(8),
    ]

    # ── 6. Exit-Codes ─────────────────────────────────────────────────────
    story += [
        h1("6. Exit-Codes", s),
        make_table(
            ["Code", "Bedeutung"],
            [
                ["0", "Erfolgreich (alle Mails gesendet bzw. Dry-Run ohne Fehler)."],
                ["1", "Konfigurationsfehler (fehlende oder ungültige config-Datei)."],
                ["2", "SMTP-Verbindungsfehler."],
                ["4", "Teilerfolg: mind. 1 Mail gesendet, mind. 1 fehlgeschlagen."],
                ["5", "Alle Mails fehlgeschlagen."],
            ],
            s,
            col_widths=[1.5*cm, 15.5*cm],
        ),
        sp(8),
    ]

    # ── 7. Hinweise ───────────────────────────────────────────────────────
    story += [
        h1("7. Hinweise", s),
        make_table(
            ["Thema", "Hinweis"],
            [
                ["Dry-Run",
                 "Für einen Test ohne echten Versand rechnung_mailto.bat mit "
                 "--dry-run aufrufen. Rechnungen_versenden.bat ist für den "
                 "produktiven Einsatz vorgesehen."],
                ["mysql-connector-python",
                 "Für Adressmodus 2 (Datenbank) muss das Python-Paket installiert sein: "
                 "py -m pip install mysql-connector-python"],
                ["Rechnungsnummer im Dateinamen",
                 "Der Dateiname muss eine Ziffernfolge enthalten, die als "
                 "Rechnungsnummer erkannt wird. Das Muster ist in &lt;RgNrRegex&gt; "
                 "konfigurierbar."],
                ["Laufzeit Python vs. EXE",
                 "Ist invoice_tools.py neuer als invoice_tools.exe, wird "
                 "automatisch Python verwendet."],
            ],
            s,
            col_widths=[4.5*cm, 12.5*cm],
        ),
    ]

    return story


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------
def main():
    s = make_styles()
    doc = SimpleDocTemplate(
        OUTPUT_PDF,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
        title="Rechnungen_versenden.bat – Dokumentation",
        author="invoice_tools",
    )
    doc.build(build_content(s))
    print(f"Erstellt: {OUTPUT_PDF}")


if __name__ == "__main__":
    main()
