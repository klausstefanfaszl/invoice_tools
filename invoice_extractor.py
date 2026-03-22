#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modulares Python-Tool zum Extrahieren von Informationsfeldern aus PDF-Rechnungen.

Extraktions-Priorität:
  1. ZUGferD XML (eingebettet im PDF)
  2. KI-Analyse (Claude / OpenAI / Gemini) — wenn in der Config konfiguriert
  3. Regex-Fallback (wenn keine KI konfiguriert)

Die Felder, Regex-Muster und KI-Konfiguration werden aus einer XML-Config-Datei gelesen.

Debug-Level:
  0  keine Ausgabe (Standard)
  1  Extraktionspfad, Anzahl Felder
  2  Textvorschau, KI-Prompt/-Antwort (gekürzt), ZUGferD-Felder
  3  Volltext, alle Regex-Versuche, vollständige KI-Antwort
"""

import base64
import csv
import io
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Union

import datetime
import fitz          # PyMuPDF
from PIL import Image
# ── Modul-globales Logging ────────────────────────────────────────────────────
_log_fh = None   # wird vom CLI auf eine offene Datei gesetzt

def _log(msg: str) -> None:
    """Schreibt eine Log-Zeile mit Zeitstempel in die Log-Datei (falls gesetzt)
    oder auf stderr."""
    ts  = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"{ts}  {msg}\n"
    if _log_fh is not None:
        _log_fh.write(line)
        _log_fh.flush()
    else:
        sys.stderr.write(line)
        sys.stderr.flush()

try:
    import pytesseract
    pytesseract.get_tesseract_version()   # wirft Exception wenn Binary fehlt
    _TESSERACT_OK = True
except Exception:
    _TESSERACT_OK = False


# ─── ZUGferD TypeCode → Zahlungsart (UNTDID 4461) ────────────────────────────
ZUGFERD_PAYMENT_TYPES = {
    '10': 'Bar',
    '20': 'Scheck',
    '21': 'Banktratte',
    '22': 'Bankscheck',
    '30': 'Überweisung',
    '31': 'Überweisung',
    '42': 'Überweisung',
    '48': 'Kreditkarte',
    '49': 'Lastschrift',
    '50': 'Postgiro',
    '57': 'Dauerauftrag',
    '58': 'Überweisung (SEPA)',
    '59': 'Lastschrift (SEPA)',
    '70': 'Abrechnung',        # Tankkarten, Rahmenverträge (z.B. DKV)
    '74': 'Überweisung',
    '75': 'Überweisung',
    '76': 'Überweisung',
    '97': 'Verrechnung',
}

ResultValue = Optional[Union[str, List[str]]]


class KiAbbruchFehler(Exception):
    """Wird geworfen wenn die KI-API einen fatalen Fehler meldet (Kontingent, Auth).
    Kein Fallback auf Regex — Verarbeitung soll abgebrochen werden."""
    pass

# ─── Datumskonvertierung ──────────────────────────────────────────────────────
_MONTH_NAMES: Dict[str, int] = {
    'jan': 1, 'january': 1, 'januar': 1,
    'feb': 2, 'february': 2, 'februar': 2,
    'mar': 3, 'march': 3, 'mär': 3, 'märz': 3,
    'apr': 4, 'april': 4,
    'may': 5, 'mai': 5,
    'jun': 6, 'june': 6, 'juni': 6,
    'jul': 7, 'july': 7, 'juli': 7,
    'aug': 8, 'august': 8,
    'sep': 9, 'september': 9,
    'oct': 10, 'october': 10, 'okt': 10, 'oktober': 10,
    'nov': 11, 'november': 11,
    'dec': 12, 'december': 12, 'dez': 12, 'dezember': 12,
}


def _normalize_date(value: str) -> str:
    """
    Konvertiert verschiedene Datumsformate nach DD.MM.YYYY.
    Unterstützt: YYYYMMDD, DD.MM.YYYY, DD.MM.YY, MM/DD/YYYY,
                 "Month DD, YYYY", "DD Month YYYY", "D. Monat YYYY"
    """
    if not value or value in ('-', ''):
        return value
    s = value.strip().rstrip('.')

    # YYYYMMDD (ZUGferD)
    m = re.fullmatch(r'(\d{4})(\d{2})(\d{2})', s)
    if m:
        return f'{m.group(3)}.{m.group(2)}.{m.group(1)}'

    # DD.MM.YYYY oder DD.MM.YY
    m = re.fullmatch(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        y = ('20' + y) if len(y) == 2 else y
        return f'{d:02d}.{mo:02d}.{y}'

    # MM/DD/YYYY (Slash → US-Format annehmen; wenn erstes Teil > 12 → DD/MM)
    m = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), m.group(3)
        mo, d = (b, a) if a > 12 else (a, b)   # a > 12 → muss Tag sein
        return f'{d:02d}.{mo:02d}.{y}'

    # "Month DD, YYYY" (Englisch: "January 5, 2026", "Jan 01, 2026")
    m = re.fullmatch(r'([A-Za-zÄÖÜäöüß]+)\.?\s+(\d{1,2}),?\s+(\d{4})', s)
    if m:
        mo = _MONTH_NAMES.get(m.group(1).lower().rstrip('.'))
        if mo:
            return f'{int(m.group(2)):02d}.{mo:02d}.{m.group(3)}'

    # "DD Month YYYY" (Englisch/Deutsch: "02 January 2026", "1. Januar 2026")
    m = re.fullmatch(r'(\d{1,2})\.?\s+([A-Za-zÄÖÜäöüß]+)\.?\s+(\d{4})', s)
    if m:
        mo = _MONTH_NAMES.get(m.group(2).lower().rstrip('.'))
        if mo:
            return f'{int(m.group(1)):02d}.{mo:02d}.{m.group(3)}'

    return value   # Fallback: Original behalten


# ═══════════════════════════════════════════════════════════════════════════════
class InvoiceExtractor:
# ═══════════════════════════════════════════════════════════════════════════════

    def __init__(self, config_path: str, debug_level: int = 0,
                 api_config_path: Optional[str] = None):
        self.config_path     = config_path
        self.debug_level     = debug_level
        self.config          = ET.parse(config_path).getroot()
        self.api_config_path = api_config_path  # None → auto-Suche
        self._dbg(1, f"Config geladen: {config_path}  (debug_level={debug_level})")

    # ── Debug-Ausgabe ─────────────────────────────────────────────────────────

    def _dbg(self, level: int, msg: str) -> None:
        """Schreibt msg ins Log wenn debug_level >= level."""
        if self.debug_level >= level:
            _log(f"[DBG{level}] {msg}")

    # ── Public entry point ────────────────────────────────────────────────────

    def extract(self, pdf_path: str) -> Dict[str, ResultValue]:
        """
        Extrahiert Felder aus einer Rechnung.
        Reihenfolge: ZUGferD XML → KI → Regex (Fallback)
        """
        self._dbg(1, f"Verarbeite: {pdf_path}")
        with fitz.open(pdf_path) as doc:
            self._dbg(1, f"PDF geöffnet: {len(doc)} Seite(n)")

            # 1. ZUGferD XML
            zugferd = self._extract_zugferd_xml(doc)
            if zugferd:
                self._dbg(1, "Extraktionspfad: ZUGferD XML")
                result = self._extract_from_zugferd_xml(zugferd)
                self._dbg(1, f"ZUGferD-Extraktion: {sum(1 for v in result.values() if v)} Felder gefüllt")
                return self._postprocess_result(result)

            # 2. KI-Extraktion (wenn konfiguriert)
            ai_cfg = self._select_ai_config()
            if ai_cfg:
                self._dbg(1, f"Extraktionspfad: KI ({ai_cfg['provider']} / {ai_cfg['model']})")
                result = self._extract_with_ai(doc, ai_cfg)
                self._dbg(1, f"KI-Extraktion: {sum(1 for v in result.values() if v)} Felder gefüllt")
                return self._postprocess_result(result)

            # 3. Regex-Fallback
            text = ''.join(page.get_text() for page in doc)
            if text.strip():
                self._dbg(1, f"Extraktionspfad: Regex (Text, {len(text)} Zeichen)")
                self._dbg(2, f"Textvorschau:\n{text[:500]}\n{'...' if len(text) > 500 else ''}")
                self._dbg(3, f"Vollständiger Text:\n{text}")
                result = self._extract_from_text(text)
            elif _TESSERACT_OK:
                self._dbg(1, "Extraktionspfad: Regex (OCR – kein selektierbarer Text)")
                ocr_text = self._ocr_doc(doc)
                self._dbg(2, f"OCR-Vorschau:\n{ocr_text[:500]}")
                result = self._extract_from_text(ocr_text)
            else:
                self._dbg(1, "Extraktionspfad: Regex übersprungen (kein Text, Tesseract nicht verfügbar)")
                result = {f.get('name'): None for f in self.config.findall('Field') if f.get('name')}

            self._dbg(1, f"Regex-Extraktion: {sum(1 for v in result.values() if v)} Felder gefüllt")
            return self._postprocess_result(result)

    # ── Nachbearbeitung ───────────────────────────────────────────────────────

    def _postprocess_result(self, result: Dict[str, ResultValue]) -> Dict[str, ResultValue]:
        """Normalisiert Betrags- und Datumsfelder; setzt VAT-Fallback auf 0,00 EUR."""
        for field in self.config.findall('Field'):
            name      = field.get('name')
            ftype     = field.get('type', '')
            if name not in result:
                continue

            if ftype == 'amount' and isinstance(result[name], str):
                result[name] = _normalize_amount(result[name])

            if ftype == 'date' and isinstance(result[name], str):
                result[name] = _normalize_date(result[name])

        # VAT: fehlend / "-" / Prozentwert → 0,00 EUR
        if 'VAT' in result:
            v = result['VAT']
            if not v or v in ('-',) or re.fullmatch(r'0+[.,]?0*\s*%?', str(v).strip()):
                result['VAT'] = '0,00 EUR'

        # Berechnete Felder: type="vat_rate" → VAT / NetAmount * 100
        for field in self.config.findall('Field'):
            name  = field.get('name')
            ftype = field.get('type', '')
            if ftype != 'vat_rate' or not name:
                continue
            vat_f = _amount_to_float(result.get('VAT'))
            net_f = _amount_to_float(result.get('NetAmount'))
            if net_f is not None and net_f > 0 and vat_f is not None:
                rate = (vat_f / net_f) * 100
                result[name] = f"{round(rate):.0f} %"
            elif vat_f == 0.0:
                result[name] = '0 %'
            else:
                result[name] = None

        return result

    # ══════════════════════════════════════════════════════════════════════════
    # ZUGferD
    # ══════════════════════════════════════════════════════════════════════════

    def _extract_zugferd_xml(self, doc) -> Optional[tuple]:
        """
        Sucht eingebettete ZUGferD/Factur-X XML im PDF.
        PyMuPDF ≥ 1.18: embfile_info(i) → Metadaten, embfile_get(i) → bytes
        Gibt (root, namespaces) zurück oder None.
        """
        count = doc.embfile_count()
        self._dbg(2, f"Eingebettete Dateien im PDF: {count}")
        for i in range(count):
            info = doc.embfile_info(i)
            filename = info.get('filename') or info.get('name') or ''
            self._dbg(2, f"  Anhang {i}: {filename}")
            if filename.lower().endswith('.xml'):
                try:
                    content: bytes = doc.embfile_get(i)
                    ns: dict = {}
                    for event, elem in ET.iterparse(io.BytesIO(content), events=['start-ns']):
                        ns[elem[0]] = elem[1]
                    root = ET.fromstring(content)
                    if 'CrossIndustryDocument' in root.tag or 'CrossIndustryInvoice' in root.tag:
                        self._dbg(1, f"ZUGferD/Factur-X XML gefunden: {filename}")
                        self._dbg(2, f"  Root-Tag: {root.tag}")
                        self._dbg(2, f"  Namespaces: {ns}")
                        return root, ns
                    else:
                        self._dbg(2, f"  XML ist kein ZUGferD (root: {root.tag})")
                except ET.ParseError as e:
                    self._dbg(2, f"  XML-Parse-Fehler bei {filename}: {e}")
                    continue
        return None

    def _extract_from_zugferd_xml(self, zugferd: tuple) -> Dict[str, ResultValue]:
        root, ns = zugferd
        result = {}
        for field in self.config.findall('Field'):
            name = field.get('name')
            if not name:
                continue
            field_type = field.get('type', '')
            multi = field.get('multi', 'false').lower() == 'true'

            if field_type == 'payment_detection':
                val = self._detect_payment_type_zugferd(root, ns)
                result[name] = val
                self._dbg(2, f"  ZUGferD {name}: {val}")
                continue

            xpath = field.findtext('XPath')
            if xpath:
                if multi:
                    found_all = root.findall(xpath, ns)
                    values = [e.text.strip() for e in found_all if e is not None and e.text]
                    val = self._clean_iban_list(values) if 'IBAN' in name else values
                    result[name] = val
                    self._dbg(2, f"  ZUGferD {name}: {val}")
                else:
                    found = root.find(xpath, ns)
                    val = found.text.strip() if found is not None and found.text else None
                    result[name] = val
                    self._dbg(2, f"  ZUGferD {name}: {val}")
            else:
                result[name] = None
                self._dbg(2, f"  ZUGferD {name}: kein XPath konfiguriert")
        return result

    def _detect_payment_type_zugferd(self, root, ns) -> Optional[str]:
        code_el = root.find('.//ram:SpecifiedTradeSettlementPaymentMeans/ram:TypeCode', ns)
        if code_el is not None and code_el.text:
            return ZUGFERD_PAYMENT_TYPES.get(code_el.text.strip(),
                                              f'TypeCode {code_el.text.strip()}')
        info_el = root.find('.//ram:SpecifiedTradeSettlementPaymentMeans/ram:Information', ns)
        if info_el is not None and info_el.text:
            return info_el.text.strip()
        return None

    # ══════════════════════════════════════════════════════════════════════════
    # KI-Extraktion
    # ══════════════════════════════════════════════════════════════════════════

    def _load_all_ai_configs(self) -> List[dict]:
        """Liest alle <AI>-Blöcke aus der API-Config. Suchreihenfolge:
           1. Explizit per --api/-a angegebene Datei (self.api_config_path)
           2. invoice_tools_api_config.xml neben der Extractor-Config
           3. invoice_tools_api_config.xml neben invoice_extractor.py
           4. <AI>-Block direkt in der Extractor-Config (Fallback)
        """
        api_cfg_root = None
        api_cfg_src  = None

        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_dir = os.path.dirname(os.path.abspath(self.config_path))
        candidates = []
        if self.api_config_path:
            candidates.append(self.api_config_path)
        candidates += [
            os.path.join(config_dir, 'invoice_tools_api_config.xml'),
            os.path.join(script_dir, 'invoice_tools_api_config.xml'),
        ]

        for cand in candidates:
            if os.path.isfile(cand):
                try:
                    api_cfg_root = ET.parse(cand).getroot()
                    api_cfg_src  = cand
                    break
                except ET.ParseError as e:
                    self._dbg(1, f"API-Config Parse-Fehler ({cand}): {e}")

        ai_elements = []
        if api_cfg_root is not None:
            ai_elements = api_cfg_root.findall('AI')
            self._dbg(2, f"KI-Config aus: {api_cfg_src}  ({len(ai_elements)} Provider konfiguriert)")
        else:
            ai_el = self.config.find('AI')
            if ai_el is not None:
                ai_elements = [ai_el]
                self._dbg(2, "KI-Config aus Extractor-Config (Fallback)")

        configs = []
        for ai_el in ai_elements:
            provider = (ai_el.findtext('Provider') or '').strip().lower()
            model    = (ai_el.findtext('Model')    or '').strip()
            api_key  = (ai_el.findtext('APIKey')   or '').strip()
            if provider and api_key:
                configs.append({'provider': provider, 'model': model, 'api_key': api_key})
        return configs

    def _test_ai(self, ai_cfg: dict) -> bool:
        """Sendet eine minimale Testanfrage. Gibt True zurück wenn erfolgreich."""
        provider = ai_cfg['provider']
        try:
            if provider == 'claude':
                from anthropic import Anthropic
                client = Anthropic(api_key=ai_cfg['api_key'])
                client.messages.create(
                    model=ai_cfg['model'] or 'claude-opus-4-6',
                    max_tokens=8,
                    messages=[{'role': 'user', 'content': 'Hi'}],
                )
            elif provider == 'openai':
                from openai import OpenAI
                client = OpenAI(api_key=ai_cfg['api_key'])
                client.chat.completions.create(
                    model=ai_cfg['model'] or 'gpt-4o-mini',
                    max_tokens=8,
                    messages=[{'role': 'user', 'content': 'Hi'}],
                )
            elif provider == 'gemini':
                from google import genai
                client = genai.Client(api_key=ai_cfg['api_key'])
                client.models.generate_content(
                    model=ai_cfg['model'] or 'gemini-2.5-flash-lite',
                    contents='Hi',
                )
            else:
                _log(f"[KI-Test] Unbekannter Provider: {provider}")
                return False
            return True
        except Exception as e:
            _log(f"[KI-Test] {provider}: {e}")
            return False

    def _select_ai_config(self) -> Optional[dict]:
        """Wählt per Testanfrage den ersten funktionierenden KI-Provider.
        Ergebnis wird gecacht — Test läuft nur beim ersten Aufruf."""
        if hasattr(self, '_active_ai_cfg'):
            return self._active_ai_cfg

        configs = self._load_all_ai_configs()
        if not configs:
            self._dbg(2, "Kein <AI>-Block gefunden — KI deaktiviert")
            self._active_ai_cfg = None
            return None

        for cfg in configs:
            key_preview = f"{cfg['api_key'][:20]}...{cfg['api_key'][-10:]}" if len(cfg['api_key']) > 30 else cfg['api_key']
            self._dbg(1, f"Teste KI-Provider: {cfg['provider']} / {cfg['model']}  key={key_preview}")
            if self._test_ai(cfg):
                self._dbg(1, f"KI-Provider ausgewählt: {cfg['provider']} / {cfg['model']}")
                self._active_ai_cfg = cfg
                return cfg
            self._dbg(1, f"  {cfg['provider']} nicht verfügbar — nächster Versuch")

        # Alle Provider fehlgeschlagen
        msg = f"Alle {len(configs)} konfigurierten KI-Provider nicht verfügbar"
        _log(f"[InvoiceExtractor] FATAL — {msg}")
        raise KiAbbruchFehler(msg)

    def _extract_with_ai(self, doc, ai_cfg: dict) -> Dict[str, ResultValue]:
        """
        Sendet den PDF-Inhalt (Text oder Seiten-Bilder) an die konfigurierte KI
        und parst die JSON-Antwort.
        """
        field_specs = self._build_field_specs()

        text = ''.join(page.get_text() for page in doc).strip()
        self._dbg(2, f"Extrahierter Text: {len(text)} Zeichen")
        if text:
            self._dbg(2, f"Textvorschau:\n{text[:500]}\n{'...' if len(text) > 500 else ''}")
        self._dbg(3, f"Vollständiger Text:\n{text}")

        # Seiten als Base64-Bilder für Vision-Modelle (max. 3 Seiten)
        if not text:
            page_images_b64 = self._render_pages_b64(doc, max_pages=3)
        else:
            # Auch bei Text optional Seite 1 als Bild mitschicken (für Layout-Kontext)
            page_images_b64 = self._render_pages_b64(doc, max_pages=1)

        self._dbg(2, f"Seiten-Bilder für KI: {len(page_images_b64)}")

        prompt = self._build_ai_prompt(field_specs, text)
        self._dbg(2, f"KI-Prompt ({len(prompt)} Zeichen):\n{prompt[:1000]}{'...' if len(prompt) > 1000 else ''}")
        self._dbg(3, f"Vollständiger KI-Prompt:\n{prompt}")

        provider = ai_cfg['provider']
        try:
            if provider == 'claude':
                raw = self._call_claude(ai_cfg, prompt, page_images_b64)
            elif provider == 'openai':
                raw = self._call_openai(ai_cfg, prompt, page_images_b64)
            elif provider == 'gemini':
                raw = self._call_gemini(ai_cfg, prompt, page_images_b64)
            else:
                raise ValueError(f"Unbekannter KI-Anbieter: '{provider}'. "
                                 f"Erlaubt: claude, openai, gemini")
        except Exception as e:
            _log(f"[InvoiceExtractor] KI-Fehler ({provider}): {e}")
            self._dbg(1, f"KI-Fehler – Fallback auf Regex")
            text_full = text or self._ocr_doc(doc)
            return self._extract_from_text(text_full)

        self._dbg(2, f"KI-Antwort ({len(raw)} Zeichen):\n{raw[:800]}{'...' if len(raw) > 800 else ''}")
        self._dbg(3, f"Vollständige KI-Antwort:\n{raw}")

        result = self._parse_ai_result(raw, field_specs)

        # Wenn KI kein einziges Feld gefüllt hat → Fallback auf Regex
        if not any(result.values()):
            self._dbg(1, "KI-Antwort leer – Fallback auf Regex")
            text_full = text or self._ocr_doc(doc)
            return self._extract_from_text(text_full)

        return result

    # ── Prompt-Aufbau ─────────────────────────────────────────────────────────

    def _build_field_specs(self) -> List[dict]:
        """Extrahiert Feld-Spezifikationen aus der Config."""
        specs = []
        for field in self.config.findall('Field'):
            name = field.get('name')
            if not name:
                continue
            specs.append({
                'name':  name,
                'multi': field.get('multi', 'false').lower() == 'true',
                'type':  field.get('type', ''),
                'desc':  field.findtext('Description', ''),
            })
        return specs

    def _build_ai_prompt(self, field_specs: List[dict], text: str) -> str:
        lines = []
        for s in field_specs:
            if s['type'] == 'vat_rate':   # wird berechnet, nicht extrahiert
                continue
            suffix = ' (Liste, z.B. ["DE12...", "DE34..."])' if s['multi'] else ''
            desc   = f' – {s["desc"]}' if s['desc'] else ''
            lines.append(f'  "{s["name"]}"{suffix}{desc}')

        field_block = '\n'.join(lines)

        text_block = ''
        if text:
            truncated = text[:8000] + ('...[gekürzt]' if len(text) > 8000 else '')
            text_block = f'\n\nExtrahierter Text der Rechnung:\n"""\n{truncated}\n"""'

        return (
            'Du bist ein Buchhalter-Assistent. Analysiere die folgende Rechnung '
            'und extrahiere die unten aufgelisteten Felder.\n\n'
            'Regeln:\n'
            '- Antworte AUSSCHLIESSLICH mit einem gültigen JSON-Objekt, KEIN erklärender Text\n'
            '- Nicht gefundene Felder: null\n'
            '- Listen-Felder (z.B. IBAN): JSON-Array, auch wenn nur ein Wert\n'
            '- IBANs: Leerzeichen entfernen, maskierte IBANs (mit * oder XXX) weglassen\n'
            '- Beträge: als String mit Original-Formatierung (z.B. "1.234,56")\n'
            '- Daten: im Original-Format belassen\n'
            '- PaymentType: "Gutschrift" (Kreditnote/Credit Note), "Lastschrift", '
            '"Überweisung", "PayPal", "Online", "Vorauszahlung" oder null\n'
            '- SupplierName: der RECHNUNGSAUSSTELLER (die Firma, die Geld verlangt / '
            'die Rechnung ausstellt), NICHT der Empfänger. '
            'Der Empfänger steht oft ganz oben als Adressfeld; ignoriere ihn für dieses Feld. '
            'Suche stattdessen nach Firmenname im Briefkopf, in der Fußzeile, '
            'bei den Bankdaten oder neben "Von:", "Absender:", "Lieferant:" o.ä.\n'
            '- RecipientName: der RECHNUNGSEMPFÄNGER (die Firma oder Person, an die die Rechnung '
            'adressiert ist). Steht meist ganz oben als Adressfeld, nach "An:", "Bill to:", '
            '"Rechnungsempfänger:" o.ä. NUR den Firmennamen, keine Straße/PLZ.\n\n'
            f'Zu extrahierende Felder:\n{field_block}'
            f'{text_block}'
        )

    # ── Claude (Anthropic) ────────────────────────────────────────────────────

    def _call_claude(self, ai_cfg: dict, prompt: str,
                     images_b64: List[str]) -> str:
        try:
            import anthropic
        except ImportError:
            raise ImportError("Paket 'anthropic' fehlt. Installation: pip install anthropic")

        self._dbg(2, "Sende Anfrage an Claude ...")
        client = anthropic.Anthropic(api_key=ai_cfg['api_key'])
        model  = ai_cfg['model'] or 'claude-opus-4-6'

        content = []
        for img in images_b64:
            content.append({
                'type': 'image',
                'source': {
                    'type':       'base64',
                    'media_type': 'image/png',
                    'data':       img,
                }
            })
        content.append({'type': 'text', 'text': prompt})

        with client.messages.stream(
            model=model,
            max_tokens=2048,
            messages=[{'role': 'user', 'content': content}],
        ) as stream:
            response = stream.get_final_message()

        self._dbg(2, f"Claude-Antwort: {response.usage}")
        return next(
            (b.text for b in response.content if b.type == 'text'), ''
        )

    # ── OpenAI ────────────────────────────────────────────────────────────────

    def _call_openai(self, ai_cfg: dict, prompt: str,
                     images_b64: List[str]) -> str:
        try:
            from openai import OpenAI
        except ImportError:
            raise ImportError("Paket 'openai' fehlt. Installation: pip install openai")

        self._dbg(2, "Sende Anfrage an OpenAI ...")
        client = OpenAI(api_key=ai_cfg['api_key'])
        model  = ai_cfg['model'] or 'gpt-4o'

        content = []
        for img in images_b64:
            content.append({
                'type': 'image_url',
                'image_url': {'url': f'data:image/png;base64,{img}', 'detail': 'high'},
            })
        content.append({'type': 'text', 'text': prompt})

        response = client.chat.completions.create(
            model=model,
            messages=[{'role': 'user', 'content': content}],
            max_tokens=2048,
            response_format={'type': 'json_object'},
        )
        self._dbg(2, f"OpenAI-Antwort: {response.usage}")
        return response.choices[0].message.content or ''

    # ── Google Gemini ─────────────────────────────────────────────────────────

    def _call_gemini(self, ai_cfg: dict, prompt: str,
                     images_b64: List[str]) -> str:
        try:
            from google import genai
            from google.genai import types as genai_types
        except ImportError:
            raise ImportError(
                "Paket 'google-genai' fehlt. "
                "Installation: pip install google-genai"
            )

        self._dbg(2, "Sende Anfrage an Gemini ...")
        client = genai.Client(api_key=ai_cfg['api_key'])
        model_name = ai_cfg['model'] or 'gemini-2.5-flash-lite'

        contents = []
        for img in images_b64:
            contents.append(genai_types.Part.from_bytes(
                data=base64.b64decode(img),
                mime_type='image/png',
            ))
        contents.append(prompt)

        response = client.models.generate_content(
            model=model_name,
            contents=contents,
            config=genai_types.GenerateContentConfig(
                response_mime_type='application/json',
            ),
        )
        return response.text or ''

    # ── Antwort-Parsing ───────────────────────────────────────────────────────

    def _parse_ai_result(self, raw: str,
                         field_specs: List[dict]) -> Dict[str, ResultValue]:
        """Parst die JSON-Antwort der KI und bereinigt die Werte."""
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        if not json_match:
            _log(f"[InvoiceExtractor] KI lieferte kein JSON: {raw[:200]}")
            return {s['name']: ([] if s['multi'] else None) for s in field_specs}

        try:
            data = json.loads(json_match.group())
        except json.JSONDecodeError as e:
            _log(f"[InvoiceExtractor] JSON-Parse-Fehler: {e}\n{raw[:200]}")
            return {s['name']: ([] if s['multi'] else None) for s in field_specs}

        result: Dict[str, ResultValue] = {}
        for spec in field_specs:
            name  = spec['name']
            multi = spec['multi']
            val   = data.get(name)

            if val is None or val == '':
                result[name] = [] if multi else None
            elif multi:
                if isinstance(val, list):
                    items = [str(v).strip() for v in val if v]
                else:
                    items = [str(val).strip()]
                result[name] = self._clean_iban_list(items) if 'IBAN' in name else items
            else:
                result[name] = str(val).strip() if val else None

            self._dbg(2, f"  KI {name}: {result[name]}")

        return result

    # ══════════════════════════════════════════════════════════════════════════
    # Regex-Extraktion (Fallback)
    # ══════════════════════════════════════════════════════════════════════════

    def _extract_from_text(self, text: str) -> Dict[str, ResultValue]:
        result = {}
        for field in self.config.findall('Field'):
            name = field.get('name')
            if not name:
                continue
            field_type = field.get('type', '')
            multi = field.get('multi', 'false').lower() == 'true'

            if field_type == 'payment_detection':
                result[name] = self._detect_payment_type_text(field, text)
            elif multi:
                result[name] = self._findall_first_regex(field, text)
            else:
                result[name] = self._search_first_regex(field, text)

            self._dbg(2, f"  Regex {name}: {result[name]}")
        return result

    def _search_first_regex(self, field: ET.Element, text: str) -> Optional[str]:
        for regex_el in field.findall('Regex'):
            pattern = regex_el.text
            if not pattern:
                continue
            m = re.search(pattern, text, re.MULTILINE | re.IGNORECASE | re.DOTALL)
            if m:
                for g in m.groups():
                    if g is not None:
                        self._dbg(3, f"    Regex TREFFER [{pattern[:60]}]: {g.strip()!r}")
                        return g.strip()
            else:
                self._dbg(3, f"    Regex kein Treffer [{pattern[:60]}]")
        return None

    def _findall_first_regex(self, field: ET.Element, text: str) -> List[str]:
        for regex_el in field.findall('Regex'):
            pattern = regex_el.text
            if not pattern:
                continue
            matches = re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)
            if matches:
                values = []
                for m in matches:
                    val = (m if isinstance(m, str) else (m[0] if m else '')).strip()
                    if val:
                        values.append(val)
                if 'IBAN' in (field.get('name') or ''):
                    values = self._clean_iban_list(values)
                self._dbg(3, f"    findall TREFFER [{pattern[:60]}]: {values}")
                return values
            else:
                self._dbg(3, f"    findall kein Treffer [{pattern[:60]}]")
        return []

    def _detect_payment_type_text(self, field: ET.Element, text: str) -> Optional[str]:
        for kw_el in field.findall('Keyword'):
            category = kw_el.get('category', '')
            keyword  = kw_el.text or ''
            if keyword and re.search(re.escape(keyword), text, re.IGNORECASE):
                self._dbg(3, f"    Keyword TREFFER [{keyword!r}] → {category}")
                return category
        return None

    # ══════════════════════════════════════════════════════════════════════════
    # Hilfsmethoden
    # ══════════════════════════════════════════════════════════════════════════

    def _ocr_doc(self, doc) -> str:
        """OCR-Verarbeitung für Bild-PDFs mit Tesseract. Leer wenn Tesseract nicht verfügbar."""
        if not _TESSERACT_OK:
            self._dbg(1, "OCR übersprungen – Tesseract nicht installiert")
            return ''
        self._dbg(2, "Starte OCR ...")
        text = ''
        for i, page in enumerate(doc):
            pix = page.get_pixmap()
            img = Image.open(io.BytesIO(pix.tobytes('png')))
            page_text = pytesseract.image_to_string(img)
            self._dbg(2, f"  OCR Seite {i+1}: {len(page_text)} Zeichen")
            text += page_text
        return text

    def _render_pages_b64(self, doc, max_pages: int = 3) -> List[str]:
        """Rendert PDF-Seiten als Base64-kodierte PNG-Bilder."""
        images = []
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            pix  = page.get_pixmap(dpi=150)
            data = base64.standard_b64encode(pix.tobytes('png')).decode('utf-8')
            self._dbg(2, f"  Seite {i+1} gerendert: {len(data)//1024} KB (base64)")
            images.append(data)
        return images

    @staticmethod
    def _validate_iban(iban: str) -> bool:
        """Prüft eine IBAN per Modulo-97-Algorithmus (ISO 13616)."""
        iban = re.sub(r'\s+', '', iban).upper()
        if not re.fullmatch(r'[A-Z]{2}[0-9]{2}[A-Z0-9]{11,30}', iban):
            return False
        rearranged = iban[4:] + iban[:4]
        numeric = ''.join(str(ord(c) - 55) if c.isalpha() else c for c in rearranged)
        return int(numeric) % 97 == 1

    @staticmethod
    def _clean_iban_list(ibans: List[str]) -> List[str]:
        """Entfernt Whitespace aus IBANs, filtert maskierte und ungültige heraus, dedupliziert."""
        result = []
        seen   = set()
        for iban in ibans:
            clean = re.sub(r'\s+', '', iban)
            if '*' in clean or re.search(r'X{3,}', clean):
                continue
            if not InvoiceExtractor._validate_iban(clean):
                continue
            if clean not in seen:
                seen.add(clean)
                result.append(clean)
        return result


# ═══════════════════════════════════════════════════════════════════════════════
# Betrags-Hilfsfunktionen
# ═══════════════════════════════════════════════════════════════════════════════

def _amount_to_float(value) -> Optional[float]:
    """Parst einen normalisierten Betrag ('1.234,56 EUR') in einen float.
    Gibt None zurück wenn der Wert nicht parsbar ist."""
    if not value or value in ('-', None):
        return None
    s = re.sub(r'[A-Z€$£\s]', '', str(value)).strip()
    if not s or re.fullmatch(r'[\-]+', s):
        return None
    # Deutsches Format: Punkt = Tausender, Komma = Dezimal
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# Betrags-Normalisierung
# ═══════════════════════════════════════════════════════════════════════════════

def _normalize_amount(value: str) -> str:
    """
    Normalisiert einen Geldbetrag:
    - Entfernt Währungssymbole (€, $, £) und ersetzt sie durch Codes (EUR, USD, GBP)
    - Hängt 'EUR' an wenn keine Währung erkennbar ist
    - Kein €-Zeichen in der Ausgabe
    Beispiele:  '€ 335,16'   → '335,16 EUR'
                '$20.00'     → '20.00 USD'
                '$34.99 USD' → '34.99 USD'
                '107,10'     → '107,10 EUR'
                '0%'         → '0%'   (Sonderformat – unverändert)
    """
    if not value or value in ('-', ''):
        return value
    s = value.strip()
    # Prozentwerte (z.B. "0%") nicht verändern
    if re.fullmatch(r'[\d.,]+\s*%', s):
        return value
    currency = None
    # Reihenfolge: längere/spezifischere Muster zuerst
    for sym, code in [('EUR€', 'EUR'), ('EUR', 'EUR'), ('CHF', 'CHF'),
                      ('USD', 'USD'), ('GBP', 'GBP'), ('€', 'EUR'),
                      ('$', 'USD'), ('£', 'GBP')]:
        if sym in s:
            currency = code
            s = s.replace(sym, '').strip()
            break
    # Verbleibende Währungssymbole entfernen (z.B. "$34.99 USD" → nach USD-Entfernung bleibt "$34.99")
    for sym in ('€', '$', '£'):
        s = s.replace(sym, '').strip()
    # Nicht ändern wenn kein Zahlenwert übrig
    if not any(c.isdigit() for c in s):
        return value
    return f'{s} {currency or "EUR"}'


# ═══════════════════════════════════════════════════════════════════════════════
# Ausgabe-Formatierung
# ═══════════════════════════════════════════════════════════════════════════════

def _csv_cell(value) -> str:
    """Konvertiert einen Feldwert in einen CSV-String (Listen mit ' | ' verbunden)."""
    if value is None:
        return ''
    if isinstance(value, list):
        return ' | '.join(str(v) for v in value)
    return str(value)


def _format_json(results: list, single: bool) -> str:
    if single:
        data = {k: v for k, v in results[0].items() if k != '_file'}
        return json.dumps(data, indent=2, ensure_ascii=False)
    out = []
    for r in results:
        entry = {'Datei': r.get('_file', '')}
        entry.update({k: v for k, v in r.items() if k != '_file'})
        out.append(entry)
    return json.dumps(out, indent=2, ensure_ascii=False)


def _format_csv(results: list, single: bool, field_names: list) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=';', quoting=csv.QUOTE_MINIMAL)
    if single:
        writer.writerow(field_names)
        writer.writerow([_csv_cell(results[0].get(n)) for n in field_names])
    else:
        writer.writerow(['Datei'] + field_names)
        for r in results:
            writer.writerow(
                [r.get('_file', '')] + [_csv_cell(r.get(n)) for n in field_names]
            )
    return buf.getvalue()


def _format_xml(results: list, single: bool) -> str:
    from xml.dom import minidom
    root_el = ET.Element('Invoices')
    for r in results:
        inv_el = ET.SubElement(root_el, 'Invoice')
        if not single:
            inv_el.set('file', r.get('_file', ''))
        for k, v in r.items():
            if k == '_file':
                continue
            child = ET.SubElement(inv_el, k)
            if isinstance(v, list):
                for item in v:
                    ET.SubElement(child, 'Item').text = str(item)
            elif v is not None:
                child.text = str(v)
    raw = ET.tostring(root_el, encoding='unicode')
    return minidom.parseString(raw).toprettyxml(indent='  ')


def _txt_cell(value) -> str:
    if value is None:
        return '-'
    if isinstance(value, list):
        return ', '.join(str(v) for v in value) if value else '-'
    return str(value)


def _fmt_sum(val: float) -> str:
    """Formatiert einen Summenbetrag im deutschen Format: 1.234,56 EUR"""
    s = f"{val:,.2f}"                                   # "1,234.56"
    s = s.replace(',', 'X').replace('.', ',').replace('X', '.')  # "1.234,56"
    return f"{s} EUR"


def _compute_subtotals(results: list, sum_fields: set) -> dict:
    """Summiert alle Betragsfelder über alle Ergebnisse."""
    totals = {}
    for name in sum_fields:
        values = [_amount_to_float(r.get(name)) for r in results]
        valid  = [v for v in values if v is not None]
        totals[name] = sum(valid) if valid else None
    return totals


def _format_txt(results: list, single: bool, field_names: list,
                amount_fields: set = None, labels: dict = None,
                subtotals: bool = False, sum_fields: set = None) -> str:
    amount_fields = amount_fields or set()
    sum_fields    = sum_fields or set()
    labels        = labels or {}
    # names: für Daten-Lookup und Ausrichtungsprüfung (interne Feldnamen)
    # hdrs:  für Spaltenüberschriften (Labels aus Config)
    names = field_names if single else ['_file'] + field_names
    hdrs  = ([labels.get(n, n) for n in field_names] if single
             else ['Datei'] + [labels.get(n, n) for n in field_names])

    rows = []
    for r in results:
        row = []
        if not single:
            row.append(r.get('_file', ''))
        row += [_txt_cell(r.get(n)) for n in field_names]
        rows.append(row)

    widths = [len(h) for h in hdrs]
    for row in rows:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(cell))

    # Summenzeile vorberechnen (beeinflusst Spaltenbreiten)
    sum_row = None
    if subtotals and not single and len(results) > 1:
        totals   = _compute_subtotals(results, sum_fields)
        sum_label = 'Gesamt'
        sum_row  = []
        if not single:
            sum_row.append(sum_label)
        for n in field_names:
            v = totals.get(n)
            sum_row.append(_fmt_sum(v) if v is not None else '')
        for i, cell in enumerate(sum_row):
            widths[i] = max(widths[i], len(cell))

    sep     = '-+-'.join('-' * w for w in widths)
    sum_sep = '=+='.join('=' * w for w in widths)
    hdr     = ' | '.join(h.ljust(widths[i]) for i, h in enumerate(hdrs))
    lines   = [hdr, sep]
    for row in rows:
        cells = []
        for i, (name, cell) in enumerate(zip(names, row)):
            cells.append(cell.rjust(widths[i]) if name in amount_fields
                         else cell.ljust(widths[i]))
        lines.append(' | '.join(cells))

    if sum_row is not None:
        lines.append(sum_sep)
        cells = []
        for i, (name, cell) in enumerate(zip(names, sum_row)):
            cells.append(cell.rjust(widths[i]) if name in amount_fields
                         else cell.ljust(widths[i]))
        lines.append(' | '.join(cells))

    return '\n'.join(lines) + '\n'


def _format_pdf(results: list, single: bool, field_names: list,
                amount_fields: set = None, labels: dict = None,
                subtotals: bool = False, sum_fields: set = None) -> bytes:
    """Erzeugt ein PDF mit einer formatierten Tabelle (Querformat A4).
    Betragsfelder werden rechtsbündig ausgerichtet.
    Lange Dateinamen und mehrere IBANs umbrechen in der Zelle.
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_RIGHT
        import io as _io
    except ImportError:
        raise ImportError("Paket 'reportlab' fehlt. Installation: pip install reportlab")

    amount_fields = amount_fields or set()
    sum_fields    = sum_fields or set()
    labels        = labels or {}

    buf = _io.BytesIO()
    margin = 10 * mm
    page_w, _ = landscape(A4)
    usable_w = page_w - 2 * margin

    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=margin, rightMargin=margin,
        topMargin=12 * mm, bottomMargin=10 * mm,
        title='Rechnungsübersicht',
    )

    cols = field_names if single else ['Datei'] + field_names

    # Spaltenbreiten proportional zum erwarteten Inhalt
    col_weights = {
        'Datei':          10, 'SupplierName':   11, 'InvoiceNumber':   8,
        'CustomerNumber':  6, 'InvoiceDate':     5, 'DueDate':         5,
        'NetAmount':       7, 'GrossAmount':     7, 'VAT':             6,
        'IBAN':           18, 'PaymentType':     7,
    }
    total_w = sum(col_weights.get(c, 7) for c in cols)
    col_widths = [usable_w * col_weights.get(c, 7) / total_w for c in cols]

    # Paragraph-Styles (Latin-1-kompatibel)
    def _style(align):
        return ParagraphStyle(
            'cell', fontName='Helvetica', fontSize=6.5, leading=8.5,
            alignment=align, wordWrap='LTR', spaceAfter=0, spaceBefore=0,
        )
    def _hdr_style():
        return ParagraphStyle(
            'hdr', fontName='Helvetica-Bold', fontSize=7, leading=9,
            alignment=TA_LEFT, textColor=colors.white,
        )

    def _safe(text: str) -> str:
        return text.encode('latin-1', errors='replace').decode('latin-1')

    def _pdf_cell(col: str, value) -> Paragraph:
        """Erzeugt einen Paragraph für eine Tabellenzelle."""
        align = TA_RIGHT if col in amount_fields else TA_LEFT
        if isinstance(value, list):
            # Mehrere Werte (z.B. IBANs) zeilenweise
            text = '<br/>'.join(_safe(str(v)) for v in value) if value else '-'
        else:
            text = _safe(str(value)) if value not in (None, '') else '-'
        return Paragraph(text, _style(align))

    # Tabellendaten aufbauen — Überschriften aus Labels, Daten aus Feldnamen
    header_row = [Paragraph(_safe(labels.get(c, c)), _hdr_style()) for c in cols]
    data_rows = []
    for r in results:
        row = []
        if not single:
            row.append(_pdf_cell('Datei', r.get('_file', '')))
        for n in field_names:
            row.append(_pdf_cell(n, r.get(n)))
        data_rows.append(row)

    # Summenzeile aufbauen
    sum_row = None
    if subtotals and not single and len(results) > 1:
        totals  = _compute_subtotals(results, sum_fields)
        def _sum_cell(col: str) -> Paragraph:
            v = totals.get(col)
            if v is not None:
                return Paragraph(_safe(_fmt_sum(v)),
                                 _style(TA_RIGHT))
            elif col == (cols[0]):
                return Paragraph(_safe('Gesamt'), _style(TA_LEFT))
            else:
                return Paragraph('', _style(TA_LEFT))
        sum_row = [_sum_cell(c) for c in cols]

    table_data = [header_row] + data_rows
    if sum_row:
        table_data.append(sum_row)

    n_data = len(data_rows)
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        # Header-Hintergrund
        ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor('#2C3E50')),
        # Zeilenhintergrund alternierend (nur Datenzeilen)
        ('ROWBACKGROUNDS', (0, 1), (-1, n_data), [colors.white, colors.HexColor('#F2F2F2')]),
        # Gitter
        ('GRID',          (0, 0), (-1, -1), 0.3, colors.HexColor('#BBBBBB')),
        # Padding
        ('TOPPADDING',    (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING',   (0, 0), (-1, -1), 3),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 3),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
    ]
    # Summenzeile: fette Schrift, dunkler Hintergrund
    if sum_row:
        last = len(table_data) - 1
        style_cmds += [
            ('BACKGROUND',  (0, last), (-1, last), colors.HexColor('#D0D8E4')),
            ('FONTNAME',    (0, last), (-1, last), 'Helvetica-Bold'),
            ('LINEABOVE',   (0, last), (-1, last), 0.8, colors.HexColor('#2C3E50')),
        ]
    table.setStyle(TableStyle(style_cmds))

    doc.build([table])
    return buf.getvalue()


def _is_invoice(data: dict) -> bool:
    """Prüft ob die extrahierten Daten wie eine Rechnung aussehen.
    Mindestens 2 der Kernfelder müssen gefüllt sein."""
    core = ['InvoiceNumber', 'GrossAmount', 'SupplierName', 'InvoiceDate']
    filled = sum(1 for f in core if data.get(f))
    return filled >= 2


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

def run(argv=None):
    import argparse
    import glob as _glob

    _script_dir  = os.path.dirname(os.path.abspath(__file__))
    _default_cfg = os.path.join(_script_dir, 'invoice_extractor_config.xml')

    parser = argparse.ArgumentParser(
        prog=os.path.basename(sys.argv[0]),
        description=(
            'Extrahiert Rechnungsfelder aus PDF-Dateien.\n'
            'Extraktionspfad: ZUGferD-XML -> KI (Claude/OpenAI/Gemini) -> Regex\n'
            '\n'
            'Beispiele:\n'
            '  invoice_tools extractor rechnung.pdf\n'
            '  invoice_tools extractor -f csv "/pfad/*.pdf"\n'
            '  invoice_tools extractor -f txt -d 1 rechnung.pdf\n'
            '  invoice_tools extractor -c andere_config.xml rechnung.pdf'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        'pdf',
        nargs='*',                      # '*' statt '+' → erlaubt 0 Argumente für --help
        metavar='PDF',
        help='Eine oder mehrere PDF-Dateien (Wildcards in Anführungszeichen, z.B. "*.pdf")',
    )
    parser.add_argument(
        '-c', '--config',
        dest='config',
        default=_default_cfg,
        metavar='CONFIGDATEI',
        help='XML-Konfigurationsdatei (Standard: invoice_extractor_config.xml)',
    )
    parser.add_argument(
        '-f', '--format',
        dest='fmt',
        default='stdout',
        choices=['stdout', 'json', 'csv', 'xml', 'txt', 'pdf'],
        metavar='FORMAT',
        help='Ausgabeformat: stdout | json | csv | xml | txt | pdf  (Standard: stdout)',
    )
    parser.add_argument(
        '-o', '--output',
        dest='output',
        default=None,
        metavar='DATEI',
        help='Ausgabedatei (Standard: automatisch aus Eingabedatei). '
             '"-o STDOUT" gibt auf stdout aus. '
             'Fehlende Dateiextension wird automatisch ergänzt.',
    )
    parser.add_argument(
        '-d', '--debug',
        dest='debug_level',
        type=int,
        default=None,
        metavar='LEVEL',
        help='Debug-Level (0=aus, 1=Pfad, 2=Details, 3=Vollausgabe)',
    )
    parser.add_argument(
        '-a', '--api',
        dest='api_config',
        default=None,
        metavar='API_CONFIG',
        help='Zentrale KI-API-Konfigurationsdatei (Standard: invoice_tools_api_config.xml '
             'neben der Extractor-Config oder neben invoice_extractor.py)',
    )
    parser.add_argument(
        '-l', '--logfile',
        dest='logfile',
        default=None,
        metavar='DATEI',
        help='Log-Datei (wird neu erstellt/überschrieben). '
             'Setzt debug-Level auf mindestens 1 wenn kein -d angegeben.',
    )

    # Hilfe anzeigen wenn keine Argumente übergeben
    if argv is not None and len(argv) == 0:
        parser.print_help()
        sys.exit(0)

    args = parser.parse_args(argv)

    # Logfile öffnen (überschreiben) und debug-Level ggf. anheben
    if args.logfile:
        def _open_logfile(path):
            global _log_fh
            _log_fh = open(path, 'w', encoding='utf-8')
        try:
            _open_logfile(args.logfile)
        except OSError as e:
            parser.error(f"Log-Datei kann nicht erstellt werden: {e}")
        if args.debug_level is None:
            args.debug_level = 1
        _log(f"InvoiceExtractor gestartet  (debug_level={args.debug_level})")
        _log(f"Config: {args.config}")

    # stdout-Format: Debug-Ausgaben würden die Ausgabe zerstören
    # Ausnahme: wenn ein Logfile gesetzt ist, landet debug dort — kein Konflikt
    # Nur unterdrücken wenn kein explizites -d angegeben wurde
    if args.fmt == 'stdout' and not args.logfile and args.debug_level is None:
        args.debug_level = 0
    elif args.debug_level is None:
        args.debug_level = 0

    if not args.pdf:
        parser.print_help()
        sys.exit(0)

    if not os.path.isfile(args.config):
        parser.error(f"Konfigurationsdatei nicht gefunden: {args.config}")

    # Wildcards expandieren
    pdf_files = []
    for pattern in args.pdf:
        expanded = _glob.glob(pattern)
        pdf_files.extend(expanded if expanded else [pattern])
    # Immer nach Dateiname sortieren (unabhängig von Eingabereihenfolge)
    pdf_files = sorted(pdf_files, key=lambda p: os.path.basename(p).lower())

    if not pdf_files:
        parser.error("Keine PDF-Dateien gefunden")

    # PDFs verarbeiten
    extractor = InvoiceExtractor(args.config, debug_level=args.debug_level,
                                 api_config_path=args.api_config)
    results   = []
    skipped   = 0
    for pdf_path in pdf_files:
        if not os.path.isfile(pdf_path):
            _log(f"[WARNUNG] Datei nicht gefunden: {pdf_path}")
            continue
        try:
            data = extractor.extract(pdf_path)
        except KiAbbruchFehler as e:
            print(f"\nFEHLER: {e}", file=sys.stderr)
            print("Verarbeitung abgebrochen — KI-API nicht verfügbar.", file=sys.stderr)
            sys.exit(2)
        if not _is_invoice(data):
            _log(f"[ÜBERSPRUNGEN] Kein Rechnungsinhalt erkannt: {os.path.basename(pdf_path)}")
            skipped += 1
            continue
        data['_file'] = os.path.basename(pdf_path)
        results.append(data)

    if not results:
        print(f"Keine Rechnungen gefunden ({skipped} Datei(en) übersprungen).", file=sys.stderr)
        sys.exit(1)

    # Feldnamen, Labels, Betragsfelder und Summieroption aus Config
    config_root = ET.parse(args.config).getroot()
    field_names   = [f.get('name') for f in config_root.findall('Field') if f.get('name')]
    amount_fields = {f.get('name') for f in config_root.findall('Field')
                     if f.get('type') in ('amount', 'vat_rate') and f.get('name')}
    sum_fields    = {f.get('name') for f in config_root.findall('Field')
                     if f.get('type') == 'amount' and f.get('name')}
    labels        = {f.get('name'): f.get('label')
                     for f in config_root.findall('Field')
                     if f.get('name') and f.get('label')}
    subtotals     = config_root.get('subtotals', 'false').lower() == 'true'

    # single: genau eine Eingabedatei angegeben (nicht Anzahl Ergebnisse)
    single = len(pdf_files) == 1
    fmt    = args.fmt.lower()

    # stdout: direkt ausgeben, keine Datei schreiben
    # Tritt ein bei fmt==stdout oder wenn -o STDOUT übergeben wurde
    if fmt == 'stdout' or (args.output and args.output.upper() == 'STDOUT'):
        if fmt in ('stdout', 'txt'):
            content = _format_txt(results, single, field_names, amount_fields, labels, subtotals, sum_fields)
        elif fmt == 'json':
            content = _format_json(results, single)
        elif fmt == 'csv':
            content = _format_csv(results, single, field_names)
        elif fmt == 'xml':
            content = _format_xml(results, single)
        else:
            parser.error("Format 'pdf' kann nicht auf stdout ausgegeben werden.")
        sys.stdout.write(content)
        skipped_info = f', {skipped} übersprungen' if skipped else ''
        print(f"({len(results)} Rechnung(en){skipped_info})", file=sys.stderr)
        sys.exit(0)

    # Ausgabedatei bestimmen
    if args.output:
        out_path = args.output
        # Extension ergänzen wenn fehlend oder falsch
        expected_ext = f'.{fmt}'
        if not out_path.lower().endswith(expected_ext):
            out_path += expected_ext
    elif single:
        out_path = os.path.splitext(os.path.basename(pdf_files[0]))[0] + f'.{fmt}'
    else:
        out_path = f'invoice_tools_output.{fmt}'

    # Formatieren und schreiben
    if fmt == 'pdf':
        binary = _format_pdf(results, single, field_names, amount_fields, labels, subtotals, sum_fields)
        with open(out_path, 'wb') as fh:
            fh.write(binary)
    else:
        if fmt == 'json':
            content = _format_json(results, single)
        elif fmt == 'csv':
            content = _format_csv(results, single, field_names)
        elif fmt == 'xml':
            content = _format_xml(results, single)
        else:
            content = _format_txt(results, single, field_names, amount_fields, labels, subtotals, sum_fields)
        with open(out_path, 'w', encoding='utf-8') as fh:
            fh.write(content)

    skipped_info = f', {skipped} übersprungen' if skipped else ''
    print(f"Ausgabe: {out_path}  ({len(results)} Rechnung(en){skipped_info})")


if __name__ == '__main__':
    run()
