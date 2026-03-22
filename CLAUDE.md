# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This workspace contains **invoice_tools** — a Python toolkit for extracting structured data from PDF invoices and for fetching/filing invoices from email inboxes. It produces output in multiple formats and can be built into a standalone Windows executable.

## Running the Tools

The unified entry point is `invoice_tools.py` (or `invoice_tools.exe` after building):

```bash
# Extract fields from a single PDF
python3 invoice_tools.py extractor rechnung.pdf

# Extract multiple PDFs → PDF summary
python3 invoice_tools.py extractor -f pdf -o output.pdf "/pfad/*.pdf"

# Extract with custom config and debug output
python3 invoice_tools.py extractor -c invoice_extractor_config_RE.xml -d 2 rechnung.pdf

# Process inbox (dry-run shows what would be saved)
python3 invoice_tools.py inbox -m dry
python3 invoice_tools.py inbox -m unread
python3 invoice_tools.py inbox -m all -c invoice_inbox_config.xml

# Help for each sub-tool
python3 invoice_tools.py extractor --help
python3 invoice_tools.py inbox --help
```

The individual scripts (`invoice_extractor.py`, `inbox_processor.py`) can also be run directly.

**Output formats:** `stdout` | `json` | `csv` | `xml` | `txt` | `pdf`

**Debug levels:** `0`=off, `1`=extraction path + field count, `2`=text preview + AI prompt/response, `3`=full text + all regex attempts

**Extractor CLI flags:** `-c CONFIG` | `-f FORMAT` | `-o OUTPUT_FILE` | `-d DEBUG_LEVEL` | `-l LOG_FILE`

## Building the Executable (PyInstaller)

```bash
# Recommended: combined exe (extractor + inbox in one binary)
pyinstaller invoice_tools.spec
# Output: dist/invoice_tools/invoice_tools.exe

# Individual executables (legacy)
pyinstaller invoice_extractor.spec   # dist/invoice_extractor/invoice_extractor
pyinstaller inbox_processor.spec     # dist/inbox_processor/inbox_processor
```

## Windows Batch Scripts

Four `.bat` scripts wrap the exe for typical monthly workflows. All read `BASE_DIR` from `F:\Dokumente\UHDE intern\Buchhaltung`.

| Script | Purpose |
|--------|---------|
| `rechnungseingang_in.bat [dry\|unread\|all]` | Fetch incoming invoices from email inbox → file as `ER_MMTT_Supplier_Nr.pdf` |
| `eingangsrechnungen.bat [dry\|unread\|all]` | Same as above (alternate name) |
| `rechnungseingang_ex.bat JJJJ/MM` | Summarise existing PDFs in `…\MM\Eingangsrechnungen\` → `Rechnungseingang_MM.pdf` |
| `rechnungsausgang_ex.bat JJJJ/MM` | Summarise existing PDFs in `…\MM\Ausgangsrechnungen\` → `Rechnungsausgang_MM.pdf` |

Month parameter accepts `JJJJ/MM` or just `MM` (current year assumed).

## Architecture

### Entry Points

- `invoice_tools.py` — Dispatcher; routes `extractor` → `invoice_extractor.main()` and `inbox` → `inbox_processor.main()`
- `invoice_extractor.py` — `InvoiceExtractor` class + CLI for PDF field extraction
- `inbox_processor.py` — Connects to Exchange/IMAP, iterates unread mails, extracts PDF attachments via `InvoiceExtractor`, and files them

### Extraction Pipeline (`InvoiceExtractor.extract()`)

Three strategies run in priority order:

1. **ZUGferD XML** — Embedded XML attachment in the PDF (`order-x.xml`, `zugferd-invoice.xml`, etc.). Parsed via XPath expressions from the config. Most reliable and fastest.
2. **AI extraction** — If an `<AI>` block is present in the config, the PDF text (or page images for scanned PDFs) is sent to Claude/OpenAI/Gemini. The AI returns a JSON object keyed by field names.
3. **Regex fallback** — Each `<Field>` has ordered `<Regex>` patterns; first match wins. Scanned PDFs (no selectable text) fall back to OCR via Tesseract if installed.

Post-processing (`_postprocess_result`) normalises amounts and dates, computes `VAT` fallback to `0,00 EUR`, and calculates `vat_rate` fields.

**Available field names:** `InvoiceNumber`, `InvoiceDate`, `GrossAmount`, `NetAmount`, `VAT`, `VATRate`, `SupplierName`, `RecipientName`, `CustomerNumber`, `IBAN`, `PaymentType`, `DueDate`

### Config Files (XML)

#### `invoice_extractor_config*.xml`

- `invoice_extractor_config.xml` — General/default
- `invoice_extractor_config_RE.xml` — Eingangsrechnungen (incoming); emphasises `SupplierName`
- `invoice_extractor_config_RA.xml` — Ausgangsrechnungen (outgoing); emphasises `RecipientName`

Each `<Field>` element supports:
- `name="..."` — Field name used in output/JSON
- `type="date"` / `type="amount"` / `type="vat_rate"` / `type="payment_detection"` — Post-processing behaviour
- `multi="true"` — Collect all regex matches as a list (e.g. IBAN)
- `<XPath>` — ZUGferD namespace path (`rsm:`, `ram:`, `udt:`)
- `<Regex>` — Ordered patterns; first match wins
- `<Keyword category="...">` — For `payment_detection` fields only

AI block in the config (remove to force regex-only mode):
```xml
<AI>
  <Provider>claude</Provider>   <!-- claude | openai | gemini -->
  <Model>claude-opus-4-6</Model>
  <APIKey>sk-ant-...</APIKey>
</AI>
```

#### `invoice_inbox_config.xml`

Copy from `invoice_inbox_config.example.xml`. Key sections:
- `<Mailbox type="exchange"|"imap">` — Credentials, server, folder, limit, `<MarkAsRead>`
- `<Storage>` — `<BaseDir>`, `<Subpath>` (`{year}/{month}`), `<FallbackDir>`
- `<Filename>` — `<Pattern>` with placeholders `{invoice_month}`, `{invoice_day}`, `{supplier}`, `{invoice_number}`; plus `<SupplierMaxLen>`, `<InvoiceNumberMaxLen>`
- `<AttachmentFilter>` — `<SkipPattern>` substrings to skip (e.g. `lastschrift`)
- `<InvoiceFilter>` — `<RequiredField>` names that must be non-empty for a PDF to be filed
- `<InvoiceExtractor>` — `<Config>` path to the extractor config (default: `invoice_extractor_config_RE.xml`)

### Dependencies

- **PyMuPDF** (`fitz`) — PDF parsing and rendering
- **Pillow** — Image handling for scanned-PDF AI extraction
- **pytesseract** — Optional OCR for scanned PDFs (requires Tesseract binary)
- **exchangelib** — Optional Exchange/EWS backend for `inbox_processor`
- **anthropic / openai / google-generativeai** — Optional AI providers

## Agent Framework Files

`AGENTS.md`, `SOUL.md`, `USER.md`, `TOOLS.md`, `HEARTBEAT.md`, `IDENTITY.md` are part of a persistent AI agent framework. Read `AGENTS.md` first if operating as an autonomous agent in this workspace.
