# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This workspace contains **invoice_tools** — a Python toolkit for extracting structured data from PDF invoices and for fetching/filing invoices from email inboxes. It produces output in multiple formats and can be built into a standalone Windows executable.

## Running the Tools

The unified entry point is `invoice_tools.py` (or `invoice_tools.exe` after building):

**Important:** On this system Python is invoked as `py`, not `python3`.

```cmd
:: Extract fields from a single PDF
py invoice_tools.py extractor rechnung.pdf

:: Extract multiple PDFs -> PDF summary
py invoice_tools.py extractor -f pdf -o output.pdf "C:\Pfad\*.pdf"

:: Extract with custom config and debug output
py invoice_tools.py extractor -c invoice_extractor_config_RE.xml -d 2 rechnung.pdf

:: Process inbox (dry-run shows what would be saved)
py invoice_tools.py inbox -m dry          :: dry-run, nur ungelesene Mails
py invoice_tools.py inbox -m dryall       :: dry-run, ALLE Mails (auch gelesene)
py invoice_tools.py inbox -m unread
py invoice_tools.py inbox -m all -c invoice_inbox_config.xml

:: inbox with Excel export
py invoice_tools.py inbox -m dry -e -d 1                      :: dry-run with cell value preview
py invoice_tools.py inbox -m unread -e                         :: file PDFs + write Excel row
py invoice_tools.py inbox -m unread -e -b export               :: + BankingZV
py invoice_tools.py inbox -m dry -B "F:\Buchhaltung"           :: override BaseDir
py invoice_tools.py inbox -m unread -B "F:\Buchhaltung" -e     :: override BaseDir + Excel

:: Help for each sub-tool
py invoice_tools.py extractor --help
py invoice_tools.py inbox --help
```

The individual scripts (`invoice_extractor.py`, `inbox_processor.py`) can also be run directly.

**Output formats:** `stdout` | `json` | `csv` | `xml` | `txt` | `pdf`

**Debug levels:** `0`=off, `1`=extraction path + field count, `2`=text preview + AI prompt/response, `3`=full text + all regex attempts

**Extractor CLI flags:** `-c CONFIG` | `-f FORMAT` | `-o OUTPUT_FILE` | `-d DEBUG_LEVEL` | `-l LOG_FILE`

## Building the Executable (PyInstaller)

```cmd
:: Recommended: combined exe (extractor + inbox in one binary)
py -m PyInstaller invoice_tools.spec -y
:: Output: dist\invoice_tools\invoice_tools.exe
:: Also copied to: src\invoice_tools.exe

:: Individual executables (legacy)
py -m PyInstaller invoice_extractor.spec -y   :: dist\invoice_extractor\invoice_extractor.exe
py -m PyInstaller inbox_processor.spec -y     :: dist\inbox_processor\inbox_processor.exe
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
- `inbox_processor.py` — Connects to Exchange/IMAP, iterates unread mails, extracts PDF attachments via `InvoiceExtractor`, files them, and optionally exports to BankingZV and/or the Excel Rechnungseingang table

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
- `<ExcelExport>` — Optional; activated via `-e / --export-excel` CLI flag
  - `<DateiPfad>` — Full path to the `.xlsx` file
  - `<Tabellenblatt>` — Sheet name (default: `ER`)
  - `<Tabellenname>` — Excel table name for column resolution (default: `tb_rechnungen`)
  - `<StandardTyp>` — Value for the TYP column (default: `E`)
  - `<DuplikatSpalte>` — Column used for duplicate detection (default: `RE-Nr`)
  - `<Spaltenmapping>` — Optional; overrides built-in field mapping. Each `<Spalte>` element:
    - `name="..."` — Excel table column name (looked up dynamically at runtime)
    - `typ="datum|betrag|iban"` — conversion type; omit for plain string
    - element text — source field name (`InvoiceDate`, `GrossAmount`, …) or special value:
      - `{StandardTyp}` — value from `<StandardTyp>` config element
      - `{KontoId}` — account ID resolved via BankingZV routing rules
  - If `<Spaltenmapping>` is absent, the built-in default (`_EXCEL_DEFAULT_MAPPING` in `inbox_processor.py`) is used
  - Formula columns (`calculatedColumnFormula`) are detected automatically and copied from the previous row
  - The table reference (`tbl.ref`) is extended by one row after every successful write

### Excel Export (`inbox_processor.py`)

The Excel export is self-contained within `inbox_processor.py`. Key components:

- **`_ExcelSpalte`** (dataclass) — one mapping entry: `name` (Excel column), `quelle` (source field or `{StandardTyp}`/`{KontoId}`), `typ` (`datum`/`betrag`/`iban`/`""`)
- **`_EXCEL_DEFAULT_MAPPING`** — module-level list of `_ExcelSpalte`; used when no `<Spaltenmapping>` in config
- **`_ExcelKfg`** (dataclass) — full Excel export config including `spaltenmapping: List[_ExcelSpalte]`
- **`_lade_excel_kfg(cfg_root)`** — parses `<ExcelExport>` from inbox config XML
- **`_schreibe_in_excel(fields, konto_id, kfg, dry_run, debug)`** — main write function:
  1. Resolves all values via `_wert()` inner function using the mapping
  2. Opens workbook (always, even in dry-run — needed for duplicate check)
  3. Looks up named table (`ws.tables[kfg.tabellenname]`) and builds `col_map: Dict[str, int]` from `tableColumns`
  4. Detects formula columns via `tc.calculatedColumnFormula is not None`
  5. Runs duplicate check on `kfg.duplikat_spalte` column
  6. In dry-run with `debug > 0`: prints each column name + resolved value
  7. Writes data cells, copies formula cells from `last_row`, extends `tbl.ref`
  8. Suppresses openpyxl `UserWarning` during `load_workbook` via `warnings.catch_warnings()`

### Dependencies

- **PyMuPDF** (`fitz`) — PDF parsing and rendering
- **Pillow** — Image handling for scanned-PDF AI extraction
- **pytesseract** — Optional OCR for scanned PDFs (requires Tesseract binary)
- **exchangelib** — Optional Exchange/EWS backend for `inbox_processor`
- **openpyxl** — Optional Excel export (`-e` flag in `inbox_processor`); must be installed for `<ExcelExport>` to work
- **reportlab** — PDF generation for `make_doku.py` (documentation builder)
- **anthropic / openai / google-generativeai** — Optional AI providers

## Agent Framework Files

`AGENTS.md`, `SOUL.md`, `USER.md`, `TOOLS.md`, `HEARTBEAT.md`, `IDENTITY.md` are part of a persistent AI agent framework. Read `AGENTS.md` first if operating as an autonomous agent in this workspace.
