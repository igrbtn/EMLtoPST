# eml2pst

A Python CLI tool that converts EML email files into PST (Outlook Personal Storage Table) format. Implements the [MS-PST open specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/) from scratch with zero external dependencies.

## Features

- Converts directories of `.eml` files to a single `.pst` file
- Subdirectory structure maps to PST folder hierarchy
- Supports message headers, plain text and HTML bodies, attachments
- Handles multipart MIME messages and charset decoding
- Unicode PST format (wVer=23) with no 2GB file size limit
- No external dependencies -- only Python stdlib

## Installation

Clone the repository:

```bash
git clone https://github.com/igrbtn/EMLtoPST.git
cd EMLtoPST
```

No `pip install` required. Python 3.8+ is the only prerequisite.

## Usage

```bash
python -m eml2pst <input_dir> -o output.pst
```

### Arguments

| Argument | Description |
|----------|-------------|
| `input_dir` | Directory containing `.eml` files (subdirectories become PST folders) |
| `-o, --output` | Output PST file path (default: `output.pst`) |
| `-n, --name` | Display name for the PST store (default: `Personal Folders`) |

### Example

```bash
# Convert a mailbox export to PST
python -m eml2pst ~/mail-export/ -o mailbox.pst

# Specify a custom store name
python -m eml2pst ~/mail-export/ -o mailbox.pst -n "Work Email"
```

Directory structure is preserved as folders:

```
mail-export/
  Inbox/
    message1.eml
    message2.eml
  Sent/
    reply.eml
```

Produces a PST with `Inbox` (2 messages) and `Sent` (1 message) folders.

## Architecture

The implementation mirrors the three layers of the MS-PST specification:

```
eml2pst/
├── ndb/          # Layer 1: Node Database (blocks, B-trees, allocation maps)
├── ltp/          # Layer 2: Lists, Tables & Properties (heap, BTH, PC, TC)
├── messaging/    # Layer 3: Messaging objects (store, folders, messages)
├── mapi/         # MAPI property tags and type definitions
├── eml_parser.py # EML parsing and MAPI property mapping
├── pst_file.py   # Top-level PST assembly and file writing
├── crc.py        # MS-PST CRC-32 algorithm
├── utils.py      # Encoding helpers, FILETIME conversion
└── cli.py        # CLI entry point
```

### Key design decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| PST version | Unicode (wVer=23) | Modern format, no 2GB limit |
| Encryption | None (bCryptMethod=0) | Simplest approach |
| Dependencies | stdlib only | Maximum portability |
| B-tree depth | Single leaf level | Sufficient for typical mailbox sizes |

## MAPI Properties Mapped from EML

| EML Header | MAPI Property |
|------------|--------------|
| `Subject` | PR_SUBJECT (0x0037) |
| `From` | PR_SENDER_NAME (0x0C1A), PR_SENDER_EMAIL_ADDRESS (0x0C1F) |
| `To` / `Cc` / `Bcc` | Recipients table (NID 0x0692) |
| `Date` | PR_MESSAGE_DELIVERY_TIME (0x0E06) |
| text/plain body | PR_BODY (0x1000) |
| text/html body | PR_HTML (0x1035) |
| MIME attachments | Attachment table (NID 0x0671) |

## Status

The converter generates structurally valid PST files that can be opened in Microsoft Outlook. Core message properties (subject, sender, recipients, body, timestamps) are correctly populated. Attachment support is implemented.

## License

MIT
