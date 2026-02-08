# eml2pst

A Python CLI tool that converts EML email files into PST (Outlook Personal Storage Table) format. Implements the [MS-PST open specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/) from scratch with zero external dependencies.

## Features

- Converts directories of `.eml` files to a single `.pst` file
- Subdirectory structure maps to PST folder hierarchy
- Supports message headers, plain text and HTML bodies, attachments
- Handles multipart MIME messages and charset decoding
- **Stdin/pipe mode** for programmatic integration (JSONL protocol)
- Scales to large mailboxes (tested with 438-message Exchange PST round-trips)
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

### Directory mode

```bash
python -m eml2pst <input_dir> -o output.pst
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

### Stdin/pipe mode (JSONL)

For programmatic integration (databases, web apps, etc.), pipe JSONL to stdin:

```bash
your_app | python -m eml2pst --stdin -o output.pst
```

Each line is a JSON object with a folder path and EML content:

```json
{"folder": "Inbox", "eml": "<base64-encoded EML>"}
{"folder": "Inbox/Projects", "eml": "<base64-encoded EML>"}
{"folder": "Sent Items", "eml": "<base64-encoded EML>"}
```

Nested folders are created automatically from the path. You can also reference files on disk:

```json
{"folder": "Inbox", "eml_file": "/path/to/message.eml"}
```

#### Example: pipe from a script

```bash
python3 -c "
import base64, json
with open('message.eml', 'rb') as f:
    eml_b64 = base64.b64encode(f.read()).decode()
print(json.dumps({'folder': 'Inbox', 'eml': eml_b64}))
" | python -m eml2pst --stdin -o output.pst
```

#### Example: generate from a database

```python
import base64, json, sys, sqlite3

conn = sqlite3.connect('mail.db')
for folder, eml_data in conn.execute('SELECT folder, eml FROM messages'):
    line = json.dumps({
        'folder': folder,
        'eml': base64.b64encode(eml_data).decode()
    })
    print(line)
```

```bash
python3 export_mail.py | python -m eml2pst --stdin -o mailbox.pst
```

### Arguments

| Argument | Description |
|----------|-------------|
| `input_dir` | Directory containing `.eml` files (subdirectories become PST folders) |
| `--stdin` | Read JSONL from stdin instead of a directory |
| `-o, --output` | Output PST file path (default: `output.pst`) |
| `-n, --name` | Display name for the PST store (default: `Personal Folders`) |

## Architecture

The implementation mirrors the three layers of the MS-PST specification:

```
eml2pst/
├── ndb/          # Layer 1: Node Database (blocks, B-trees, allocation maps)
│   ├── header.py     # 544-byte Unicode PST header
│   ├── block.py      # Data blocks (64-byte aligned, 8176 max payload)
│   ├── btree.py      # Multi-level NBT/BBT B-tree pages
│   ├── amap.py       # Allocation Map pages
│   ├── subnode.py    # SLBLOCK subnode lists
│   └── xblock.py     # XBLOCK data trees for multi-block nodes
├── ltp/          # Layer 2: Lists, Tables & Properties (heap, BTH, PC, TC)
│   ├── heap.py       # Multi-page Heap-on-Node allocator
│   ├── bth.py        # BTree-on-Heap
│   ├── pc.py         # Property Context (key-value store)
│   └── tc.py         # Table Context (2D row/column tables)
├── messaging/    # Layer 3: Messaging objects (store, folders, messages)
├── mapi/         # MAPI property tags and type definitions
├── eml_parser.py # EML parsing and MAPI property mapping
├── pst_file.py   # Top-level PST assembly and file writing
├── crc.py        # MS-PST CRC-32 algorithm
├── utils.py      # Encoding helpers, FILETIME conversion
└── cli.py        # CLI entry point (directory + stdin modes)
```

### Key design decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| PST version | Unicode (wVer=23) | Modern format, no 2GB limit |
| Encryption | None (bCryptMethod=0) | Simplest approach |
| Dependencies | stdlib only | Maximum portability |
| B-tree depth | Multi-level (automatic) | Scales to thousands of messages |
| Large values | Subnode + XBLOCK storage | Properties and row data >3580 bytes handled correctly |

## MAPI Properties Mapped from EML

| EML Header | MAPI Property |
|------------|--------------|
| `Subject` | PR_SUBJECT (0x0037) |
| `From` | PR_SENDER_NAME (0x0C1A), PR_SENDER_EMAIL_ADDRESS (0x0C1F) |
| `To` / `Cc` / `Bcc` | Recipients table (NID 0x0692) |
| `Date` | PR_MESSAGE_DELIVERY_TIME (0x0E06) |
| text/plain body | PR_BODY (0x1000) |
| text/html body | PR_HTML (0x1035) |
| MIME attachments | Attachment table (NID 0x0671) + per-attachment PCs |

## Validation

Generated PST files are validated with `readpst` (libpst). Test suite:

- 3-message test (Inbox + Sent, with HTML body, attachments, recipients)
- 438-message Exchange PST round-trip (mock01.pst)
- 26-message multi-folder round-trip with nested subfolders (administrator.pst)

All tests pass with 0 errors, 0 items skipped.

## License

MIT
