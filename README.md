# DocParser

**Document-to-Structured-Data Converter for LLM Consumption**

Converts PDFs, Word documents, and text files into clean, structured Markdown, JSON, or RAG-ready chunks. Designed to dramatically reduce token waste when feeding documents to AI tools.

## Why This Exists

Raw PDFs and Office documents are designed for human eyes, not machines. A 50-page PDF might consume 80,000+ tokens when pasted into an LLM context window, but the actual information content might only need 15,000 tokens in structured Markdown. DocParser bridges that gap.

## Supported Formats

| Input | Output |
|-------|--------|
| PDF (.pdf) | Markdown (.md) |
| Word (.docx) | JSON (.json) |
| Text (.txt) | RAG Chunks (.chunks.json) |
| Markdown (.md) | |

## Installation

```bash
# Install dependencies (works on Windows, macOS, and Linux)
pip install -r requirements.txt

# That's it. No API keys, no cloud services, runs 100% locally.
```

## Two Ways to Use

### GUI Application (recommended for most users)
```bash
python docparser_gui.py
```
This opens a desktop application where you can:
- Add files individually or entire folders
- Choose output format (Markdown, JSON, or Chunked JSON)
- Set output directory with a folder browser
- Configure chunk size/overlap for RAG pipelines
- Click Process and watch per-file progress with size reduction stats
- Open the output folder directly when done

### Command Line (for scripting / automation)
```bash
# Convert a single PDF to Markdown
python docparser.py report.pdf

# Convert an entire folder
python docparser.py ./documents/ -o ./parsed/

# Generate JSON for programmatic use
python docparser.py report.pdf -f json

# Generate RAG-ready chunks (for vector databases)
python docparser.py ./documents/ -f chunked --chunk-size 500 --overlap 50

# Skip metadata header in output
python docparser.py report.pdf --no-metadata

# Generate a manifest of all processed files
python docparser.py ./documents/ --manifest
```

## Output Formats

### Markdown (default)
Clean, structured Markdown with YAML frontmatter metadata. Headings are preserved from the source document. Tables are converted to Markdown tables. Ideal for direct LLM consumption.

### JSON
Structured JSON with metadata, sections, and tables as separate objects. Useful for building custom pipelines or feeding into databases.

### Chunked JSON
Pre-split content chunks with configurable size and overlap. Ready for embedding into vector databases (Pinecone, Weaviate, ChromaDB, etc.) for RAG pipelines.

## How It Works

1. **Detection** - Identifies file type and selects appropriate parser
2. **Extraction** - Pulls text, headings, tables, and metadata from the source
3. **Structure Analysis** - Uses font sizes (PDF) or style names (DOCX) to identify document hierarchy
4. **Formatting** - Converts to requested output format with clean structure
5. **Output** - Writes files preserving directory structure

### PDF Parsing
Uses PyMuPDF for extraction. Analyzes font sizes and bold/weight properties to identify heading hierarchy. Extracts tables using PyMuPDF's built-in table finder. Handles multi-column layouts and preserves reading order.

### DOCX Parsing
Uses python-docx to read Word XML structure directly. Heading levels come from document styles (Heading 1, Heading 2, etc.). Tables are extracted with full header/row structure.

## Integration Examples

### Feed to Claude API
```python
from pathlib import Path

# Parse document
# python docparser.py briefing.pdf -o briefing.md

content = Path("briefing.md").read_text()

# Now use in API call - content is clean, structured, token-efficient
messages = [{"role": "user", "content": f"Analyze this document:\n\n{content}"}]
```

### Build a RAG Pipeline
```python
import json

# Parse with chunking
# python docparser.py ./docs/ -f chunked --chunk-size 500

chunks = json.loads(Path("parsed_output/report.chunks.json").read_text())

# Each chunk is ready for embedding
for chunk in chunks["chunks"]:
    embedding = your_embedding_model.encode(chunk["content"])
    vector_db.upsert(id=chunk["id"], vector=embedding, metadata={
        "source": chunks["metadata"]["filename"],
        "section": chunk["section_title"],
    })
```

## Token Efficiency

Typical reduction when converting to Markdown:

| Document Type | Raw Tokens | Parsed Tokens | Reduction |
|--------------|-----------|---------------|-----------|
| 20-page PDF report | ~40,000 | ~12,000 | ~70% |
| Complex DOCX with tables | ~25,000 | ~8,000 | ~68% |
| Scanned PDF (text-based) | ~35,000 | ~10,000 | ~71% |

*Actual results vary based on document complexity and formatting density.*

## Limitations

- **Scanned PDFs**: Requires text layer. For image-only scans, you'll need OCR preprocessing (Tesseract, PaddleOCR) before running DocParser.
- **Complex layouts**: Multi-column PDFs are handled but very complex layouts (newspapers, magazines) may need manual review.
- **Images**: Text and tables are extracted; images are noted but not embedded in output.
- **Password-protected files**: Not currently supported.

## Packaging as a Standalone App

If you want to share DocParser with people who don't have Python installed:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name DocParser docparser_gui.py
```

This produces a single executable in the `dist/` folder that runs on the target OS without needing Python.

## Platform Notes

- **Windows**: Works out of the box. Python 3.10+ required. tkinter is included with the standard Windows Python installer.
- **macOS**: Works out of the box. If using Homebrew Python, tkinter may need `brew install python-tk`.
- **Linux**: May need `sudo apt install python3-tk` (Ubuntu/Debian) or `sudo dnf install python3-tkinter` (Fedora) for the GUI version. The CLI version has no tkinter dependency.

## License

MIT - Use however you want.
