# Two Modes: Direct Mode vs Memory Mode

che-word-mcp provides two operation modes. Each tool belongs to one mode only and modes do not mix in a single call chain (see [Disambiguation Principles](../../docs/DISAMBIGUATION.md)).

## Direct Mode (`source_path`)

Pass a `.docx` path directly and complete the operation in one call. You do not need to call `open_document` first.

```
get_text(source_path="/path/to/file.docx")
export_markdown(source_path="/path/to/file.docx", path="/path/to/output.md")
```

### Characteristics

- **Single call**: parse -> process -> return/write output
- **Stateless**: does not keep server memory after completion
- **Read-only**: does not modify the source document

### Tools and Fidelity Tier

| Tool | Tier | Output | Required Parameters |
|------|------|--------|---------------------|
| `get_text` | 1 | Plain text (returned content) | `source_path` |
| `get_document_text` | 1 | Plain text (`get_text` alias) | `source_path` |
| `export_markdown` | 2 | Markdown + images (writes files) | `source_path`, `path` |
| `compare_documents` | - | Difference comparison | `path_a`, `path_b` |

### Performance

```
get_text:        source_path -> DocxReader.read(~0.64s) -> getText(<1ms) -> return text
export_markdown: source_path -> DocxReader.read(~0.64s) -> WordConverter(<1ms) -> write .md + figures/
```

---

## Memory Mode (`doc_id`)

Load a document into memory first with `open_document`, then perform multiple operations through `doc_id`.

```
open_document(path="/path/to/file.docx", doc_id="report")
get_paragraphs(doc_id="report")
insert_paragraph(doc_id="report", text="New paragraph")
format_text(doc_id="report", ...)
save_document(doc_id="report", path="/path/to/output.docx")
close_document(doc_id="report")
```

### Characteristics

- **Stateful**: keeps the document in memory and avoids repeated parsing
- **Multi-step operations**: open once, then run many read/write operations
- **Best for editing**: insert/delete/format/table operations

### Tools

All tools requiring `doc_id`, such as `get_paragraphs`, `insert_paragraph`, `format_text`, `insert_table`, and `save_document`.

### Performance

```
open_document    -> DocxReader.read()   -> keep in memory (~0.64s, one-time)
get_paragraphs   -> read from memory    -> <1ms
insert_paragraph -> mutate in memory    -> <1ms
save_document    -> DocxWriter.write()  -> ~20ms
```

---

## Selection Guide

```
Need to modify the document?
  |- Yes -> Memory Mode (open -> edit -> save)
  `- No  -> What output do you need?
            |- Plain text -> get_text(source_path=...)           Tier 1
            |- Markdown + images -> export_markdown(...)         Tier 2
            `- Diff comparison -> compare_documents(...)
```

| Need | Mode | Tool | Why |
|------|------|------|-----|
| Fast AI reading | Direct | `get_text` | One call, plain text |
| Structured AI reading | Direct | `export_markdown` | Markdown format + images |
| Modify documents | Memory | `open` -> edit -> `save` | Multi-step editing workflow |
| Create new documents | Memory | `create` -> compose -> `save` | Multi-step creation workflow |
| Compare two documents | Direct | `compare_documents` | Direct file-to-file comparison |
