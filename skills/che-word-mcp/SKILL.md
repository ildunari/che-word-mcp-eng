# che-word-mcp

A Swift-native MCP server for Microsoft Word (.docx) document manipulation. Provides 149 tools for reading, writing, and modifying Word documents without requiring Microsoft Word installation.

## When to Use

Use `che-word-mcp` when you need to:

- Read content from existing .docx files
- Create new Word documents programmatically
- Modify document content, formatting, or structure
- Convert .docx to plain text or Markdown
- Work with tables, images, comments, and other Word features

## Core Workflows

### Reading Documents

```text
1. open_document(path: "/path/to/document.docx", doc_id: "report")
   → Returns document ID
   → Track changes is enabled by default for edits

2. get_text(source_path: "/path/to/document.docx")
   → Returns plain text content

   OR

   get_paragraphs(doc_id: "report")
   → Returns paragraphs with formatting info

3. finalize_document(doc_id: "report")
   → Saves and closes using the original opened path when possible

   Optional check before finalizing:
   get_document_session_state(doc_id: "report")
   → Shows dirty state, save/finalize readiness, and tracking status
```

### Creating Documents

```text
1. create_document(doc_id: "my_document")
   → Returns document ID
   → Track changes is enabled by default for edits

2. insert_paragraph(doc_id: "...", text: "Hello World", style: "Heading1")
   insert_table(doc_id: "...", rows: 3, cols: 4, data: [...])
   insert_image_from_path(doc_id: "...", path: "/path/to/image.png", width: 320, height: 200)

3. finalize_document(doc_id: "...", path: "/path/to/output.docx")
```

### Modifying Documents

```text
1. open_document(path: "/path/to/document.docx", doc_id: "report")

2. update_paragraph(doc_id: "report", index: 0, text: "New content")
   format_text(doc_id: "report", paragraph_index: 0, bold: true, highlight: "yellow")
   format_text_range(doc_id: "report", paragraph_index: 0, start: 5, end: 12, highlight: "none")
   insert_comment(doc_id: "report", paragraph_index: 0, author: "Claude", text: "Review needed")

3. finalize_document(doc_id: "report")
```

### Exporting

```text
export_text(doc_id: "...")
→ Returns plain text

export_markdown(source_path: "/path/to/document.docx", path: "/path/to/output.md")
→ Returns Markdown format
```

### Installation Notes

- GitHub release installs should create `~/bin` first if it does not already exist.
- The current npm package name is `che-word-mcp-kosta`.
- npm installs download the matching GitHub release asset `CheWordMCP` during `postinstall`, so GitHub release assets must already be available.

### Comments and Tracked Revisions

```text
1. list_comments(doc_id: "report")
   → Returns comment IDs for follow-up actions

2. reply_to_comment(doc_id: "report", parent_comment_id: 12, text: "Thanks, updated.", author: "Claude")
   → Adds a reply to an existing comment

3. get_revisions(doc_id: "report")
   → Returns native tracked revisions, including changes that survive save/reopen

4. accept_revision(doc_id: "report", revision_id: 4)
   reject_revision(doc_id: "report", revision_id: 5)
   accept_all_revisions(doc_id: "report")
   reject_all_revisions(doc_id: "report")
```

## Safety Rules

- Do not assume in-memory edits are persisted until `save_document` succeeds.
- Prefer `finalize_document` when the task is done and you want the file safely written and closed in one step.
- If the task is complex or the path behavior is unclear, call `get_document_session_state` before the final save/close step.
- If `close_document` returns an unsaved-changes error, call `save_document` or ask the user whether it should be saved now.
- Prefer `save_document(doc_id: "...")` after `open_document(...)` so the server can reuse the original path safely.
- Use `open_document(..., autosave: true)` only when save-after-each-edit is explicitly desired.
- Paragraph-indexed tools operate on visible paragraphs, so fully deleted tracked paragraph shells do not count toward later indices.
- `replace_text` and `search_text` use exact character matching for punctuation variants in this release: `-` does not match `–`, `'` does not match `’`, and `"` does not match curly double quotes.

## Tool Categories

### Document Management

- `create_document` - Create new document
- `open_document` - Open existing .docx
- `save_document` - Save to file
- `close_document` - Close document
- `list_open_documents` - List all open
- `get_document_info` - Get statistics

### Content

- `get_text` - Get plain text
- `get_paragraphs` - Get with formatting
- `insert_paragraph` - Add paragraph
- `update_paragraph` - Modify paragraph
- `delete_paragraph` - Remove paragraph
- `replace_text` - Find and replace

### Formatting

- `format_text` - Bold, italic, color, font, and run highlight (`highlight: "none"` clears it)
- `format_text_range` - Range-scoped formatting including run highlight (`highlight: "none"` clears it)
- `set_paragraph_format` - Alignment, spacing
- `apply_style` - Apply Word styles
- `set_character_spacing` - Letter spacing
- `set_text_effect` - Text effects

### Tables

- `insert_table` - Create table
- `get_tables` - List tables
- `update_cell` - Modify cell
- `delete_table` - Remove table
- `merge_cells` - Merge cells
- `set_table_style` - Borders, shading

### Images

- `insert_image` - Inline image
- `insert_floating_image` - With text wrap
- `update_image` - Modify properties
- `delete_image` - Remove image
- `list_images` - List all images
- `set_image_style` - Border, effects

### Headers & Footers

- `add_header` / `update_header`
- `add_footer` / `update_footer`
- `insert_page_number`

### Comments & Revisions

- `insert_comment` / `update_comment` / `delete_comment`
- `list_comments` - Get all comments
- `reply_to_comment` - Add reply with `parent_comment_id` + `text`
- `resolve_comment` - Mark resolved
- `enable_track_changes` / `disable_track_changes`
- `get_revisions` - Inspect native tracked revisions, including changes that survive save/reopen
- `accept_revision` / `reject_revision`
- `accept_all_revisions` / `reject_all_revisions`

### Lists

- `insert_bullet_list`
- `insert_numbered_list`
- `set_list_level`

### Page Setup

- `set_page_size` - A4, Letter, etc.
- `set_page_margins`
- `set_page_orientation`
- `insert_page_break`
- `insert_section_break`

### Advanced

- `insert_toc` - Table of contents
- `insert_footnote` / `insert_endnote`
- `insert_hyperlink` / `insert_bookmark`
- `insert_equation` - Math equations
- `insert_checkbox` / `insert_dropdown` - Form fields
- `insert_if_field` / `insert_date_field` - Field codes

## Tips

1. **Always save after modifications** - Changes are in-memory until saved
2. **Close documents when done** - Free up resources
3. **Use styles for consistency** - `apply_style` instead of manual formatting
4. **Check document info first** - Use `get_document_info` to understand structure
5. **Export for AI processing** - Use `export_markdown` for easier text analysis

## Examples

### Create a Report

```text
Create a Word document with:
- Title "Monthly Report" (Heading1)
- Date field that auto-updates
- Executive summary paragraph
- A table with 3 columns: Metric, Value, Change
- Page numbers in footer
Save to ~/Documents/report.docx
```

### Extract and Analyze

```text
Open ~/Documents/thesis.docx
Get all paragraphs with formatting
List all comments
Export as Markdown for analysis
Close the document
```

### Batch Edit Comments

```text
Open the document
List all comments
Reply to each comment with analysis
Mark resolved comments as done
Save the updated document
```
