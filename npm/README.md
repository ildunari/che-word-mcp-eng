# che-word-mcp

A Swift-native MCP (Model Context Protocol) server for Microsoft Word (.docx) manipulation. **149 tools**, no runtime dependencies, single binary.

## Quick Start

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "che-word-mcp": {
      "command": "npx",
      "args": ["-y", "che-word-mcp-kosta"]
    }
  }
}
```

### Claude Code

```bash
claude mcp add che-word-mcp -- npx -y che-word-mcp-kosta
```

## Requirements

- **macOS 13.0+** (Ventura or later) — this is a native Swift binary
- GitHub release asset `CheWordMCP` must exist for the matching version, because the npm package downloads that binary during `postinstall`

## What It Does

Create, read, edit, and export Word documents (.docx) directly from Claude or any MCP-compatible AI agent. No Microsoft Word installation needed.

**149 tools** covering: document management, content operations, formatting, range-aware inline editing, tables, styles, lists, page setup, headers/footers, images, export (text & markdown), hyperlinks, bookmarks, comments, track changes, footnotes, endnotes, field codes, and more.

Comment replies use `parent_comment_id` + `text`, and tracked revisions can be listed or accepted/rejected either one at a time or all at once.
`format_text` and `format_text_range` support highlight clearing plus rich run formatting like `strikethrough`, `vertical_align`, `small_caps`, `all_caps`, and `underline_style`, and `false` now clears the new boolean rich-format toggles.
`search_by_formatting` can filter underline presence/style, strikethrough, vertical alignment, and caps in addition to bold/italic/color/highlight.
`replace_text` remains exact-match for typographic variants in this release, so `-` does not match `–`, and straight quotes do not match curly quotes.

## Links

- [Full documentation](https://github.com/ildunari/che-word-mcp-eng)
- [GitHub Releases](https://github.com/ildunari/che-word-mcp-eng/releases)

## License

MIT
