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

## What It Does

Create, read, edit, and export Word documents (.docx) directly from Claude or any MCP-compatible AI agent. No Microsoft Word installation needed.

**149 tools** covering: document management, content operations, formatting, range-aware inline editing, tables, styles, lists, page setup, headers/footers, images, export (text & markdown), hyperlinks, bookmarks, comments, track changes, footnotes, endnotes, field codes, and more.

## Links

- [Full documentation](https://github.com/ildunari/che-word-mcp-eng)
- [GitHub Releases](https://github.com/ildunari/che-word-mcp-eng/releases)

## License

MIT
