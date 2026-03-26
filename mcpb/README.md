# che-word-mcp MCPB Package

This directory contains the MCPB (MCP Bundle) package files for distribution.

## Structure

```
mcpb/
├── manifest.json    # Package metadata
├── server/          # Binary files (Universal Binary)
│   └── CheWordMCP   # The MCP server executable
├── README.md        # This file
└── che-word-mcp.mcpb # Built MCPB archive
```

## Building the Package

```bash
# Build Universal Binary
swift build -c release --build-path .build-arm64 --arch arm64
swift build -c release --build-path .build-x86_64 --arch x86_64
mkdir -p dist/release
lipo -create \
  .build-arm64/release/CheWordMCP \
  .build-x86_64/release/CheWordMCP \
  -output dist/release/CheWordMCP
cp dist/release/CheWordMCP mcpb/server/CheWordMCP

# Create .mcpb package
cd mcpb && zip -r che-word-mcp.mcpb manifest.json README.md server
```

## Installation

The `.mcpb` file can be installed via:

1. **Claude Desktop**: Drag and drop the `.mcpb` file
2. **Manual**: Extract and configure in `claude_desktop_config.json`

## Release Notes

- Current package target: `v1.20.0`
- This bundle ships the rich run-formatting release: `format_text` / `format_text_range` support `strikethrough`, `vertical_align`, `small_caps`, `all_caps`, and `underline_style`
- Keep `manifest.json`, `CHANGELOG.md`, `README.md`, `npm/package.json`, and `Sources/CheWordMCP/Server.swift` on the same version before packaging.
- Upload the raw GitHub release asset as `CheWordMCP` so npm `postinstall` can download it for the matching version.
