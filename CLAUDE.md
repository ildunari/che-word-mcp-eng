# che-word-mcp Development Guide

## Project Structure

```
che-word-mcp-eng/
├── Sources/
│   └── CheWordMCP/
│       └── Server.swift          # MCP server entry point (148 tools)
├── mcpb/                         # MCPB packaging directory
│   ├── manifest.json             # MCPB manifest
│   ├── server/
│   │   └── CheWordMCP            # Built binary (manual copy required)
│   ├── che-word-mcp.mcpb         # Packaged MCPB archive
│   └── README.md
├── Package.swift                 # Swift package definition
├── Package.resolved              # Dependency lock file
├── CHANGELOG.md                  # Version history
├── README.md                     # Main documentation
├── README_zh-TW.md               # Compatibility filename (English content)
└── LICENSE
```

## Important Path Rules

### Binary Install Locations
- **Local development**: `~/bin/CheWordMCP`
- **MCPB packaging**: `mcpb/server/CheWordMCP`

### MCPB Archive Location
- **Correct**: `mcpb/che-word-mcp.mcpb`
- **Incorrect**: repository root (do not place it there)

### Build and Deploy Flow
```bash
# 1) Build
swift build -c release

# 2) Copy binary to install/package locations
cp .build/release/CheWordMCP ~/bin/
cp .build/release/CheWordMCP mcpb/server/

# 3) Package mcpb (run inside mcpb/)
cd mcpb && zip -r che-word-mcp.mcpb . && mv che-word-mcp.mcpb ../mcpb/
# or
cd mcpb && zip -r che-word-mcp.mcpb .
```

## Version Update Checklist

When updating versions, edit:
1. `mcpb/manifest.json` - `version` field
2. `CHANGELOG.md` - add a new release entry
3. `README.md` - tool counts or feature notes if changed

## GitHub Release

When publishing a new version:
```bash
# Create and push tag
git tag v1.x.0
git push origin v1.x.0

# Create release and upload mcpb
gh release create v1.x.0 --title "v1.x.0 - Feature summary" --notes "..."
gh release upload v1.x.0 mcpb/che-word-mcp.mcpb
```

## Related Projects

- **ooxml-swift**: https://github.com/kiki830621/ooxml-swift (core OOXML library)
- **macdoc**: /Users/che/Developer/macdoc (Word->Markdown CLI, delegated target for `export_markdown`)
- **che-claude-plugins**: plugin definitions including this project
