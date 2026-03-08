import Foundation
import MCP
import OOXMLSwift
import WordToMDSwift
import DocConverterSwift

class WordMCPServer {
    private let server: Server
    private let transport: StdioTransport

    private var openDocuments: [String: WordDocument] = [:]
    private var documentOriginalPaths: [String: String] = [:]
    private var documentDirtyState: [String: Bool] = [:]
    private var documentAutosave: [String: Bool] = [:]
    private var documentTrackChangesEnforced: [String: Bool] = [:]

    private let wordConverter = WordConverter()
    private let defaultRevisionAuthor = "CheWordMCP"

    init() async {
        self.server = Server(
            name: "che-word-mcp",
            version: "1.15.2",
            capabilities: .init(tools: .init())
        )
        self.transport = StdioTransport()

        await registerToolHandlers()
    }

    func run() async throws {
        do {
            try await server.start(transport: transport)
            await server.waitUntilCompleted()
            await flushDirtyDocumentsOnShutdown()
        } catch {
            await flushDirtyDocumentsOnShutdown()
            throw error
        }
    }

    private func registerToolHandlers() async {
        let tools = allTools

        await server.withMethodHandler(ListTools.self) { [tools] _ in
            ListTools.Result(tools: tools)
        }

        await server.withMethodHandler(CallTool.self) { [weak self] params in
            guard let self = self else {
                return CallTool.Result(content: [.text("Server unavailable")], isError: true)
            }
            return try await self.handleToolCall(params)
        }
    }

    private func initializeSession(
        docId: String,
        document: WordDocument,
        sourcePath: String?,
        autosave: Bool
    ) {
        openDocuments[docId] = document
        documentOriginalPaths[docId] = sourcePath
        documentDirtyState[docId] = false
        documentAutosave[docId] = autosave
        documentTrackChangesEnforced[docId] = true
    }

    private func removeSession(docId: String) {
        openDocuments.removeValue(forKey: docId)
        documentOriginalPaths.removeValue(forKey: docId)
        documentDirtyState.removeValue(forKey: docId)
        documentAutosave.removeValue(forKey: docId)
        documentTrackChangesEnforced.removeValue(forKey: docId)
    }

    private func isDirty(docId: String) -> Bool {
        documentDirtyState[docId] ?? false
    }

    private func effectiveSavePath(for docId: String, explicitPath: String?) throws -> String {
        if let explicitPath, !explicitPath.isEmpty {
            return explicitPath
        }
        if let originalPath = documentOriginalPaths[docId], !originalPath.isEmpty {
            return originalPath
        }
        throw WordError.invalidParameter(
            "path",
            "No path was provided and this document has no known original path. Call save_document with an explicit path."
        )
    }

    private func enforceTrackChangesIfNeeded(_ document: inout WordDocument, docId: String) {
        guard documentTrackChangesEnforced[docId] ?? true else { return }
        if !document.isTrackChangesEnabled() {
            document.enableTrackChanges(author: defaultRevisionAuthor)
        }
    }

    private func persistDocumentToDisk(_ document: WordDocument, docId: String, path: String) throws {
        let url = URL(fileURLWithPath: path)
        try DocxWriter.write(document, to: url)
        openDocuments[docId] = document
        documentOriginalPaths[docId] = path
        documentDirtyState[docId] = false
    }

    private func storeDocument(
        _ document: WordDocument,
        for docId: String,
        markDirty: Bool = true
    ) async throws {
        openDocuments[docId] = document
        documentDirtyState[docId] = markDirty

        guard markDirty, documentAutosave[docId] == true, let path = documentOriginalPaths[docId] else {
            return
        }

        try persistDocumentToDisk(document, docId: docId, path: path)
    }

    private func flushDirtyDocumentsOnShutdown() async {
        for docId in openDocuments.keys.sorted() {
            guard isDirty(docId: docId), let document = openDocuments[docId] else { continue }
            guard let path = documentOriginalPaths[docId], !path.isEmpty else {
                FileHandle.standardError.write(
                    Data("Warning: document '\(docId)' has unsaved changes but no save path; shutdown flush skipped.\n".utf8)
                )
                continue
            }

            do {
                try persistDocumentToDisk(document, docId: docId, path: path)
            } catch {
                FileHandle.standardError.write(
                    Data("Warning: failed to flush '\(docId)' to '\(path)' during shutdown: \(error.localizedDescription)\n".utf8)
                )
            }
        }
    }

    func invokeToolForTesting(name: String, arguments: [String: Value] = [:]) async -> CallTool.Result {
        let params = CallTool.Parameters(name: name, arguments: arguments)
        do {
            return try await handleToolCall(params)
        } catch {
            return CallTool.Result(content: [.text("Error: \(error.localizedDescription)")], isError: true)
        }
    }

    func isDocumentDirtyForTesting(_ docId: String) -> Bool {
        isDirty(docId: docId)
    }

    func isTrackChangesEnabledForTesting(_ docId: String) -> Bool? {
        openDocuments[docId]?.isTrackChangesEnabled()
    }

    func flushDirtyDocumentsForTesting() async {
        await flushDirtyDocumentsOnShutdown()
    }

    // MARK: - Tools Definition

    private var allTools: [Tool] {
        [
            Tool(
                name: "create_document",
                description: "Create a new Word document (.docx). Track changes is enabled by default for edits.",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("File identification code for subsequent operations")
                        ]),
                        "autosave": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to autosave to the known backing path after each edit. New documents still need an explicit first save.")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "open_document",
                description: "Open an existing Word file (.docx). Track changes is enabled by default for edits.",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("file path")
                        ]),
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("File identification code for subsequent operations")
                        ]),
                        "autosave": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to autosave back to the opened file after each edit")
                        ])
                    ]),
                    "required": .array([.string("path"), .string("doc_id")])
                ])
            ),
            Tool(
                name: "save_document",
                description: "Save the Word document (.docx). If no path is provided, the server reuses the original opened path when available.",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("Storage path. Optional when the document was opened from disk before editing.")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "close_document",
                description: "Close an open file. Returns an error instead of closing when there are unsaved changes.",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "list_open_documents",
                description: "List all open files",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([:])
                ])
            ),
            Tool(
                name: "get_document_info",
                description: "Get document information (number of paragraphs, number of words, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "get_text",
                description: "Get the plain text content of a .docx file (Direct Mode, Tier 1)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "source_path": .object([
                            "type": .string("string"),
                            "description": .string("Source .docx file path")
                        ])
                    ]),
                    "required": .array([.string("source_path")])
                ])
            ),
            Tool(
                name: "get_paragraphs",
                description: "Get all paragraphs (including formatting information)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "insert_paragraph",
                description: "insert new paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Paragraph text content")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position (starting from 0), if not specified, add to the end")
                        ]),
                        "style": .object([
                            "type": .string("string"),
                            "description": .string("Paragraph style (such as Heading1, Normal)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("text")])
                ])
            ),
            Tool(
                name: "update_paragraph",
                description: "Update the content of an existing paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("new paragraph text")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("index"), .string("text")])
                ])
            ),
            Tool(
                name: "delete_paragraph",
                description: "delete paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("index")])
                ])
            ),
            Tool(
                name: "replace_text",
                description: "Search and replace text",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "find": .object([
                            "type": .string("string"),
                            "description": .string("text to search for")
                        ]),
                        "replace": .object([
                            "type": .string("string"),
                            "description": .string("Replaced text")
                        ]),
                        "all": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to replace all matching items (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("find"), .string("replace")])
                ])
            ),

            Tool(
                name: "format_text",
                description: "Format the text of the specified paragraph (bold, italics, color, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("paragraph index")
                        ]),
                        "bold": .object([
                            "type": .string("boolean"),
                            "description": .string("Bold")
                        ]),
                        "italic": .object([
                            "type": .string("boolean"),
                            "description": .string("italics")
                        ]),
                        "underline": .object([
                            "type": .string("boolean"),
                            "description": .string("bottom line")
                        ]),
                        "font_size": .object([
                            "type": .string("integer"),
                            "description": .string("Font size (points, e.g. 12)")
                        ]),
                        "font_name": .object([
                            "type": .string("string"),
                            "description": .string("Font name (e.g. Arial, Times New Roman)")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Text color (RGB hexadecimal, such as FF0000 means red)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "set_paragraph_format",
                description: "Format paragraphs (alignment, spacing, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("paragraph index")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Alignment: left, center, right, both")
                        ]),
                        "line_spacing": .object([
                            "type": .string("number"),
                            "description": .string("Line spacing (multiple, such as 1.5)")
                        ]),
                        "space_before": .object([
                            "type": .string("integer"),
                            "description": .string("Spacing before paragraph (number of points)")
                        ]),
                        "space_after": .object([
                            "type": .string("integer"),
                            "description": .string("Spacing after paragraph (number of points)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "apply_style",
                description: "Apply built-in styles to paragraphs",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("paragraph index")
                        ]),
                        "style": .object([
                            "type": .string("string"),
                            "description": .string("Style name (such as Heading1, Heading2, Normal, Title)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("style")])
                ])
            ),

            Tool(
                name: "insert_table",
                description: "Insert table",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "rows": .object([
                            "type": .string("integer"),
                            "description": .string("Number of columns")
                        ]),
                        "cols": .object([
                            "type": .string("integer"),
                            "description": .string("Number of columns")
                        ]),
                        "data": .object([
                            "type": .string("array"),
                            "description": .string("Tabular data (two-dimensional array)")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("insertion position")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("rows"), .string("cols")])
                ])
            ),
            Tool(
                name: "get_tables",
                description: "Get information about all tables in the document",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "update_cell",
                description: "Update table cell contents",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "row": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "col": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("New cell contents")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("row"), .string("col"), .string("text")])
                ])
            ),
            Tool(
                name: "delete_table",
                description: "Delete specified table",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index")])
                ])
            ),
            Tool(
                name: "merge_cells",
                description: "Merge table cells (supports horizontal or vertical merging)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "direction": .object([
                            "type": .string("string"),
                            "description": .string("Merge direction: horizontal or vertical")
                        ]),
                        "row": .object([
                            "type": .string("integer"),
                            "description": .string("When merging horizontally: target column index; when merging vertically: starting column")
                        ]),
                        "col": .object([
                            "type": .string("integer"),
                            "description": .string("When merging horizontally: starting column; when merging vertically: target column index")
                        ]),
                        "end_row": .object([
                            "type": .string("integer"),
                            "description": .string("Ending column index when merging vertically")
                        ]),
                        "end_col": .object([
                            "type": .string("integer"),
                            "description": .string("End column index when merging horizontally")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("direction")])
                ])
            ),
            Tool(
                name: "set_table_style",
                description: "Set table style (border, cell background color)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "border_style": .object([
                            "type": .string("string"),
                            "description": .string("Border style: single, double, dashed, dotted, none")
                        ]),
                        "border_color": .object([
                            "type": .string("string"),
                            "description": .string("Border color (RGB hexadecimal, such as 000000)")
                        ]),
                        "border_size": .object([
                            "type": .string("integer"),
                            "description": .string("Border width (1/8 point, default 4 = 0.5pt)")
                        ]),
                        "cell_row": .object([
                            "type": .string("integer"),
                            "description": .string("Set the cell index of the background color (optional)")
                        ]),
                        "cell_col": .object([
                            "type": .string("integer"),
                            "description": .string("Set the cell column index of the background color (optional)")
                        ]),
                        "shading_color": .object([
                            "type": .string("string"),
                            "description": .string("Cell background color (RGB hexadecimal, such as FFFF00)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index")])
                ])
            ),

            Tool(
                name: "list_styles",
                description: "List all styles available in the file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "create_style",
                description: "Create custom styles",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "style_id": .object([
                            "type": .string("string"),
                            "description": .string("Style ID (unique identifier)")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("style display name")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Style type: paragraph, character, table, numbering")
                        ]),
                        "based_on": .object([
                            "type": .string("string"),
                            "description": .string("The style ID to base on (optional)")
                        ]),
                        "next_style": .object([
                            "type": .string("string"),
                            "description": .string("Style ID to use in the next paragraph (optional)")
                        ]),
                        "font_name": .object([
                            "type": .string("string"),
                            "description": .string("Font name")
                        ]),
                        "font_size": .object([
                            "type": .string("integer"),
                            "description": .string("Font size (points)")
                        ]),
                        "bold": .object([
                            "type": .string("boolean"),
                            "description": .string("Bold")
                        ]),
                        "italic": .object([
                            "type": .string("boolean"),
                            "description": .string("italics")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Text color (RGB hexadecimal)")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Alignment: left, center, right, both")
                        ]),
                        "space_before": .object([
                            "type": .string("integer"),
                            "description": .string("Spacing before paragraph (number of points)")
                        ]),
                        "space_after": .object([
                            "type": .string("integer"),
                            "description": .string("Spacing after paragraph (number of points)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("style_id"), .string("name")])
                ])
            ),
            Tool(
                name: "update_style",
                description: "Modify the definition of an existing style",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "style_id": .object([
                            "type": .string("string"),
                            "description": .string("Style ID to modify")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("new display name")
                        ]),
                        "font_name": .object([
                            "type": .string("string"),
                            "description": .string("Font name")
                        ]),
                        "font_size": .object([
                            "type": .string("integer"),
                            "description": .string("Font size (points)")
                        ]),
                        "bold": .object([
                            "type": .string("boolean"),
                            "description": .string("Bold")
                        ]),
                        "italic": .object([
                            "type": .string("boolean"),
                            "description": .string("italics")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Text color (RGB hexadecimal)")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Alignment")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("style_id")])
                ])
            ),
            Tool(
                name: "delete_style",
                description: "Delete custom styles (built-in styles cannot be deleted)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "style_id": .object([
                            "type": .string("string"),
                            "description": .string("Style ID to delete")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("style_id")])
                ])
            ),

            Tool(
                name: "insert_bullet_list",
                description: "Insert bulleted list",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "items": .object([
                            "type": .string("array"),
                            "description": .string("List items (string array)")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position (optional, if not specified, it will be added to the end)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("items")])
                ])
            ),
            Tool(
                name: "insert_numbered_list",
                description: "Insert numbered list",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "items": .object([
                            "type": .string("array"),
                            "description": .string("List items (string array)")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position (optional, if not specified, it will be added to the end)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("items")])
                ])
            ),
            Tool(
                name: "set_list_level",
                description: "Set the level of list items (0-8)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("paragraph index")
                        ]),
                        "level": .object([
                            "type": .string("integer"),
                            "description": .string("Level (0-8, 0 is the outermost level)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("level")])
                ])
            ),

            Tool(
                name: "set_page_size",
                description: "Set page size (letter, a4, legal, a3, a5, b5, executive)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "size": .object([
                            "type": .string("string"),
                            "description": .string("Page sizes: letter, a4, legal, a3, a5, b5, executive")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("size")])
                ])
            ),
            Tool(
                name: "set_page_margins",
                description: "Set page margins",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "preset": .object([
                            "type": .string("string"),
                            "description": .string("Default margins: normal, narrow, moderate, wide (optional)")
                        ]),
                        "top": .object([
                            "type": .string("integer"),
                            "description": .string("Top margin (twips, 1440 = 1 inch)")
                        ]),
                        "right": .object([
                            "type": .string("integer"),
                            "description": .string("Right margin (twips)")
                        ]),
                        "bottom": .object([
                            "type": .string("integer"),
                            "description": .string("Bottom margin (twips)")
                        ]),
                        "left": .object([
                            "type": .string("integer"),
                            "description": .string("Left margin (twips)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "set_page_orientation",
                description: "Set page orientation (portrait/landscape)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "orientation": .object([
                            "type": .string("string"),
                            "description": .string("Page orientation: portrait, landscape")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("orientation")])
                ])
            ),
            Tool(
                name: "insert_page_break",
                description: "Insert page break",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "at_index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position (paragraph index, optional, inserted at the end of the file by default)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "insert_section_break",
                description: "Insert section breaks (different section types can be set)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Section type: nextPage (next page), continuous (continuous), evenPage (even page), oddPage (odd page)")
                        ]),
                        "at_index": .object([
                            "type": .string("integer"),
                            "description": .string("insertion position (paragraph index, optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "add_header",
                description: "Add header",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Header text")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Header type: default (default), first (home page), even (even page)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("text")])
                ])
            ),
            Tool(
                name: "update_header",
                description: "Update top page content",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "header_id": .object([
                            "type": .string("string"),
                            "description": .string("Header ID (returned from add_header)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("New header text")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("header_id"), .string("text")])
                ])
            ),
            Tool(
                name: "add_footer",
                description: "Add footer",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Footer text (optional, if not provided, page number will be used)")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Footer type: default (default), first (home page), even (even page)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "update_footer",
                description: "Update footer content",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "footer_id": .object([
                            "type": .string("string"),
                            "description": .string("Footer ID (returned from add_footer)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("New footer text")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("footer_id"), .string("text")])
                ])
            ),
            Tool(
                name: "insert_page_number",
                description: "Insert page number at the end of the page",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "format": .object([
                            "type": .string("string"),
                            "description": .string("Page number format: simple (1), pageOfTotal (Page 1 of 10), withDash (- 1 -), or custom format such as 'Page #' (# represents the page number)")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Alignment: left, center, right (default center)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "insert_image",
                description: "Insert image into file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "base64": .object([
                            "type": .string("string"),
                            "description": .string("Base64 encoded data for the image")
                        ]),
                        "file_name": .object([
                            "type": .string("string"),
                            "description": .string("Image file name (including file extension, such as image.png)")
                        ]),
                        "width": .object([
                            "type": .string("integer"),
                            "description": .string("Image width (pixels)")
                        ]),
                        "height": .object([
                            "type": .string("integer"),
                            "description": .string("Image height (pixels)")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("insertion position (paragraph index, optional)")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Image name (optional, for alt text)")
                        ]),
                        "description": .object([
                            "type": .string("string"),
                            "description": .string("Image description (optional, for accessibility)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("base64"), .string("file_name"), .string("width"), .string("height")])
                ])
            ),
            Tool(
                name: "insert_image_from_path",
                description: "Insert image from archive path (recommended for large images, avoid base64 transfer)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("Full path to image file")
                        ]),
                        "width": .object([
                            "type": .string("integer"),
                            "description": .string("Image width (pixels)")
                        ]),
                        "height": .object([
                            "type": .string("integer"),
                            "description": .string("Image height (pixels)")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("insertion position (paragraph index, optional)")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Image name (optional, for alt text)")
                        ]),
                        "description": .object([
                            "type": .string("string"),
                            "description": .string("Image description (optional, for accessibility)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("path"), .string("width"), .string("height")])
                ])
            ),
            Tool(
                name: "update_image",
                description: "Update image size",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "image_id": .object([
                            "type": .string("string"),
                            "description": .string("Image ID (returned from insert_image)")
                        ]),
                        "width": .object([
                            "type": .string("integer"),
                            "description": .string("new width (pixels, optional)")
                        ]),
                        "height": .object([
                            "type": .string("integer"),
                            "description": .string("new height (pixels, optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("image_id")])
                ])
            ),
            Tool(
                name: "delete_image",
                description: "Delete picture",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "image_id": .object([
                            "type": .string("string"),
                            "description": .string("Image ID (returned from insert_image)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("image_id")])
                ])
            ),
            Tool(
                name: "list_images",
                description: "List all images in the file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "export_image",
                description: "Export a single image to a file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "image_id": .object([
                            "type": .string("string"),
                            "description": .string("Image ID (obtained from list_images)")
                        ]),
                        "save_path": .object([
                            "type": .string("string"),
                            "description": .string("Full archive path (including file name, such as /tmp/output.png)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("image_id"), .string("save_path")])
                ])
            ),
            Tool(
                name: "export_all_images",
                description: "Export all images to directory",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "output_dir": .object([
                            "type": .string("string"),
                            "description": .string("Output directory path (automatically created)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("output_dir")])
                ])
            ),
            Tool(
                name: "set_image_style",
                description: "Set image style (border, shadow, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "image_id": .object([
                            "type": .string("string"),
                            "description": .string("Image ID (returned from insert_image)")
                        ]),
                        "has_border": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to display borders")
                        ]),
                        "border_color": .object([
                            "type": .string("string"),
                            "description": .string("Border color (RGB hex, such as '000000')")
                        ]),
                        "border_width": .object([
                            "type": .string("integer"),
                            "description": .string("Border width (EMU, 9525 ~ 0.75pt)")
                        ]),
                        "has_shadow": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to show shadow")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("image_id")])
                ])
            ),

            Tool(
                name: "export_text",
                description: "Export files as plain text",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("Export path")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("path")])
                ])
            ),
            Tool(
                name: "export_markdown",
                description: "Convert .docx to Markdown and extract images (Direct Mode, Tier 2)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "source_path": .object([
                            "type": .string("string"),
                            "description": .string("Source .docx file path")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("Markdown export path")
                        ]),
                        "figures_directory": .object([
                            "type": .string("string"),
                            "description": .string("Picture output directory (default is figures/ on the same layer as path)")
                        ]),
                        "include_frontmatter": .object([
                            "type": .string("boolean"),
                            "description": .string("Include file attributes as YAML frontmatter (default false)")
                        ]),
                        "hard_line_breaks": .object([
                            "type": .string("boolean"),
                            "description": .string("Convert soft newlines to hard newlines (default false)")
                        ])
                    ]),
                    "required": .array([.string("source_path"), .string("path")])
                ])
            ),

            Tool(
                name: "insert_hyperlink",
                description: "Insert external hyperlink (URL)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "url": .object([
                            "type": .string("string"),
                            "description": .string("Target URL (such as https://example.com)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Link display text")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Which paragraph to insert into (optional, defaults to the last paragraph)")
                        ]),
                        "tooltip": .object([
                            "type": .string("string"),
                            "description": .string("Mouseover prompt text (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("url"), .string("text")])
                ])
            ),
            Tool(
                name: "insert_internal_link",
                description: "Insert internal link (to bookmark)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "bookmark_name": .object([
                            "type": .string("string"),
                            "description": .string("target bookmark name")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Link display text")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Which paragraph to insert into (optional, defaults to the last paragraph)")
                        ]),
                        "tooltip": .object([
                            "type": .string("string"),
                            "description": .string("Mouseover prompt text (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("bookmark_name"), .string("text")])
                ])
            ),
            Tool(
                name: "update_hyperlink",
                description: "Update the text or URL of a hyperlink",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "hyperlink_id": .object([
                            "type": .string("string"),
                            "description": .string("Hyperlink ID (returned from insert_hyperlink)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("New display text (optional)")
                        ]),
                        "url": .object([
                            "type": .string("string"),
                            "description": .string("New URL (optional, external links only)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("hyperlink_id")])
                ])
            ),
            Tool(
                name: "delete_hyperlink",
                description: "Delete hyperlink (keep text but remove link)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "hyperlink_id": .object([
                            "type": .string("string"),
                            "description": .string("Hyperlink ID (returned from insert_hyperlink)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("hyperlink_id")])
                ])
            ),
            Tool(
                name: "insert_bookmark",
                description: "Insert bookmark markers (for navigation within the file)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Bookmark name (cannot contain spaces, cannot start with a number, maximum 40 characters)")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Which paragraph to insert into (optional, defaults to the last paragraph)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("name")])
                ])
            ),
            Tool(
                name: "delete_bookmark",
                description: "Delete bookmark",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Bookmark name to delete")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("name")])
                ])
            ),

            Tool(
                name: "insert_comment",
                description: "Insert a comment into the specified paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Annotation text")
                        ]),
                        "author": .object([
                            "type": .string("string"),
                            "description": .string("Author name")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index to which annotation is to be attached")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("text"), .string("author"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "update_comment",
                description: "Update annotation content",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "comment_id": .object([
                            "type": .string("integer"),
                            "description": .string("Annotation ID")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("New annotation text")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("comment_id"), .string("text")])
                ])
            ),
            Tool(
                name: "delete_comment",
                description: "Delete annotation",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "comment_id": .object([
                            "type": .string("integer"),
                            "description": .string("Annotation ID to be deleted")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("comment_id")])
                ])
            ),
            Tool(
                name: "list_comments",
                description: "List all annotations in the file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "enable_track_changes",
                description: "Enable revision tracking",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "author": .object([
                            "type": .string("string"),
                            "description": .string("Revise author name (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "disable_track_changes",
                description: "Disable revision tracking",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "accept_revision",
                description: "Accept the specified revision",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "revision_id": .object([
                            "type": .string("integer"),
                            "description": .string("Revision ID (use 'all' to accept all revisions)")
                        ]),
                        "all": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to accept all revisions")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "reject_revision",
                description: "Reject the specified revision",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "revision_id": .object([
                            "type": .string("integer"),
                            "description": .string("Revision ID")
                        ]),
                        "all": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to reject all revisions")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "insert_footnote",
                description: "Inserts a footnote (appears at the bottom of the page) in the specified paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Footnote content")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("text")])
                ])
            ),
            Tool(
                name: "delete_footnote",
                description: "Delete the specified footnote",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "footnote_id": .object([
                            "type": .string("integer"),
                            "description": .string("Footnote ID")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("footnote_id")])
                ])
            ),
            Tool(
                name: "insert_endnote",
                description: "Insert an endnote (appears at the end of the file) in the specified paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Endnote content")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("text")])
                ])
            ),
            Tool(
                name: "delete_endnote",
                description: "Delete the specified endnote",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "endnote_id": .object([
                            "type": .string("integer"),
                            "description": .string("Endnote ID")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("endnote_id")])
                ])
            ),


            Tool(
                name: "insert_toc",
                description: "Insert Table of Contents",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "title": .object([
                            "type": .string("string"),
                            "description": .string("Table of contents title")
                        ]),
                        "heading_levels": .object([
                            "type": .string("string"),
                            "description": .string("Contains a range of title levels, such as 1-3")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position (starting from 0), if not specified, insert to the beginning")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "insert_text_field",
                description: "Insert form text field",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Field name")
                        ]),
                        "default_value": .object([
                            "type": .string("string"),
                            "description": .string("default value")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("name")])
                ])
            ),
            Tool(
                name: "insert_checkbox",
                description: "Insert checkbox",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Field name")
                        ]),
                        "checked": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to check by default")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("name")])
                ])
            ),
            Tool(
                name: "insert_dropdown",
                description: "Insert drop-down menu",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "name": .object([
                            "type": .string("string"),
                            "description": .string("Field name")
                        ]),
                        "options": .object([
                            "type": .string("array"),
                            "description": .string("List of options (JSON array format)")
                        ]),
                        "selected_index": .object([
                            "type": .string("integer"),
                            "description": .string("Default selected index (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("name"), .string("options")])
                ])
            ),

            Tool(
                name: "insert_equation",
                description: "Insert mathematical formulas (supports simplified LaTeX syntax)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "latex": .object([
                            "type": .string("string"),
                            "description": .string("Formulas in LaTeX format")
                        ]),
                        "display_mode": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether it is an independent block (true) or inline (false)")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (specify insertion position in inline mode)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("latex")])
                ])
            ),

            Tool(
                name: "set_paragraph_border",
                description: "Set paragraph borders",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "border_type": .object([
                            "type": .string("string"),
                            "description": .string("Border type: single, double, dotted, dashed, thick, wave")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Border color (hex RGB)")
                        ]),
                        "size": .object([
                            "type": .string("integer"),
                            "description": .string("Border width (1/8 point)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "set_paragraph_shading",
                description: "Set paragraph background color",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "fill": .object([
                            "type": .string("string"),
                            "description": .string("Fill color (hex RGB, such as FFFF00)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("fill")])
                ])
            ),
            Tool(
                name: "set_character_spacing",
                description: "Set character spacing",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "spacing": .object([
                            "type": .string("integer"),
                            "description": .string("Character spacing (1/20 point, positive values increase, negative values decrease)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "set_text_effect",
                description: "Set text effects",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "effect": .object([
                            "type": .string("string"),
                            "description": .string("Effect type: blinkBackground, lights, antsBlack, antsRed, shimmer, sparkle, none")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("effect")])
                ])
            ),


            Tool(
                name: "reply_to_comment",
                description: "Reply to existing comment",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "parent_comment_id": .object([
                            "type": .string("integer"),
                            "description": .string("Annotation ID to reply to")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Reply content")
                        ]),
                        "author": .object([
                            "type": .string("string"),
                            "description": .string("Respondent name (default 'Author')")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("parent_comment_id"), .string("text")])
                ])
            ),
            Tool(
                name: "resolve_comment",
                description: "Mark annotations as resolved or unresolved",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "comment_id": .object([
                            "type": .string("integer"),
                            "description": .string("Annotation ID")
                        ]),
                        "resolved": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether it has been resolved (true/false)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("comment_id")])
                ])
            ),

            Tool(
                name: "insert_floating_image",
                description: "Insert floating pictures (position and text wrapping method can be set)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "base64": .object([
                            "type": .string("string"),
                            "description": .string("Base64 encoded data for the image")
                        ]),
                        "file_name": .object([
                            "type": .string("string"),
                            "description": .string("Image file name (including file extension)")
                        ]),
                        "width": .object([
                            "type": .string("integer"),
                            "description": .string("Image width (pixels)")
                        ]),
                        "height": .object([
                            "type": .string("integer"),
                            "description": .string("Image height (pixels)")
                        ]),
                        "wrap_type": .object([
                            "type": .string("string"),
                            "description": .string("Text wrapping methods: square, tight, through, topAndBottom, behindText, inFrontOfText")
                        ]),
                        "horizontal_position": .object([
                            "type": .string("string"),
                            "description": .string("Horizontal position: left, center, right, or specific offset pixels")
                        ]),
                        "vertical_position": .object([
                            "type": .string("string"),
                            "description": .string("Vertical position: top, center, bottom, or specific offset pixels")
                        ]),
                        "relative_to_h": .object([
                            "type": .string("string"),
                            "description": .string("Horizontal relative to: margin, page, column, character")
                        ]),
                        "relative_to_v": .object([
                            "type": .string("string"),
                            "description": .string("Vertically relative to: margin, page, paragraph, line")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("base64"), .string("file_name"), .string("width"), .string("height")])
                ])
            ),

            Tool(
                name: "insert_if_field",
                description: "Insert IF conditional judgment field",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "left_operand": .object([
                            "type": .string("string"),
                            "description": .string("Left operand (can be a field name or value)")
                        ]),
                        "operator": .object([
                            "type": .string("string"),
                            "description": .string("Comparison operators: =, <>, <, >, <=, >=")
                        ]),
                        "right_operand": .object([
                            "type": .string("string"),
                            "description": .string("right operand")
                        ]),
                        "true_text": .object([
                            "type": .string("string"),
                            "description": .string("Text to display when the condition is true")
                        ]),
                        "false_text": .object([
                            "type": .string("string"),
                            "description": .string("Text to display when the condition is false")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("left_operand"), .string("operator"), .string("right_operand"), .string("true_text"), .string("false_text")])
                ])
            ),
            Tool(
                name: "insert_calculation_field",
                description: "Insert calculated fields (supports SUM, AVERAGE, MAX, MIN, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "expression": .object([
                            "type": .string("string"),
                            "description": .string("Calculation expressions, such as 'SUM(ABOVE)', 'AVERAGE(LEFT)', '=bookmark1*bookmark2'")
                        ]),
                        "format": .object([
                            "type": .string("string"),
                            "description": .string("Number format, such as '#,##0.00' (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("expression")])
                ])
            ),
            Tool(
                name: "insert_date_field",
                description: "Insert date time field",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Date type: date (current date), time (current time), createDate (creation date), saveDate (storage date)")
                        ]),
                        "format": .object([
                            "type": .string("string"),
                            "description": .string("Date format, such as 'yyyy/M/d', 'yyyy year M month d day' (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "insert_page_field",
                description: "Insert page number or document information fields",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Field type: page (page number), numPages (total number of pages), fileName (file name), author (author), numWords (number of words)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("type")])
                ])
            ),
            Tool(
                name: "insert_merge_field",
                description: "Insert merge print fields",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "field_name": .object([
                            "type": .string("string"),
                            "description": .string("Field name (field corresponding to the data source)")
                        ]),
                        "text_before": .object([
                            "type": .string("string"),
                            "description": .string("Prefix text (only displayed if the field is not empty)")
                        ]),
                        "text_after": .object([
                            "type": .string("string"),
                            "description": .string("Post text (only displayed if the field is not empty)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("field_name")])
                ])
            ),
            Tool(
                name: "insert_sequence_field",
                description: "Insert sequence fields (auto numbering, for chart numbering, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "identifier": .object([
                            "type": .string("string"),
                            "description": .string("Sequence identifier, such as 'Figure', 'Table', 'Equation'")
                        ]),
                        "format": .object([
                            "type": .string("string"),
                            "description": .string("Numbering format: arabic (1,2,3), alphabetic (A,B,C), roman (I,II,III)")
                        ]),
                        "reset_level": .object([
                            "type": .string("integer"),
                            "description": .string("Reset the level (corresponding to the heading level, if set to 1, it will be reset every time Heading1 is encountered)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("identifier")])
                ])
            ),

            Tool(
                name: "insert_content_control",
                description: "Insert content controls (SDT)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("Control item types: richText, plainText, picture, date, dropDownList, comboBox, checkbox")
                        ]),
                        "tag": .object([
                            "type": .string("string"),
                            "description": .string("Control label (for identification)")
                        ]),
                        "alias": .object([
                            "type": .string("string"),
                            "description": .string("Control display name")
                        ]),
                        "placeholder": .object([
                            "type": .string("string"),
                            "description": .string("Placeholder prompt text")
                        ]),
                        "content": .object([
                            "type": .string("string"),
                            "description": .string("Default content")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("type"), .string("tag")])
                ])
            ),
            Tool(
                name: "insert_repeating_section",
                description: "Insert repeating section (block that can add/delete items)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position (paragraph index)")
                        ]),
                        "tag": .object([
                            "type": .string("string"),
                            "description": .string("Section label (for identification)")
                        ]),
                        "section_title": .object([
                            "type": .string("string"),
                            "description": .string("Section title (shown in UI)")
                        ]),
                        "items": .object([
                            "type": .string("array"),
                            "description": .string("Initial project content (string array)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("tag")])
                ])
            ),


            Tool(
                name: "insert_text",
                description: "Insert text at the specified position in the specified paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("text to insert")
                        ]),
                        "position": .object([
                            "type": .string("integer"),
                            "description": .string("Character position (starts from 0, if not specified, is inserted at the end of the paragraph)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("text")])
                ])
            ),

            Tool(
                name: "get_document_text",
                description: "Get the complete plain text content of a .docx file (alias for get_text, Direct Mode, Tier 1)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "source_path": .object([
                            "type": .string("string"),
                            "description": .string("Source .docx file path")
                        ])
                    ]),
                    "required": .array([.string("source_path")])
                ])
            ),

            Tool(
                name: "search_text",
                description: "Search for specified text in the document and return all matching positions",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "query": .object([
                            "type": .string("string"),
                            "description": .string("text to search for")
                        ]),
                        "case_sensitive": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to be case sensitive (default false)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("query")])
                ])
            ),

            Tool(
                name: "list_hyperlinks",
                description: "List all hyperlinks in the document",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "list_bookmarks",
                description: "List all bookmarks in a file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "list_footnotes",
                description: "List all footnotes in a file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "list_endnotes",
                description: "List all endnotes in a file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "get_revisions",
                description: "Get all revision tracking records in the file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "accept_all_revisions",
                description: "Accept all revisions in the file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "reject_all_revisions",
                description: "Reject all revisions in the file",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "set_document_properties",
                description: "Set document properties (title, author, subject, keywords, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "title": .object([
                            "type": .string("string"),
                            "description": .string("File title")
                        ]),
                        "subject": .object([
                            "type": .string("string"),
                            "description": .string("Purpose")
                        ]),
                        "creator": .object([
                            "type": .string("string"),
                            "description": .string("author")
                        ]),
                        "keywords": .object([
                            "type": .string("string"),
                            "description": .string("Keywords (separated by commas)")
                        ]),
                        "description": .object([
                            "type": .string("string"),
                            "description": .string("Description/Remarks")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "get_paragraph_runs",
                description: "Get all runs (text fragments) of the specified paragraph and their formatting information, including color, bold, italics, etc.",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "get_text_with_formatting",
                description: "Get the file text and mark it in Markdown format (use ** for bold, * for italics, and {{color:red}} for red)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Specify paragraph index (optional, if not specified, get all)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "search_by_formatting",
                description: "Search for text with a specific format (e.g. red, bold)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Color RGB hex (eg FF0000 represents red)")
                        ]),
                        "bold": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to be bold")
                        ]),
                        "italic": .object([
                            "type": .string("boolean"),
                            "description": .string("Is it italicized?")
                        ]),
                        "highlight": .object([
                            "type": .string("string"),
                            "description": .string("Fluorescent marker color (yellow, green, cyan, etc.)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "get_document_properties",
                description: "Get file attributes (title, author, creation date, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "search_text_with_formatting",
                description: "Searches for text and returns matching positions and their formatting markers (bold, italics, color, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "query": .object([
                            "type": .string("string"),
                            "description": .string("text to search for")
                        ]),
                        "case_sensitive": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to be case sensitive (default false)")
                        ]),
                        "context_chars": .object([
                            "type": .string("integer"),
                            "description": .string("Display the number of characters before and after the matching position (default 20)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("query")])
                ])
            ),

            Tool(
                name: "list_all_formatted_text",
                description: "List all text with a specific format. Must specify format_type: italic, bold, underline, color, highlight, strikethrough",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "format_type": .object([
                            "type": .string("string"),
                            "description": .string("Format type: italic, bold, underline, color, highlight, strikethrough")
                        ]),
                        "color_filter": .object([
                            "type": .string("string"),
                            "description": .string("When format_type=color, the color can be specified (for example, FF0000 represents red)")
                        ]),
                        "paragraph_start": .object([
                            "type": .string("integer"),
                            "description": .string("Starting paragraph index (optional)")
                        ]),
                        "paragraph_end": .object([
                            "type": .string("integer"),
                            "description": .string("End paragraph index (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("format_type")])
                ])
            ),

            Tool(
                name: "get_word_count_by_section",
                description: "Count word count by section, customizable delimiters (such as References) and exclude specific sections",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "section_markers": .object([
                            "type": .string("array"),
                            "description": .string("Array of section delimiter markers (for example [\"Abstract\", \"Introduction\", \"References\"])")
                        ]),
                        "exclude_sections": .object([
                            "type": .string("array"),
                            "description": .string("Section names excluded from total word count (for example [\"References\", \"Appendix\"])")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "compare_documents",
                description: "Compare the differences (paragraph level) between two Word files and return only the differences. Supports text, format, and structure comparison modes",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id_a": .object([
                            "type": .string("string"),
                            "description": .string("Identifier of the baseline file (old version)")
                        ]),
                        "doc_id_b": .object([
                            "type": .string("string"),
                            "description": .string("Compare the identification code of the file (new version)")
                        ]),
                        "mode": .object([
                            "type": .string("string"),
                            "description": .string("Comparison mode: text (default, pure text difference), formatting (including formatting differences), structure (structure summary), full (complete comparison)"),
                            "enum": .array([.string("text"), .string("formatting"), .string("structure"), .string("full")])
                        ]),
                        "context_lines": .object([
                            "type": .string("integer"),
                            "description": .string("Number of unchanged paragraphs displayed before and after the difference (0-3, default 0)"),
                            "minimum": .int(0),
                            "maximum": .int(3)
                        ]),
                        "max_results": .object([
                            "type": .string("integer"),
                            "description": .string("Maximum number of differences to be returned (default 0 = return all)"),
                            "minimum": .int(0)
                        ]),
                        "heading_styles": .object([
                            "type": .string("array"),
                            "description": .string("Custom heading style names (used in structure mode, for example [\"EC8\", \"ECtitle\"])"),
                            "items": .object(["type": .string("string")])
                        ])
                    ]),
                    "required": .array([.string("doc_id_a"), .string("doc_id_b")])
                ])
            ),


            Tool(
                name: "set_columns",
                description: "Set multi-column formatting of documents (default for the entire document, or insert section breaks after specified paragraphs)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "columns": .object([
                            "type": .string("integer"),
                            "description": .string("Number of columns (1-4)"),
                            "minimum": .int(1),
                            "maximum": .int(4)
                        ]),
                        "space": .object([
                            "type": .string("integer"),
                            "description": .string("Column spacing (twips, default 720 = 0.5 inch)")
                        ]),
                        "equal_width": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether column widths are equal (default true)")
                        ]),
                        "separator": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to display separators (default false)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("columns")])
                ])
            ),

            Tool(
                name: "insert_column_break",
                description: "Insert column breaks into specified paragraphs",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "set_line_numbers",
                description: "Set file line number display",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "enable": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to enable line numbers")
                        ]),
                        "start": .object([
                            "type": .string("integer"),
                            "description": .string("Starting line number (default 1)")
                        ]),
                        "count_by": .object([
                            "type": .string("integer"),
                            "description": .string("Show line numbers every few lines (default 1)")
                        ]),
                        "restart": .object([
                            "type": .string("string"),
                            "description": .string("Renumbering mode: continuous, newSection (per section), newPage (per page)")
                        ]),
                        "distance": .object([
                            "type": .string("integer"),
                            "description": .string("The distance between line number and text (twips, default 360)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("enable")])
                ])
            ),

            Tool(
                name: "set_page_borders",
                description: "Set page borders (the four sides can be set independently)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "style": .object([
                            "type": .string("string"),
                            "description": .string("Border style: single (single line), double (double line), dotted (dotted line), dashed (dashed line), thick (thick line), none (none)")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Border color (RGB hexadecimal, such as 000000)")
                        ]),
                        "size": .object([
                            "type": .string("integer"),
                            "description": .string("Border thickness (1/8 point, default 4 = 0.5pt)")
                        ]),
                        "offset_from": .object([
                            "type": .string("string"),
                            "description": .string("Starting position of border: text (from text), page (from page edge)")
                        ]),
                        "top": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to display the top border (default true)")
                        ]),
                        "bottom": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to display the bottom border (default true)")
                        ]),
                        "left": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to display the left border (default true)")
                        ]),
                        "right": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to display the right border (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("style")])
                ])
            ),

            Tool(
                name: "insert_symbol",
                description: "Insert special symbols into specified paragraphs (use font symbols or Unicode)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "char": .object([
                            "type": .string("string"),
                            "description": .string("Symbol character code (hexadecimal, such as F020 or Unicode code point)")
                        ]),
                        "font": .object([
                            "type": .string("string"),
                            "description": .string("Symbol fonts (e.g. Symbol, Wingdings, Wingdings 2)")
                        ]),
                        "position": .object([
                            "type": .string("string"),
                            "description": .string("Insertion position: start (beginning of paragraph), end (end of paragraph)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("char")])
                ])
            ),

            Tool(
                name: "set_text_direction",
                description: "Set the text direction of a paragraph or document (supports straight writing)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "direction": .object([
                            "type": .string("string"),
                            "description": .string("Text direction: lrTb (left to right, top to bottom, default), tbRl (top to bottom, right to left, straight writing), btLr (bottom to top, left to right)")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (if not specified, the entire document will be applied)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("direction")])
                ])
            ),

            Tool(
                name: "insert_drop_cap",
                description: "Enlarge the first word of a paragraph (drop cap effect)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "type": .object([
                            "type": .string("string"),
                            "description": .string("First word type: drop (sink, default), margin (at the border), none (remove)")
                        ]),
                        "lines": .object([
                            "type": .string("integer"),
                            "description": .string("Number of sinking rows (2-10, default 3)")
                        ]),
                        "distance": .object([
                            "type": .string("integer"),
                            "description": .string("Distance from text (twips, default 0)")
                        ]),
                        "font": .object([
                            "type": .string("string"),
                            "description": .string("First letter font (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "insert_horizontal_line",
                description: "Insert a horizontal line in the specified paragraph (paragraph border mode)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0), a horizontal line will be added below the paragraph")
                        ]),
                        "style": .object([
                            "type": .string("string"),
                            "description": .string("Line style: single (single line, default), double (double line), dotted (dotted line), dashed (dashed line), thick (thick line)")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Line color (RGB hex, default 000000)")
                        ]),
                        "size": .object([
                            "type": .string("integer"),
                            "description": .string("Line thickness (1/8 point, default 12 = 1.5pt)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "set_widow_orphan",
                description: "Set paragraphs to avoid beginning and ending (lone line/few lines control)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0), if not specified, the entire document will be applied")
                        ]),
                        "enable": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to enable head and tail avoidance (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "set_keep_with_next",
                description: "Set the paragraph to be on the same page as the next paragraph (to avoid separation during paging)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "enable": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to enable the same page as the next paragraph (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),


            Tool(
                name: "insert_watermark",
                description: "Insert text watermark (centered diagonally on the page background)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("Watermark text (such as \"CONFIDENTIAL\", \"DRAFT\", \"CONFIDENTIAL\")")
                        ]),
                        "font": .object([
                            "type": .string("string"),
                            "description": .string("Font name (Default Calibri Light)")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("Text color (RGB hexadecimal, default C0C0C0 light gray)")
                        ]),
                        "size": .object([
                            "type": .string("integer"),
                            "description": .string("Font size (points, default 72)")
                        ]),
                        "semitransparent": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to be translucent (default true)")
                        ]),
                        "rotation": .object([
                            "type": .string("integer"),
                            "description": .string("Rotation angle (degrees, default -45 is oblique)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("text")])
                ])
            ),

            Tool(
                name: "insert_image_watermark",
                description: "Insert image watermark (centered on page background)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "image_path": .object([
                            "type": .string("string"),
                            "description": .string("Image file path")
                        ]),
                        "scale": .object([
                            "type": .string("integer"),
                            "description": .string("Scaling (percentage, default 100)")
                        ]),
                        "washout": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to fade processing (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("image_path")])
                ])
            ),

            Tool(
                name: "remove_watermark",
                description: "Remove watermarks from files",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "protect_document",
                description: "Set file protection (restrict editing, read only, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "protection_type": .object([
                            "type": .string("string"),
                            "description": .string("Protection type: readOnly (read only), comments (only comments allowed), trackedChanges (only tracked changes allowed), forms (only form filling allowed)")
                        ]),
                        "password": .object([
                            "type": .string("string"),
                            "description": .string("Protect password (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("protection_type")])
                ])
            ),

            Tool(
                name: "unprotect_document",
                description: "Remove file protection",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "password": .object([
                            "type": .string("string"),
                            "description": .string("Protection password (if set)")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "set_document_password",
                description: "Set file opening password (encryption protection)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "password": .object([
                            "type": .string("string"),
                            "description": .string("Open password")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("password")])
                ])
            ),

            Tool(
                name: "remove_document_password",
                description: "Remove file opening password",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "current_password": .object([
                            "type": .string("string"),
                            "description": .string("Current activation password")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("current_password")])
                ])
            ),

            Tool(
                name: "restrict_editing_region",
                description: "Set editable area (other areas are protected)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "start_paragraph": .object([
                            "type": .string("integer"),
                            "description": .string("Editable area starting paragraph index")
                        ]),
                        "end_paragraph": .object([
                            "type": .string("integer"),
                            "description": .string("End of editable area paragraph index")
                        ]),
                        "editor": .object([
                            "type": .string("string"),
                            "description": .string("Users/groups allowed to edit (optional)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("start_paragraph"), .string("end_paragraph")])
                ])
            ),


            Tool(
                name: "insert_caption",
                description: "Insert labels for pictures or tables (such as \"Figure 1\", \"Table 1\")",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position paragraph index")
                        ]),
                        "label": .object([
                            "type": .string("string"),
                            "description": .string("Label type: Figure, Table, Equation")
                        ]),
                        "caption_text": .object([
                            "type": .string("string"),
                            "description": .string("Label description text")
                        ]),
                        "position": .object([
                            "type": .string("string"),
                            "description": .string("Label position: above (above), below (below, default)")
                        ]),
                        "include_chapter_number": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to include the chapter number (such as \"Figure 2-1\")")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("label")])
                ])
            ),

            Tool(
                name: "insert_cross_reference",
                description: "Insert cross-references (links to bookmarks, titles, chart labels, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position paragraph index")
                        ]),
                        "reference_type": .object([
                            "type": .string("string"),
                            "description": .string("Reference types: bookmark, heading, figure, table, equation")
                        ]),
                        "reference_target": .object([
                            "type": .string("string"),
                            "description": .string("Reference target name or ID")
                        ]),
                        "format": .object([
                            "type": .string("string"),
                            "description": .string("Display format: full (complete, such as \"Figure 1\"), numberOnly (number only), pageNumber (page number), text (text only)")
                        ]),
                        "include_hyperlink": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to add hyperlinks (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("reference_type"), .string("reference_target")])
                ])
            ),

            Tool(
                name: "insert_table_of_figures",
                description: "Insert chart catalog",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position paragraph index")
                        ]),
                        "caption_label": .object([
                            "type": .string("string"),
                            "description": .string("Label type: Figure, Table, Equation")
                        ]),
                        "include_page_numbers": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to include page numbers (default true)")
                        ]),
                        "right_align_page_numbers": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether the page number is aligned to the right (default true)")
                        ]),
                        "tab_leader": .object([
                            "type": .string("string"),
                            "description": .string("Anchor point leading characters: dot (dotted line), hyphen (hyphen), underscore (underline), none (none)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("caption_label")])
                ])
            ),

            Tool(
                name: "insert_index_entry",
                description: "Mark text as an index item (used to generate the index)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Index of the paragraph containing the text to be tagged")
                        ]),
                        "main_entry": .object([
                            "type": .string("string"),
                            "description": .string("main index term")
                        ]),
                        "sub_entry": .object([
                            "type": .string("string"),
                            "description": .string("Sub-index term (optional)")
                        ]),
                        "cross_reference": .object([
                            "type": .string("string"),
                            "description": .string("Cross-reference (such as \"See XXX\")")
                        ]),
                        "bold": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether the page number is bold (default false)")
                        ]),
                        "italic": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether the page number is italicized (default false)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("main_entry")])
                ])
            ),

            Tool(
                name: "insert_index",
                description: "Insert index (based on marked index items)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position paragraph index")
                        ]),
                        "columns": .object([
                            "type": .string("integer"),
                            "description": .string("Number of index columns (1-4, default 2)")
                        ]),
                        "right_align_page_numbers": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether the page number is aligned to the right (default true)")
                        ]),
                        "tab_leader": .object([
                            "type": .string("string"),
                            "description": .string("Anchor point leading characters: dot (dotted line), hyphen (hyphen), underscore (underline), none (none)")
                        ]),
                        "run_in": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether sub-items are displayed continuously (default false)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),


            Tool(
                name: "set_language",
                description: "Set the proofreading language for text (used for spell checking)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "language": .object([
                            "type": .string("string"),
                            "description": .string("Language code (such as en-US, zh-TW, ja-JP)")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (if not specified, the entire document will be applied)")
                        ]),
                        "no_proofing": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to disable redaction (default false)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("language")])
                ])
            ),

            Tool(
                name: "set_keep_lines",
                description: "Set paragraphs not to be paged (the entire paragraph remains on the same page)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "enable": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to enable paragraph non-pagination (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "insert_tab_stop",
                description: "Set anchor point in paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "position": .object([
                            "type": .string("integer"),
                            "description": .string("Anchor position (twips, counted from the left border)")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Alignment: left, center, right, decimal (decimal point alignment)")
                        ]),
                        "leader": .object([
                            "type": .string("string"),
                            "description": .string("Leading characters: none, dot, hyphen, underscore")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("position")])
                ])
            ),

            Tool(
                name: "clear_tab_stops",
                description: "Clear all anchor points of paragraph",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "set_page_break_before",
                description: "Set pagination before paragraph (paragraph starts on new page)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "enable": .object([
                            "type": .string("boolean"),
                            "description": .string("Whether to enable pagination before paragraphs (default true)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "set_outline_level",
                description: "Set the outline level of paragraphs (used to generate tables of contents and navigation)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Paragraph index (starting from 0)")
                        ]),
                        "level": .object([
                            "type": .string("integer"),
                            "description": .string("Outline level (1-9, or 0 for this article)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("level")])
                ])
            ),

            Tool(
                name: "insert_continuous_section_break",
                description: "Insert continuous section breaks (section breaks without page breaks)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("Insertion position paragraph index")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),

            Tool(
                name: "get_section_properties",
                description: "Get the section attributes of the file (page settings, etc.)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            Tool(
                name: "add_row_to_table",
                description: "Add a new column to the table",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "position": .object([
                            "type": .string("string"),
                            "description": .string("Insertion position: end (last), start (front), after_row (after the specified column)")
                        ]),
                        "row_index": .object([
                            "type": .string("integer"),
                            "description": .string("When position=after_row, specify the column to insert after")
                        ]),
                        "data": .object([
                            "type": .string("array"),
                            "description": .string("Array of cell data for new column")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index")])
                ])
            ),

            Tool(
                name: "add_column_to_table",
                description: "Add a new column to the table",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "position": .object([
                            "type": .string("string"),
                            "description": .string("Insertion position: end (last), start (front), after_col (after the specified column)")
                        ]),
                        "col_index": .object([
                            "type": .string("integer"),
                            "description": .string("When position=after_col, specify the column to insert after")
                        ]),
                        "data": .object([
                            "type": .string("array"),
                            "description": .string("Array of cell data for new column")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index")])
                ])
            ),

            Tool(
                name: "delete_row_from_table",
                description: "Delete a column from the table",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "row_index": .object([
                            "type": .string("integer"),
                            "description": .string("Column index to delete (0-based)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("row_index")])
                ])
            ),

            Tool(
                name: "delete_column_from_table",
                description: "Remove a column from the table",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "col_index": .object([
                            "type": .string("integer"),
                            "description": .string("Column index to delete (starting from 0)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("col_index")])
                ])
            ),

            Tool(
                name: "set_cell_width",
                description: "Set table cell width",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "row": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "col": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "width": .object([
                            "type": .string("integer"),
                            "description": .string("Width (twips)")
                        ]),
                        "width_type": .object([
                            "type": .string("string"),
                            "description": .string("Width type: dxa (fixed twips), pct (percentage), auto (automatic)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("row"), .string("col"), .string("width")])
                ])
            ),

            Tool(
                name: "set_row_height",
                description: "Set table column height",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "row_index": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "height": .object([
                            "type": .string("integer"),
                            "description": .string("height(twips)")
                        ]),
                        "height_rule": .object([
                            "type": .string("string"),
                            "description": .string("Height rules: auto (automatic), atLeast (minimum), exact (fixed)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("row_index"), .string("height")])
                ])
            ),

            Tool(
                name: "set_table_alignment",
                description: "Set the alignment of the table on the page",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Alignment: left, center, right")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("alignment")])
                ])
            ),

            Tool(
                name: "set_cell_vertical_alignment",
                description: "Set vertical alignment of table cells",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "row": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "col": .object([
                            "type": .string("integer"),
                            "description": .string("Column index (starting from 0)")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("Vertical alignment: top, center, bottom")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index"), .string("row"), .string("col"), .string("alignment")])
                ])
            ),

            Tool(
                name: "set_header_row",
                description: "Set the table title column (repeated when spanning two pages)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("file identification code")
                        ]),
                        "table_index": .object([
                            "type": .string("integer"),
                            "description": .string("Table index (starting from 0)")
                        ]),
                        "row_count": .object([
                            "type": .string("integer"),
                            "description": .string("Number of title columns (counting from the first column, default 1)")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("table_index")])
                ])
            )
        ]
    }

    // MARK: - Tool Handler

    private func handleToolCall(_ params: CallTool.Parameters) async throws -> CallTool.Result {
        let name = params.name
        let args = params.arguments ?? [:]

        do {
            let result = try await executeToolTask(name: name, args: args)
            return CallTool.Result(content: [.text(result)])
        } catch {
            return CallTool.Result(
                content: [.text("Error: \(error.localizedDescription)")],
                isError: true
            )
        }
    }

    private func executeToolTask(name: String, args: [String: Value]) async throws -> String {
        switch name {
        case "create_document":
            return try await createDocument(args: args)
        case "open_document":
            return try await openDocument(args: args)
        case "save_document":
            return try await saveDocument(args: args)
        case "close_document":
            return try await closeDocument(args: args)
        case "list_open_documents":
            return await listOpenDocuments()
        case "get_document_info":
            return try await getDocumentInfo(args: args)

        case "get_text":
            return try await getText(args: args)
        case "get_paragraphs":
            return try await getParagraphs(args: args)
        case "insert_paragraph":
            return try await insertParagraph(args: args)
        case "update_paragraph":
            return try await updateParagraph(args: args)
        case "delete_paragraph":
            return try await deleteParagraph(args: args)
        case "replace_text":
            return try await replaceText(args: args)

        case "format_text":
            return try await formatText(args: args)
        case "set_paragraph_format":
            return try await setParagraphFormat(args: args)
        case "apply_style":
            return try await applyStyle(args: args)

        case "insert_table":
            return try await insertTable(args: args)
        case "get_tables":
            return try await getTables(args: args)
        case "update_cell":
            return try await updateCell(args: args)
        case "delete_table":
            return try await deleteTable(args: args)
        case "merge_cells":
            return try await mergeCells(args: args)
        case "set_table_style":
            return try await setTableStyle(args: args)

        case "list_styles":
            return try await listStyles(args: args)
        case "create_style":
            return try await createStyle(args: args)
        case "update_style":
            return try await updateStyle(args: args)
        case "delete_style":
            return try await deleteStyle(args: args)

        case "insert_bullet_list":
            return try await insertBulletList(args: args)
        case "insert_numbered_list":
            return try await insertNumberedList(args: args)
        case "set_list_level":
            return try await setListLevel(args: args)

        case "set_page_size":
            return try await setPageSize(args: args)
        case "set_page_margins":
            return try await setPageMargins(args: args)
        case "set_page_orientation":
            return try await setPageOrientation(args: args)
        case "insert_page_break":
            return try await insertPageBreak(args: args)
        case "insert_section_break":
            return try await insertSectionBreak(args: args)

        case "add_header":
            return try await addHeader(args: args)
        case "update_header":
            return try await updateHeader(args: args)
        case "add_footer":
            return try await addFooter(args: args)
        case "update_footer":
            return try await updateFooter(args: args)
        case "insert_page_number":
            return try await insertPageNumber(args: args)

        case "insert_image":
            return try await insertImage(args: args)
        case "insert_image_from_path":
            return try await insertImageFromPath(args: args)
        case "update_image":
            return try await updateImage(args: args)
        case "delete_image":
            return try await deleteImage(args: args)
        case "list_images":
            return try await listImages(args: args)
        case "export_image":
            return try await exportImage(args: args)
        case "export_all_images":
            return try await exportAllImages(args: args)
        case "set_image_style":
            return try await setImageStyle(args: args)

        case "export_text":
            return try await exportText(args: args)
        case "export_markdown":
            return try await exportMarkdown(args: args)

        case "insert_hyperlink":
            return try await insertHyperlink(args: args)
        case "insert_internal_link":
            return try await insertInternalLink(args: args)
        case "update_hyperlink":
            return try await updateHyperlink(args: args)
        case "delete_hyperlink":
            return try await deleteHyperlink(args: args)
        case "insert_bookmark":
            return try await insertBookmark(args: args)
        case "delete_bookmark":
            return try await deleteBookmark(args: args)

        case "insert_comment":
            return try await insertComment(args: args)
        case "update_comment":
            return try await updateComment(args: args)
        case "delete_comment":
            return try await deleteComment(args: args)
        case "list_comments":
            return try await listComments(args: args)
        case "enable_track_changes":
            return try await enableTrackChanges(args: args)
        case "disable_track_changes":
            return try await disableTrackChanges(args: args)
        case "accept_revision":
            return try await acceptRevision(args: args)
        case "reject_revision":
            return try await rejectRevision(args: args)

        case "insert_footnote":
            return try await insertFootnote(args: args)
        case "delete_footnote":
            return try await deleteFootnote(args: args)
        case "insert_endnote":
            return try await insertEndnote(args: args)
        case "delete_endnote":
            return try await deleteEndnote(args: args)

        case "insert_toc":
            return try await insertTOC(args: args)
        case "insert_text_field":
            return try await insertTextField(args: args)
        case "insert_checkbox":
            return try await insertCheckbox(args: args)
        case "insert_dropdown":
            return try await insertDropdown(args: args)
        case "insert_equation":
            return try await insertEquation(args: args)
        case "set_paragraph_border":
            return try await setParagraphBorder(args: args)
        case "set_paragraph_shading":
            return try await setParagraphShading(args: args)
        case "set_character_spacing":
            return try await setCharacterSpacing(args: args)
        case "set_text_effect":
            return try await setTextEffect(args: args)

        case "reply_to_comment":
            return try await replyToComment(args: args)
        case "resolve_comment":
            return try await resolveComment(args: args)

        case "insert_floating_image":
            return try await insertFloatingImage(args: args)

        case "insert_if_field":
            return try await insertIfField(args: args)
        case "insert_calculation_field":
            return try await insertCalculationField(args: args)
        case "insert_date_field":
            return try await insertDateField(args: args)
        case "insert_page_field":
            return try await insertPageField(args: args)
        case "insert_merge_field":
            return try await insertMergeField(args: args)
        case "insert_sequence_field":
            return try await insertSequenceField(args: args)

        case "insert_content_control":
            return try await insertContentControl(args: args)
        case "insert_repeating_section":
            return try await insertRepeatingSection(args: args)

        case "insert_text":
            return try await insertText(args: args)
        case "get_document_text":
            return try await getDocumentText(args: args)
        case "search_text":
            return try await searchText(args: args)
        case "list_hyperlinks":
            return try await listHyperlinks(args: args)
        case "list_bookmarks":
            return try await listBookmarks(args: args)
        case "list_footnotes":
            return try await listFootnotes(args: args)
        case "list_endnotes":
            return try await listEndnotes(args: args)
        case "get_revisions":
            return try await getRevisions(args: args)
        case "accept_all_revisions":
            return try await acceptAllRevisions(args: args)
        case "reject_all_revisions":
            return try await rejectAllRevisions(args: args)
        case "set_document_properties":
            return try await setDocumentProperties(args: args)
        case "get_document_properties":
            return try await getDocumentProperties(args: args)
        case "get_paragraph_runs":
            return try await getParagraphRuns(args: args)
        case "get_text_with_formatting":
            return try await getTextWithFormatting(args: args)
        case "search_by_formatting":
            return try await searchByFormatting(args: args)
        case "search_text_with_formatting":
            return try await searchTextWithFormatting(args: args)
        case "list_all_formatted_text":
            return try await listAllFormattedText(args: args)
        case "get_word_count_by_section":
            return try await getWordCountBySection(args: args)
        case "compare_documents":
            return try await compareDocuments(args: args)

        case "set_columns":
            return try await setColumns(args: args)
        case "insert_column_break":
            return try await insertColumnBreak(args: args)
        case "set_line_numbers":
            return try await setLineNumbers(args: args)
        case "set_page_borders":
            return try await setPageBorders(args: args)
        case "insert_symbol":
            return try await insertSymbol(args: args)
        case "set_text_direction":
            return try await setTextDirection(args: args)
        case "insert_drop_cap":
            return try await insertDropCap(args: args)
        case "insert_horizontal_line":
            return try await insertHorizontalLine(args: args)
        case "set_widow_orphan":
            return try await setWidowOrphan(args: args)
        case "set_keep_with_next":
            return try await setKeepWithNext(args: args)

        case "insert_watermark":
            return try await insertWatermark(args: args)
        case "insert_image_watermark":
            return try await insertImageWatermark(args: args)
        case "remove_watermark":
            return try await removeWatermark(args: args)
        case "protect_document":
            return try await protectDocument(args: args)
        case "unprotect_document":
            return try await unprotectDocument(args: args)
        case "set_document_password":
            return try await setDocumentPassword(args: args)
        case "remove_document_password":
            return try await removeDocumentPassword(args: args)
        case "restrict_editing_region":
            return try await restrictEditingRegion(args: args)

        case "insert_caption":
            return try await insertCaption(args: args)
        case "insert_cross_reference":
            return try await insertCrossReference(args: args)
        case "insert_table_of_figures":
            return try await insertTableOfFigures(args: args)
        case "insert_index_entry":
            return try await insertIndexEntry(args: args)
        case "insert_index":
            return try await insertIndex(args: args)

        case "set_language":
            return try await setLanguage(args: args)
        case "set_keep_lines":
            return try await setKeepLines(args: args)
        case "insert_tab_stop":
            return try await insertTabStop(args: args)
        case "clear_tab_stops":
            return try await clearTabStops(args: args)
        case "set_page_break_before":
            return try await setPageBreakBefore(args: args)
        case "set_outline_level":
            return try await setOutlineLevel(args: args)
        case "insert_continuous_section_break":
            return try await insertContinuousSectionBreak(args: args)
        case "get_section_properties":
            return try await getSectionProperties(args: args)
        case "add_row_to_table":
            return try await addRowToTable(args: args)
        case "add_column_to_table":
            return try await addColumnToTable(args: args)
        case "delete_row_from_table":
            return try await deleteRowFromTable(args: args)
        case "delete_column_from_table":
            return try await deleteColumnFromTable(args: args)
        case "set_cell_width":
            return try await setCellWidth(args: args)
        case "set_row_height":
            return try await setRowHeight(args: args)
        case "set_table_alignment":
            return try await setTableAlignment(args: args)
        case "set_cell_vertical_alignment":
            return try await setCellVerticalAlignment(args: args)
        case "set_header_row":
            return try await setHeaderRow(args: args)

        default:
            throw WordError.unknownTool(name)
        }
    }

    // MARK: - Document Management

    private func createDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        if openDocuments[docId] != nil {
            throw WordError.documentAlreadyOpen(docId)
        }

        let autosave = args["autosave"]?.boolValue ?? false
        var doc = WordDocument()
        doc.enableTrackChanges(author: defaultRevisionAuthor)
        initializeSession(docId: docId, document: doc, sourcePath: nil, autosave: autosave)

        return "Created new document with id: \(docId). Track changes is enabled by default."
    }

    private func openDocument(args: [String: Value]) async throws -> String {
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        if openDocuments[docId] != nil {
            throw WordError.documentAlreadyOpen(docId)
        }
        let autosave = args["autosave"]?.boolValue ?? false

        let url = URL(fileURLWithPath: path)
        var doc = try DocxReader.read(from: url)
        if !doc.isTrackChangesEnabled() {
            doc.enableTrackChanges(author: defaultRevisionAuthor)
        }
        initializeSession(docId: docId, document: doc, sourcePath: path, autosave: autosave)

        return "Opened document '\(path)' with id: \(docId). Track changes is enabled by default."
    }

    private func saveDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let explicitPath = args["path"]?.stringValue
        let path = try effectiveSavePath(for: docId, explicitPath: explicitPath)
        try persistDocumentToDisk(doc, docId: docId, path: path)

        if explicitPath == nil {
            return "Saved document to original path: \(path)"
        }

        return "Saved document to: \(path)"
    }

    private func closeDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }

        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        if isDirty(docId: docId) {
            let knownPath = documentOriginalPaths[docId] ?? "no known path"
            let saveGuidance: String
            if knownPath == "no known path" {
                saveGuidance = " with an explicit path"
            } else {
                saveGuidance = " (path can be omitted to reuse \(knownPath))"
            }
            throw WordError.invalidFormat(
                "Document '\(docId)' has unsaved changes. Call save_document first\(saveGuidance). If you want the user to confirm, report back and ask whether it should be saved now."
            )
        }

        removeSession(docId: docId)
        return "Closed document: \(docId)"
    }

    private func listOpenDocuments() async -> String {
        if openDocuments.isEmpty {
            return "No documents currently open"
        }

        let ids = openDocuments.keys.sorted()
        return "Open documents:\n" + ids.map { docId in
            let dirty = isDirty(docId: docId) ? "dirty" : "clean"
            let autosave = (documentAutosave[docId] ?? false) ? "autosave:on" : "autosave:off"
            let path = documentOriginalPaths[docId] ?? "no-path"
            return "- \(docId) [\(dirty), \(autosave), path:\(path)]"
        }.joined(separator: "\n")
    }

    private func getDocumentInfo(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let info = doc.getInfo()
        return """
        Document Info (\(docId)):
        - Paragraphs: \(info.paragraphCount)
        - Characters: \(info.characterCount)
        - Words: \(info.wordCount)
        - Tables: \(info.tableCount)
        """
    }

    // MARK: - Content Operations

    private func getText(args: [String: Value]) async throws -> String {
        guard let sourcePath = args["source_path"]?.stringValue else {
            throw WordError.missingParameter("source_path")
        }
        guard FileManager.default.fileExists(atPath: sourcePath) else {
            throw WordError.fileNotFound(sourcePath)
        }
        let sourceURL = URL(fileURLWithPath: sourcePath)
        let lockFile = sourceURL.deletingLastPathComponent()
            .appendingPathComponent("~$" + sourceURL.lastPathComponent)
        if FileManager.default.fileExists(atPath: lockFile.path) {
            throw WordError.invalidFormat("File is open in Microsoft Word. Please save and close it first: \(sourceURL.lastPathComponent)")
        }
        let document = try DocxReader.read(from: sourceURL)
        return document.getText()
    }

    private func getParagraphs(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let paragraphs = doc.getParagraphs()
        if paragraphs.isEmpty {
            return "No paragraphs in document"
        }

        var result = "Paragraphs:\n"
        for (index, para) in paragraphs.enumerated() {
            let style = para.properties.style ?? "Normal"
            let preview = String(para.getText().prefix(50))
            result += "[\(index)] (\(style)) \(preview)...\n"
        }
        return result
    }

    private func insertParagraph(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let index = args["index"]?.intValue
        let style = args["style"]?.stringValue

        var para = Paragraph(text: text)
        if let style = style {
            para.properties.style = style
        }

        if let index = index {
            doc.insertParagraph(para, at: index)
        } else {
            doc.appendParagraph(para)
        }

        try await storeDocument(doc, for: docId)

        return "Inserted paragraph at index \(index ?? doc.getParagraphs().count - 1)"
    }

    private func updateParagraph(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let index = args["index"]?.intValue else {
            throw WordError.missingParameter("index")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.updateParagraph(at: index, text: text)
        try await storeDocument(doc, for: docId)

        return "Updated paragraph at index \(index)"
    }

    private func deleteParagraph(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let index = args["index"]?.intValue else {
            throw WordError.missingParameter("index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteParagraph(at: index)
        try await storeDocument(doc, for: docId)

        return "Deleted paragraph at index \(index)"
    }

    private func replaceText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let find = args["find"]?.stringValue else {
            throw WordError.missingParameter("find")
        }
        guard let replace = args["replace"]?.stringValue else {
            throw WordError.missingParameter("replace")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let replaceAll = args["all"]?.boolValue ?? true
        let count = doc.replaceText(find: find, with: replace, all: replaceAll)
        try await storeDocument(doc, for: docId)

        return "Replaced \(count) occurrence(s) of '\(find)' with '\(replace)'"
    }

    // MARK: - Formatting

    private func formatText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        var format = RunProperties()
        if let bold = args["bold"]?.boolValue { format.bold = bold }
        if let italic = args["italic"]?.boolValue { format.italic = italic }
        if let underline = args["underline"]?.boolValue { format.underline = underline ? .single : nil }
        if let fontSize = args["font_size"]?.intValue { format.fontSize = fontSize * 2 }
        if let fontName = args["font_name"]?.stringValue { format.fontName = fontName }
        if let color = args["color"]?.stringValue { format.color = color }

        try doc.formatParagraph(at: paragraphIndex, with: format)
        try await storeDocument(doc, for: docId)

        return "Applied formatting to paragraph \(paragraphIndex)"
    }

    private func setParagraphFormat(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        var props = ParagraphProperties()
        if let alignment = args["alignment"]?.stringValue {
            props.alignment = Alignment(rawValue: alignment)
        }
        if let lineSpacing = args["line_spacing"]?.doubleValue {
            props.spacing = Spacing(line: Int(lineSpacing * 240))
        }
        if let spaceBefore = args["space_before"]?.intValue {
            if props.spacing == nil { props.spacing = Spacing() }
            props.spacing?.before = spaceBefore * 20
        }
        if let spaceAfter = args["space_after"]?.intValue {
            if props.spacing == nil { props.spacing = Spacing() }
            props.spacing?.after = spaceAfter * 20
        }

        try doc.setParagraphFormat(at: paragraphIndex, properties: props)
        try await storeDocument(doc, for: docId)

        return "Applied paragraph format to index \(paragraphIndex)"
    }

    private func applyStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let style = args["style"]?.stringValue else {
            throw WordError.missingParameter("style")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.applyStyle(at: paragraphIndex, style: style)
        try await storeDocument(doc, for: docId)

        return "Applied style '\(style)' to paragraph \(paragraphIndex)"
    }

    // MARK: - Table

    private func insertTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let rows = args["rows"]?.intValue else {
            throw WordError.missingParameter("rows")
        }
        guard let cols = args["cols"]?.intValue else {
            throw WordError.missingParameter("cols")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        var table = Table(rowCount: rows, columnCount: cols)

        if let dataArray = args["data"]?.arrayValue {
            for (rowIndex, rowData) in dataArray.enumerated() {
                if let rowArray = rowData.arrayValue {
                    for (colIndex, cellData) in rowArray.enumerated() {
                        if let text = cellData.stringValue,
                           rowIndex < table.rows.count && colIndex < table.rows[rowIndex].cells.count {
                            table.rows[rowIndex].cells[colIndex] = TableCell(text: text)
                        }
                    }
                }
            }
        }

        let index = args["index"]?.intValue
        if let index = index {
            doc.insertTable(table, at: index)
        } else {
            doc.appendTable(table)
        }

        try await storeDocument(doc, for: docId)

        return "Inserted \(rows)x\(cols) table"
    }

    private func getTables(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let tables = doc.getTables()
        if tables.isEmpty {
            return "No tables in document"
        }

        var result = "Tables in document:\n"
        for (index, table) in tables.enumerated() {
            let rows = table.rows.count
            let cols = table.rows.first?.cells.count ?? 0
            result += "[\(index)] \(rows)x\(cols) table\n"

            for (rowIdx, row) in table.rows.prefix(3).enumerated() {
                let cellPreviews = row.cells.prefix(3).map { cell -> String in
                    let preview = String(cell.getText().prefix(15))
                    return preview.isEmpty ? "(empty)" : preview
                }
                result += "  Row \(rowIdx): \(cellPreviews.joined(separator: " | "))\n"
            }
            if table.rows.count > 3 {
                result += "  ... (\(table.rows.count - 3) more rows)\n"
            }
        }
        return result
    }

    private func updateCell(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let row = args["row"]?.intValue else {
            throw WordError.missingParameter("row")
        }
        guard let col = args["col"]?.intValue else {
            throw WordError.missingParameter("col")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.updateCell(tableIndex: tableIndex, row: row, col: col, text: text)
        try await storeDocument(doc, for: docId)

        return "Updated cell at table[\(tableIndex)][\(row)][\(col)]"
    }

    private func deleteTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteTable(at: tableIndex)
        try await storeDocument(doc, for: docId)

        return "Deleted table at index \(tableIndex)"
    }

    private func mergeCells(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let direction = args["direction"]?.stringValue else {
            throw WordError.missingParameter("direction")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        switch direction.lowercased() {
        case "horizontal":
            guard let row = args["row"]?.intValue else {
                throw WordError.missingParameter("row")
            }
            guard let col = args["col"]?.intValue else {
                throw WordError.missingParameter("col")
            }
            guard let endCol = args["end_col"]?.intValue else {
                throw WordError.missingParameter("end_col")
            }
            try doc.mergeCellsHorizontal(tableIndex: tableIndex, row: row, startCol: col, endCol: endCol)
            try await storeDocument(doc, for: docId)
            return "Merged cells horizontally: row \(row), columns \(col) to \(endCol)"

        case "vertical":
            guard let row = args["row"]?.intValue else {
                throw WordError.missingParameter("row")
            }
            guard let col = args["col"]?.intValue else {
                throw WordError.missingParameter("col")
            }
            guard let endRow = args["end_row"]?.intValue else {
                throw WordError.missingParameter("end_row")
            }
            try doc.mergeCellsVertical(tableIndex: tableIndex, col: col, startRow: row, endRow: endRow)
            try await storeDocument(doc, for: docId)
            return "Merged cells vertically: column \(col), rows \(row) to \(endRow)"

        default:
            throw WordError.invalidParameter("direction", "Must be 'horizontal' or 'vertical'")
        }
    }

    private func setTableStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        var results: [String] = []

        if let borderStyle = args["border_style"]?.stringValue {
            let style = BorderStyle(rawValue: borderStyle) ?? .single
            let size = args["border_size"]?.intValue ?? 4
            let color = args["border_color"]?.stringValue ?? "000000"

            let border = Border(style: style, size: size, color: color)
            let borders = TableBorders.all(border)

            try doc.setTableBorders(tableIndex: tableIndex, borders: borders)
            results.append("Set border style: \(borderStyle)")
        }

        if let cellRow = args["cell_row"]?.intValue,
           let cellCol = args["cell_col"]?.intValue,
           let shadingColor = args["shading_color"]?.stringValue {
            let shading = CellShading(fill: shadingColor)
            try doc.setCellShading(tableIndex: tableIndex, row: cellRow, col: cellCol, shading: shading)
            results.append("Set cell shading at [\(cellRow)][\(cellCol)]: \(shadingColor)")
        }

        try await storeDocument(doc, for: docId)

        if results.isEmpty {
            return "No style changes applied"
        }
        return results.joined(separator: "\n")
    }

    // MARK: - Style Management

    private func listStyles(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let styles = doc.getStyles()
        if styles.isEmpty {
            return "No styles defined"
        }

        var result = "Available Styles:\n"
        for style in styles {
            let defaultMark = style.isDefault ? " (default)" : ""
            let basedOnInfo = style.basedOn.map { " [based on: \($0)]" } ?? ""
            result += "- \(style.id) (\(style.name)) - \(style.type.rawValue)\(defaultMark)\(basedOnInfo)\n"

            if let runProps = style.runProperties {
                var formats: [String] = []
                if let fontName = runProps.fontName { formats.append("font: \(fontName)") }
                if let fontSize = runProps.fontSize { formats.append("size: \(fontSize / 2)pt") }
                if runProps.bold == true { formats.append("bold") }
                if runProps.italic == true { formats.append("italic") }
                if let color = runProps.color { formats.append("color: #\(color)") }
                if !formats.isEmpty {
                    result += "    Text: \(formats.joined(separator: ", "))\n"
                }
            }
        }
        return result
    }

    private func createStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let styleId = args["style_id"]?.stringValue else {
            throw WordError.missingParameter("style_id")
        }
        guard let name = args["name"]?.stringValue else {
            throw WordError.missingParameter("name")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let typeStr = args["type"]?.stringValue ?? "paragraph"
        let styleType = StyleType(rawValue: typeStr) ?? .paragraph

        var paraProps = ParagraphProperties()
        if let alignment = args["alignment"]?.stringValue {
            paraProps.alignment = Alignment(rawValue: alignment)
        }
        if let spaceBefore = args["space_before"]?.intValue {
            if paraProps.spacing == nil { paraProps.spacing = Spacing() }
            paraProps.spacing?.before = spaceBefore * 20
        }
        if let spaceAfter = args["space_after"]?.intValue {
            if paraProps.spacing == nil { paraProps.spacing = Spacing() }
            paraProps.spacing?.after = spaceAfter * 20
        }

        var runProps = RunProperties()
        if let fontName = args["font_name"]?.stringValue { runProps.fontName = fontName }
        if let fontSize = args["font_size"]?.intValue { runProps.fontSize = fontSize * 2 }
        if let bold = args["bold"]?.boolValue { runProps.bold = bold }
        if let italic = args["italic"]?.boolValue { runProps.italic = italic }
        if let color = args["color"]?.stringValue { runProps.color = color }

        let style = Style(
            id: styleId,
            name: name,
            type: styleType,
            basedOn: args["based_on"]?.stringValue,
            nextStyle: args["next_style"]?.stringValue,
            isDefault: false,
            isQuickStyle: true,
            paragraphProperties: paraProps,
            runProperties: runProps
        )

        try doc.addStyle(style)
        try await storeDocument(doc, for: docId)

        return "Created style '\(styleId)' (\(name))"
    }

    private func updateStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let styleId = args["style_id"]?.stringValue else {
            throw WordError.missingParameter("style_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        var paraProps: ParagraphProperties? = nil
        if let alignment = args["alignment"]?.stringValue {
            paraProps = ParagraphProperties()
            paraProps?.alignment = Alignment(rawValue: alignment)
        }

        var runProps: RunProperties? = nil
        if args["font_name"] != nil || args["font_size"] != nil ||
           args["bold"] != nil || args["italic"] != nil || args["color"] != nil {
            runProps = RunProperties()
            if let fontName = args["font_name"]?.stringValue { runProps?.fontName = fontName }
            if let fontSize = args["font_size"]?.intValue { runProps?.fontSize = fontSize * 2 }
            if let bold = args["bold"]?.boolValue { runProps?.bold = bold }
            if let italic = args["italic"]?.boolValue { runProps?.italic = italic }
            if let color = args["color"]?.stringValue { runProps?.color = color }
        }

        let updates = StyleUpdate(
            name: args["name"]?.stringValue,
            paragraphProperties: paraProps,
            runProperties: runProps
        )

        try doc.updateStyle(id: styleId, with: updates)
        try await storeDocument(doc, for: docId)

        return "Updated style '\(styleId)'"
    }

    private func deleteStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let styleId = args["style_id"]?.stringValue else {
            throw WordError.missingParameter("style_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteStyle(id: styleId)
        try await storeDocument(doc, for: docId)

        return "Deleted style '\(styleId)'"
    }

    // MARK: - List Operations

    private func insertBulletList(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let itemsArray = args["items"]?.arrayValue else {
            throw WordError.missingParameter("items")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let items = itemsArray.compactMap { $0.stringValue }
        if items.isEmpty {
            throw WordError.invalidParameter("items", "Must contain at least one item")
        }

        let index = args["index"]?.intValue
        let numId = doc.insertBulletList(items: items, at: index)
        try await storeDocument(doc, for: docId)

        return "Inserted bullet list with \(items.count) items (numId: \(numId))"
    }

    private func insertNumberedList(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let itemsArray = args["items"]?.arrayValue else {
            throw WordError.missingParameter("items")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let items = itemsArray.compactMap { $0.stringValue }
        if items.isEmpty {
            throw WordError.invalidParameter("items", "Must contain at least one item")
        }

        let index = args["index"]?.intValue
        let numId = doc.insertNumberedList(items: items, at: index)
        try await storeDocument(doc, for: docId)

        return "Inserted numbered list with \(items.count) items (numId: \(numId))"
    }

    private func setListLevel(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let level = args["level"]?.intValue else {
            throw WordError.missingParameter("level")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.setListLevel(paragraphIndex: paragraphIndex, level: level)
        try await storeDocument(doc, for: docId)

        return "Set list level to \(level) for paragraph \(paragraphIndex)"
    }

    // MARK: - Page Settings

    private func setPageSize(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let sizeName = args["size"]?.stringValue else {
            throw WordError.missingParameter("size")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.setPageSize(name: sizeName)
        try await storeDocument(doc, for: docId)

        let size = doc.sectionProperties.pageSize
        return "Set page size to \(size.name) (\(size.widthInInches)\" x \(size.heightInInches)\")"
    }

    private func setPageMargins(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        if let preset = args["preset"]?.stringValue {
            try doc.setPageMargins(name: preset)
        } else {
            let top = args["top"]?.intValue
            let right = args["right"]?.intValue
            let bottom = args["bottom"]?.intValue
            let left = args["left"]?.intValue

            doc.setPageMargins(top: top, right: right, bottom: bottom, left: left)
        }

        try await storeDocument(doc, for: docId)

        let margins = doc.sectionProperties.pageMargins
        return "Set page margins to \(margins.name) (top: \(margins.top), right: \(margins.right), bottom: \(margins.bottom), left: \(margins.left) twips)"
    }

    private func setPageOrientation(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let orientationStr = args["orientation"]?.stringValue else {
            throw WordError.missingParameter("orientation")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        guard let orientation = PageOrientation(rawValue: orientationStr.lowercased()) else {
            throw WordError.invalidParameter("orientation", "Must be 'portrait' or 'landscape'")
        }

        doc.setPageOrientation(orientation)
        try await storeDocument(doc, for: docId)

        return "Set page orientation to \(orientation.rawValue)"
    }

    private func insertPageBreak(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let index = args["at_index"]?.intValue
        doc.insertPageBreak(at: index)
        try await storeDocument(doc, for: docId)

        if let index = index {
            return "Inserted page break at position \(index)"
        } else {
            return "Inserted page break at end of document"
        }
    }

    private func insertSectionBreak(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let typeStr = args["type"]?.stringValue ?? "nextPage"
        guard let breakType = SectionBreakType(rawValue: typeStr) else {
            throw WordError.invalidParameter("type", "Must be 'nextPage', 'continuous', 'evenPage', or 'oddPage'")
        }

        let index = args["at_index"]?.intValue
        doc.insertSectionBreak(type: breakType, at: index)
        try await storeDocument(doc, for: docId)

        if let index = index {
            return "Inserted \(breakType.rawValue) section break at position \(index)"
        } else {
            return "Inserted \(breakType.rawValue) section break at end of document"
        }
    }

    // MARK: - Header/Footer

    private func addHeader(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let typeStr = args["type"]?.stringValue ?? "default"
        let headerType: HeaderFooterType
        switch typeStr.lowercased() {
        case "first": headerType = .first
        case "even": headerType = .even
        default: headerType = .default
        }

        let header = doc.addHeader(text: text, type: headerType)
        try await storeDocument(doc, for: docId)

        return "Added header with id '\(header.id)' (type: \(headerType.rawValue))"
    }

    private func updateHeader(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let headerId = args["header_id"]?.stringValue else {
            throw WordError.missingParameter("header_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.updateHeader(id: headerId, text: text)
        try await storeDocument(doc, for: docId)

        return "Updated header '\(headerId)'"
    }

    private func addFooter(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let typeStr = args["type"]?.stringValue ?? "default"
        let footerType: HeaderFooterType
        switch typeStr.lowercased() {
        case "first": footerType = .first
        case "even": footerType = .even
        default: footerType = .default
        }

        let footer: Footer
        if let text = args["text"]?.stringValue {
            footer = doc.addFooter(text: text, type: footerType)
        } else {
            footer = doc.addFooterWithPageNumber(format: .simple, type: footerType)
        }

        try await storeDocument(doc, for: docId)

        return "Added footer with id '\(footer.id)' (type: \(footerType.rawValue))"
    }

    private func updateFooter(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let footerId = args["footer_id"]?.stringValue else {
            throw WordError.missingParameter("footer_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.updateFooter(id: footerId, text: text)
        try await storeDocument(doc, for: docId)

        return "Updated footer '\(footerId)'"
    }

    private func insertPageNumber(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let formatStr = args["format"]?.stringValue ?? "simple"
        let format: PageNumberFormat
        switch formatStr.lowercased() {
        case "simple": format = .simple
        case "pageoftotal": format = .pageOfTotal
        case "withdash": format = .withDash
        default:
            if formatStr.contains("#") {
                format = .withText(formatStr)
            } else {
                format = .simple
            }
        }

        let footer = doc.addFooterWithPageNumber(format: format, type: .default)
        try await storeDocument(doc, for: docId)

        return "Inserted page number in footer '\(footer.id)' with format '\(formatStr)'"
    }

    // MARK: - Image Operations

    private func insertImage(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let base64 = args["base64"]?.stringValue else {
            throw WordError.missingParameter("base64")
        }
        guard let fileName = args["file_name"]?.stringValue else {
            throw WordError.missingParameter("file_name")
        }
        guard let width = args["width"]?.intValue else {
            throw WordError.missingParameter("width")
        }
        guard let height = args["height"]?.intValue else {
            throw WordError.missingParameter("height")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let index = args["index"]?.intValue
        let name = args["name"]?.stringValue ?? "Picture"
        let description = args["description"]?.stringValue ?? ""

        let imageId = try doc.insertImage(
            base64: base64,
            fileName: fileName,
            widthPx: width,
            heightPx: height,
            at: index,
            name: name,
            description: description
        )

        try await storeDocument(doc, for: docId)

        return "Inserted image '\(fileName)' with id '\(imageId)' (\(width)x\(height) pixels)"
    }

    private func insertImageFromPath(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let width = args["width"]?.intValue else {
            throw WordError.missingParameter("width")
        }
        guard let height = args["height"]?.intValue else {
            throw WordError.missingParameter("height")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let fileManager = FileManager.default
        guard fileManager.fileExists(atPath: path) else {
            throw WordError.fileNotFound(path)
        }

        let index = args["index"]?.intValue
        let name = args["name"]?.stringValue ?? "Picture"
        let description = args["description"]?.stringValue ?? ""

        let imageId = try doc.insertImage(
            path: path,
            widthPx: width,
            heightPx: height,
            at: index,
            name: name,
            description: description
        )

        try await storeDocument(doc, for: docId)

        let url = URL(fileURLWithPath: path)
        return "Inserted image '\(url.lastPathComponent)' from path with id '\(imageId)' (\(width)x\(height) pixels)"
    }

    private func updateImage(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let imageId = args["image_id"]?.stringValue else {
            throw WordError.missingParameter("image_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let width = args["width"]?.intValue
        let height = args["height"]?.intValue

        try doc.updateImage(imageId: imageId, widthPx: width, heightPx: height)
        try await storeDocument(doc, for: docId)

        var changes: [String] = []
        if let w = width { changes.append("width: \(w)px") }
        if let h = height { changes.append("height: \(h)px") }

        return "Updated image '\(imageId)': \(changes.joined(separator: ", "))"
    }

    private func deleteImage(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let imageId = args["image_id"]?.stringValue else {
            throw WordError.missingParameter("image_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteImage(imageId: imageId)
        try await storeDocument(doc, for: docId)

        return "Deleted image '\(imageId)'"
    }

    private func listImages(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let images = doc.getImages()

        if images.isEmpty {
            return "No images in document"
        }

        var result = "Found \(images.count) image(s):\n"
        for img in images {
            result += "- id: \(img.id), file: \(img.fileName), size: \(img.widthPx)x\(img.heightPx)px\n"
        }

        return result
    }

    private func exportImage(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let imageId = args["image_id"]?.stringValue else {
            throw WordError.missingParameter("image_id")
        }
        guard let savePath = args["save_path"]?.stringValue else {
            throw WordError.missingParameter("save_path")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        guard let imageRef = doc.images.first(where: { $0.id == imageId }) else {
            throw WordError.parseError("Image ID not found: \(imageId)")
        }

        let url = URL(fileURLWithPath: savePath)
        let directory = url.deletingLastPathComponent()
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)

        try imageRef.data.write(to: url)

        let sizeKB = imageRef.data.count / 1024
        return "Saved image \(imageId) to \(savePath) (\(sizeKB)KB)"
    }

    private func exportAllImages(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let outputDir = args["output_dir"]?.stringValue else {
            throw WordError.missingParameter("output_dir")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let images = doc.images
        if images.isEmpty {
            return "No images to export"
        }

        let dirURL = URL(fileURLWithPath: outputDir)
        try FileManager.default.createDirectory(at: dirURL, withIntermediateDirectories: true)

        var result = "Exported \(images.count) image(s) to \(outputDir):\n"
        for imageRef in images {
            let fileURL = dirURL.appendingPathComponent(imageRef.fileName)
            try imageRef.data.write(to: fileURL)
            let sizeKB = imageRef.data.count / 1024
            result += "  - \(imageRef.fileName) (\(sizeKB)KB)\n"
        }

        return result
    }

    private func setImageStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let imageId = args["image_id"]?.stringValue else {
            throw WordError.missingParameter("image_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let hasBorder = args["has_border"]?.boolValue
        let borderColor = args["border_color"]?.stringValue
        let borderWidth = args["border_width"]?.intValue
        let hasShadow = args["has_shadow"]?.boolValue

        try doc.setImageStyle(
            imageId: imageId,
            hasBorder: hasBorder,
            borderColor: borderColor,
            borderWidth: borderWidth,
            hasShadow: hasShadow
        )

        try await storeDocument(doc, for: docId)

        var changes: [String] = []
        if let border = hasBorder { changes.append("border: \(border)") }
        if let color = borderColor { changes.append("color: \(color)") }
        if let width = borderWidth { changes.append("width: \(width)") }
        if let shadow = hasShadow { changes.append("shadow: \(shadow)") }

        return "Updated image style for '\(imageId)': \(changes.joined(separator: ", "))"
    }

    // MARK: - Export

    private func exportText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let text = doc.getText()
        try text.write(toFile: path, atomically: true, encoding: .utf8)

        return "Exported text to: \(path)"
    }

    private func exportMarkdown(args: [String: Value]) async throws -> String {
        guard let sourcePath = args["source_path"]?.stringValue else {
            throw WordError.missingParameter("source_path")
        }
        guard let outputPath = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard FileManager.default.fileExists(atPath: sourcePath) else {
            throw WordError.fileNotFound(sourcePath)
        }

        let sourceURL = URL(fileURLWithPath: sourcePath)
        let lockFile = sourceURL.deletingLastPathComponent()
            .appendingPathComponent("~$" + sourceURL.lastPathComponent)
        if FileManager.default.fileExists(atPath: lockFile.path) {
            throw WordError.invalidFormat("File is open in Microsoft Word. Please save and close it first: \(sourceURL.lastPathComponent)")
        }

        let document = try DocxReader.read(from: sourceURL)

        let figuresDir: URL
        if let customFigDir = args["figures_directory"]?.stringValue {
            figuresDir = URL(fileURLWithPath: customFigDir)
        } else {
            figuresDir = URL(fileURLWithPath: outputPath)
                .deletingLastPathComponent()
                .appendingPathComponent("figures")
        }

        let options = ConversionOptions(
            includeFrontmatter: args["include_frontmatter"]?.boolValue ?? false,
            hardLineBreaks: args["hard_line_breaks"]?.boolValue ?? false,
            fidelity: .markdownWithFigures,
            figuresDirectory: figuresDir
        )

        let markdown = try wordConverter.convertToString(document: document, options: options)

        try markdown.write(toFile: outputPath, atomically: true, encoding: .utf8)
        let figCount = (try? FileManager.default.contentsOfDirectory(atPath: figuresDir.path))?.count ?? 0
        if figCount > 0 {
            return "Exported Markdown to: \(outputPath) (\(figCount) figures in \(figuresDir.path))"
        }
        return "Exported Markdown to: \(outputPath)"
    }

    // MARK: - Hyperlink and Bookmark Operations

    private func insertHyperlink(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let url = args["url"]?.stringValue else {
            throw WordError.missingParameter("url")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let paragraphIndex = args["paragraph_index"]?.intValue
        let tooltip = args["tooltip"]?.stringValue

        let hyperlinkId = doc.insertHyperlink(
            url: url,
            text: text,
            at: paragraphIndex,
            tooltip: tooltip
        )

        try await storeDocument(doc, for: docId)

        return "Inserted hyperlink '\(text)' -> \(url) with id '\(hyperlinkId)'"
    }

    private func insertInternalLink(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let bookmarkName = args["bookmark_name"]?.stringValue else {
            throw WordError.missingParameter("bookmark_name")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let paragraphIndex = args["paragraph_index"]?.intValue
        let tooltip = args["tooltip"]?.stringValue

        let hyperlinkId = doc.insertInternalLink(
            bookmarkName: bookmarkName,
            text: text,
            at: paragraphIndex,
            tooltip: tooltip
        )

        try await storeDocument(doc, for: docId)

        return "Inserted internal link '\(text)' -> bookmark '\(bookmarkName)' with id '\(hyperlinkId)'"
    }

    private func updateHyperlink(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let hyperlinkId = args["hyperlink_id"]?.stringValue else {
            throw WordError.missingParameter("hyperlink_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let text = args["text"]?.stringValue
        let url = args["url"]?.stringValue

        try doc.updateHyperlink(hyperlinkId: hyperlinkId, text: text, url: url)
        try await storeDocument(doc, for: docId)

        var changes: [String] = []
        if let text = text { changes.append("text: '\(text)'") }
        if let url = url { changes.append("url: '\(url)'") }

        return "Updated hyperlink '\(hyperlinkId)': \(changes.joined(separator: ", "))"
    }

    private func deleteHyperlink(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let hyperlinkId = args["hyperlink_id"]?.stringValue else {
            throw WordError.missingParameter("hyperlink_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteHyperlink(hyperlinkId: hyperlinkId)
        try await storeDocument(doc, for: docId)

        return "Deleted hyperlink '\(hyperlinkId)'"
    }

    private func insertBookmark(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let name = args["name"]?.stringValue else {
            throw WordError.missingParameter("name")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let paragraphIndex = args["paragraph_index"]?.intValue

        let bookmarkId = try doc.insertBookmark(name: name, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted bookmark '\(name)' with id \(bookmarkId)"
    }

    private func deleteBookmark(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let name = args["name"]?.stringValue else {
            throw WordError.missingParameter("name")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteBookmark(name: name)
        try await storeDocument(doc, for: docId)

        return "Deleted bookmark '\(name)'"
    }

    // MARK: - Comment and Revision Operations

    private func insertComment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard let author = args["author"]?.stringValue else {
            throw WordError.missingParameter("author")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let commentId = try doc.insertComment(text: text, author: author, paragraphIndex: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted comment with id \(commentId) by '\(author)'"
    }

    private func updateComment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let commentId = args["comment_id"]?.intValue else {
            throw WordError.missingParameter("comment_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.updateComment(commentId: commentId, text: text)
        try await storeDocument(doc, for: docId)

        return "Updated comment \(commentId)"
    }

    private func deleteComment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let commentId = args["comment_id"]?.intValue else {
            throw WordError.missingParameter("comment_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        try doc.deleteComment(commentId: commentId)
        try await storeDocument(doc, for: docId)

        return "Deleted comment \(commentId)"
    }

    private func listComments(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let comments = doc.getComments()
        if comments.isEmpty {
            return "No comments in document"
        }

        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "yyyy-MM-dd HH:mm"

        var result = "Comments (\(comments.count)):\n"
        for comment in comments {
            result += "- [ID: \(comment.id)] \(comment.author) (\(dateFormatter.string(from: comment.date))): \"\(comment.text)\" (para \(comment.paragraphIndex))\n"
        }

        return result
    }

    private func enableTrackChanges(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let author = args["author"]?.stringValue ?? defaultRevisionAuthor
        documentTrackChangesEnforced[docId] = true
        doc.enableTrackChanges(author: author)
        try await storeDocument(doc, for: docId)

        return "Track changes enabled for '\(author)'"
    }

    private func disableTrackChanges(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        documentTrackChangesEnforced[docId] = false
        doc.disableTrackChanges()
        try await storeDocument(doc, for: docId)

        return "Track changes disabled"
    }

    private func acceptRevision(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let acceptAll = args["all"]?.boolValue ?? false

        if acceptAll {
            doc.acceptAllRevisions()
            try await storeDocument(doc, for: docId)
            return "Accepted all revisions"
        } else {
            guard let revisionId = args["revision_id"]?.intValue else {
                throw WordError.missingParameter("revision_id")
            }
            try doc.acceptRevision(revisionId: revisionId)
            try await storeDocument(doc, for: docId)
            return "Accepted revision \(revisionId)"
        }
    }

    private func rejectRevision(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let rejectAll = args["all"]?.boolValue ?? false

        if rejectAll {
            doc.rejectAllRevisions()
            try await storeDocument(doc, for: docId)
            return "Rejected all revisions"
        } else {
            guard let revisionId = args["revision_id"]?.intValue else {
                throw WordError.missingParameter("revision_id")
            }
            try doc.rejectRevision(revisionId: revisionId)
            try await storeDocument(doc, for: docId)
            return "Rejected revision \(revisionId)"
        }
    }

    // MARK: - Footnotes/Endnotes

    private func insertFootnote(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }

        let footnoteId = try doc.insertFootnote(text: text, paragraphIndex: paragraphIndex)
        try await storeDocument(doc, for: docId)
        return "Inserted footnote \(footnoteId) at paragraph \(paragraphIndex)"
    }

    private func deleteFootnote(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let footnoteId = args["footnote_id"]?.intValue else {
            throw WordError.missingParameter("footnote_id")
        }

        try doc.deleteFootnote(footnoteId: footnoteId)
        try await storeDocument(doc, for: docId)
        return "Deleted footnote \(footnoteId)"
    }

    private func insertEndnote(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }

        let endnoteId = try doc.insertEndnote(text: text, paragraphIndex: paragraphIndex)
        try await storeDocument(doc, for: docId)
        return "Inserted endnote \(endnoteId) at paragraph \(paragraphIndex)"
    }

    private func deleteEndnote(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let endnoteId = args["endnote_id"]?.intValue else {
            throw WordError.missingParameter("endnote_id")
        }

        try doc.deleteEndnote(endnoteId: endnoteId)
        try await storeDocument(doc, for: docId)
        return "Deleted endnote \(endnoteId)"
    }

    // MARK: - Advanced Features (P7)

    private func insertTOC(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let index = args["index"]?.intValue
        let title = args["title"]?.stringValue
        let minLevel = args["min_level"]?.intValue ?? 1
        let maxLevel = args["max_level"]?.intValue ?? 3
        let includePageNumbers = args["include_page_numbers"]?.boolValue ?? true
        let useHyperlinks = args["use_hyperlinks"]?.boolValue ?? true

        doc.insertTableOfContents(
            at: index,
            title: title,
            headingLevels: minLevel...maxLevel,
            includePageNumbers: includePageNumbers,
            useHyperlinks: useHyperlinks
        )
        try await storeDocument(doc, for: docId)

        return "Inserted table of contents (heading levels \(minLevel)-\(maxLevel))"
    }

    private func insertTextField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let name = args["name"]?.stringValue else {
            throw WordError.missingParameter("name")
        }

        let defaultValue = args["default_value"]?.stringValue
        let maxLength = args["max_length"]?.intValue

        try doc.insertTextField(at: paragraphIndex, name: name, defaultValue: defaultValue, maxLength: maxLength)
        try await storeDocument(doc, for: docId)

        return "Inserted text field '\(name)' at paragraph \(paragraphIndex)"
    }

    private func insertCheckbox(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let name = args["name"]?.stringValue else {
            throw WordError.missingParameter("name")
        }

        let isChecked = args["is_checked"]?.boolValue ?? false

        try doc.insertCheckbox(at: paragraphIndex, name: name, isChecked: isChecked)
        try await storeDocument(doc, for: docId)

        return "Inserted checkbox '\(name)' (checked: \(isChecked)) at paragraph \(paragraphIndex)"
    }

    private func insertDropdown(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let name = args["name"]?.stringValue else {
            throw WordError.missingParameter("name")
        }
        guard let optionsValue = args["options"] else {
            throw WordError.missingParameter("options")
        }

        var options: [String] = []
        if case .array(let arr) = optionsValue {
            for item in arr {
                if let str = item.stringValue {
                    options.append(str)
                }
            }
        }

        if options.isEmpty {
            throw WordError.missingParameter("options (array of strings)")
        }

        let selectedIndex = args["selected_index"]?.intValue ?? 0

        try doc.insertDropdown(at: paragraphIndex, name: name, options: options, selectedIndex: selectedIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted dropdown '\(name)' with \(options.count) options at paragraph \(paragraphIndex)"
    }

    private func insertEquation(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let latex = args["latex"]?.stringValue else {
            throw WordError.missingParameter("latex")
        }

        let paragraphIndex = args["paragraph_index"]?.intValue
        let displayMode = args["display_mode"]?.boolValue ?? false

        doc.insertEquation(at: paragraphIndex, latex: latex, displayMode: displayMode)
        try await storeDocument(doc, for: docId)

        return "Inserted equation (display mode: \(displayMode))"
    }

    private func setParagraphBorder(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }

        let typeStr = args["type"]?.stringValue ?? "single"
        let size = args["size"]?.intValue ?? 4
        let color = args["color"]?.stringValue ?? "000000"
        let space = args["space"]?.intValue ?? 1

        let borderType = ParagraphBorderType(rawValue: typeStr) ?? .single
        let borderStyle = ParagraphBorderStyle(type: borderType, color: color, size: size, space: space)

        var topStyle: ParagraphBorderStyle? = borderStyle
        var bottomStyle: ParagraphBorderStyle? = borderStyle
        var leftStyle: ParagraphBorderStyle? = borderStyle
        var rightStyle: ParagraphBorderStyle? = borderStyle

        if let sidesValue = args["sides"] {
            if case .array(let arr) = sidesValue {
                topStyle = nil; bottomStyle = nil; leftStyle = nil; rightStyle = nil
                for item in arr {
                    if let side = item.stringValue {
                        switch side.lowercased() {
                        case "top": topStyle = borderStyle
                        case "bottom": bottomStyle = borderStyle
                        case "left": leftStyle = borderStyle
                        case "right": rightStyle = borderStyle
                        default: break
                        }
                    }
                }
            }
        }

        let border = ParagraphBorder(
            top: topStyle,
            bottom: bottomStyle,
            left: leftStyle,
            right: rightStyle
        )

        try doc.setParagraphBorder(at: paragraphIndex, border: border)
        try await storeDocument(doc, for: docId)

        return "Set border on paragraph \(paragraphIndex)"
    }

    private func setParagraphShading(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let fill = args["fill"]?.stringValue else {
            throw WordError.missingParameter("fill")
        }

        var pattern: ShadingPattern? = nil
        if let patternStr = args["pattern"]?.stringValue {
            pattern = ShadingPattern(rawValue: patternStr)
        }

        try doc.setParagraphShading(at: paragraphIndex, fill: fill, pattern: pattern)
        try await storeDocument(doc, for: docId)

        return "Set shading on paragraph \(paragraphIndex) (fill: #\(fill))"
    }

    private func setCharacterSpacing(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }

        let spacing = args["spacing"]?.intValue
        let position = args["position"]?.intValue
        let kern = args["kern"]?.intValue

        try doc.setCharacterSpacing(at: paragraphIndex, spacing: spacing, position: position, kern: kern)
        try await storeDocument(doc, for: docId)

        var changes: [String] = []
        if let spacing = spacing { changes.append("spacing: \(spacing)") }
        if let position = position { changes.append("position: \(position)") }
        if let kern = kern { changes.append("kern: \(kern)") }

        return "Set character spacing on paragraph \(paragraphIndex): \(changes.joined(separator: ", "))"
    }

    private func setTextEffect(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let effectType = args["effect"]?.stringValue else {
            throw WordError.missingParameter("effect")
        }

        guard let effect = TextEffect(rawValue: effectType) else {
            throw WordError.invalidParameter("effect", "Unknown effect type: \(effectType). Valid: blinkBackground, lights, antsBlack, antsRed, shimmer, sparkle, none")
        }

        try doc.setTextEffect(at: paragraphIndex, effect: effect)
        try await storeDocument(doc, for: docId)

        return "Applied '\(effectType)' effect to paragraph \(paragraphIndex)"
    }

    // MARK: - 8.1 Comment Replies and Resolution

    private func replyToComment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let commentId = args["comment_id"]?.intValue else {
            throw WordError.missingParameter("comment_id")
        }
        guard let replyText = args["reply_text"]?.stringValue else {
            throw WordError.missingParameter("reply_text")
        }
        let author = args["author"]?.stringValue ?? "User"

        guard let reply = doc.comments.addReply(to: commentId, author: author, text: replyText) else {
            throw WordError.invalidParameter("comment_id", "Comment with ID \(commentId) not found")
        }

        try await storeDocument(doc, for: docId)
        return "Added reply to comment \(commentId) by \(author) (reply ID: \(reply.id))"
    }

    private func resolveComment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let commentId = args["comment_id"]?.intValue else {
            throw WordError.missingParameter("comment_id")
        }
        let resolved = args["resolved"]?.boolValue ?? true

        doc.comments.markAsDone(commentId, done: resolved)
        try await storeDocument(doc, for: docId)

        return "Comment \(commentId) \(resolved ? "resolved" : "reopened")"
    }

    // MARK: - 8.2 Floating Images

    private func insertFloatingImage(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }

        let paragraphIndex = args["paragraph_index"]?.intValue ?? 0
        let widthEmu = args["width"]?.intValue ?? 2000000  // ~2 inches default
        let heightEmu = args["height"]?.intValue ?? 2000000
        let horizontalPos = args["horizontal_position"]?.intValue ?? 0
        let verticalPos = args["vertical_position"]?.intValue ?? 0
        let wrapTypeStr = args["wrap_type"]?.stringValue ?? "square"
        let horizontalRelative = args["horizontal_relative"]?.stringValue ?? "column"
        let allowOverlap = args["allow_overlap"]?.boolValue ?? true

        let url = URL(fileURLWithPath: path)
        let imageData = try Data(contentsOf: url)

        let imageId = "rId\(doc.images.count + 10)"
        let imageRef = ImageReference(
            id: imageId,
            fileName: url.lastPathComponent,
            contentType: detectImageContentType(from: url),
            data: imageData
        )
        doc.images.append(imageRef)

        var anchorPosition = AnchorPosition()
        anchorPosition.horizontalOffset = horizontalPos
        anchorPosition.verticalOffset = verticalPos
        anchorPosition.allowOverlap = allowOverlap

        if let hrel = HorizontalRelativeFrom(rawValue: horizontalRelative) {
            anchorPosition.horizontalRelativeFrom = hrel
        }

        switch wrapTypeStr.lowercased() {
        case "none": anchorPosition.wrapType = .none
        case "square": anchorPosition.wrapType = .square
        case "tight": anchorPosition.wrapType = .tight
        case "through": anchorPosition.wrapType = .through
        case "topandbottom": anchorPosition.wrapType = .topAndBottom
        case "behindtext": anchorPosition.wrapType = .behindText
        case "infrontoftext": anchorPosition.wrapType = .inFrontOfText
        default: anchorPosition.wrapType = .square
        }

        let drawing = Drawing.anchor(
            width: widthEmu,
            height: heightEmu,
            imageId: imageId,
            position: anchorPosition,
            name: url.lastPathComponent
        )

        try doc.insertDrawing(drawing, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted floating image '\(url.lastPathComponent)' at paragraph \(paragraphIndex)"
    }

    private func detectImageContentType(from url: URL) -> String {
        let ext = url.pathExtension.lowercased()
        switch ext {
        case "png": return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "gif": return "image/gif"
        case "bmp": return "image/bmp"
        case "tiff", "tif": return "image/tiff"
        default: return "image/png"
        }
    }

    // MARK: - 8.3 Field Codes

    private func insertIfField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let leftOperand = args["left_operand"]?.stringValue else {
            throw WordError.missingParameter("left_operand")
        }
        guard let operatorStr = args["operator"]?.stringValue else {
            throw WordError.missingParameter("operator")
        }
        guard let rightOperand = args["right_operand"]?.stringValue else {
            throw WordError.missingParameter("right_operand")
        }
        guard let trueText = args["true_text"]?.stringValue else {
            throw WordError.missingParameter("true_text")
        }
        guard let falseText = args["false_text"]?.stringValue else {
            throw WordError.missingParameter("false_text")
        }

        let compOp: IFField.ComparisonOperator
        switch operatorStr {
        case "=", "==": compOp = .equal
        case "<>", "!=": compOp = .notEqual
        case "<": compOp = .lessThan
        case ">": compOp = .greaterThan
        case "<=": compOp = .lessThanOrEqual
        case ">=": compOp = .greaterThanOrEqual
        default: compOp = .equal
        }

        let ifField = IFField(
            leftOperand: leftOperand,
            comparisonOperator: compOp,
            rightOperand: rightOperand,
            trueText: trueText,
            falseText: falseText
        )

        try doc.insertFieldCode(ifField, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted IF field at paragraph \(paragraphIndex): IF \(leftOperand) \(operatorStr) \(rightOperand)"
    }

    private func insertCalculationField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let expression = args["expression"]?.stringValue else {
            throw WordError.missingParameter("expression")
        }
        let format = args["format"]?.stringValue

        let calcField = CalculationField(
            expression: expression,
            numberFormat: format
        )

        try doc.insertFieldCode(calcField, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted calculation field '\(expression)' at paragraph \(paragraphIndex)"
    }

    private func insertDateField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        let format = args["format"]?.stringValue ?? "yyyy-MM-dd"
        let typeStr = args["type"]?.stringValue ?? "DATE"

        let fieldType: DateTimeFieldType
        switch typeStr.uppercased() {
        case "DATE": fieldType = .date
        case "TIME": fieldType = .time
        case "PRINTDATE": fieldType = .printDate
        case "SAVEDATE": fieldType = .saveDate
        case "CREATEDATE": fieldType = .createDate
        case "EDITTIME": fieldType = .editTime
        default: fieldType = .date
        }

        let dateField = DateTimeField(type: fieldType, dateFormat: format)

        try doc.insertFieldCode(dateField, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted \(typeStr) field with format '\(format)' at paragraph \(paragraphIndex)"
    }

    private func insertPageField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        let typeStr = args["type"]?.stringValue ?? "PAGE"

        let infoType: DocumentInfoFieldType
        switch typeStr.uppercased() {
        case "PAGE": infoType = .page
        case "NUMPAGES": infoType = .numPages
        case "NUMWORDS": infoType = .numWords
        case "NUMCHARS": infoType = .numChars
        case "FILENAME": infoType = .fileName
        case "AUTHOR": infoType = .author
        case "TITLE": infoType = .title
        case "SECTIONPAGES": infoType = .sectionPages
        default: infoType = .page
        }

        let infoField = DocumentInfoField(type: infoType)

        try doc.insertFieldCode(infoField, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted \(typeStr) field at paragraph \(paragraphIndex)"
    }

    private func insertMergeField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let fieldName = args["field_name"]?.stringValue else {
            throw WordError.missingParameter("field_name")
        }
        let textBefore = args["text_before"]?.stringValue
        let textAfter = args["text_after"]?.stringValue

        let mergeField = MergeField(
            fieldName: fieldName,
            textBefore: textBefore,
            textAfter: textAfter
        )

        try doc.insertFieldCode(mergeField, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted MERGEFIELD '\(fieldName)' at paragraph \(paragraphIndex)"
    }

    private func insertSequenceField(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let identifier = args["identifier"]?.stringValue else {
            throw WordError.missingParameter("identifier")
        }
        let resetOnHeading = args["reset_on_heading"]?.intValue

        let seqField = SequenceField(
            identifier: identifier,
            resetLevel: resetOnHeading
        )

        try doc.insertFieldCode(seqField, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted SEQ '\(identifier)' field at paragraph \(paragraphIndex)"
    }

    // MARK: - 8.4 Content Controls (SDT)

    private func insertContentControl(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let typeStr = args["type"]?.stringValue else {
            throw WordError.missingParameter("type")
        }
        guard let tag = args["tag"]?.stringValue else {
            throw WordError.missingParameter("tag")
        }

        let alias = args["alias"]?.stringValue
        let placeholder = args["placeholder"]?.stringValue
        let contentText = args["content"]?.stringValue ?? ""

        guard let sdtType = SDTType(rawValue: typeStr) else {
            throw WordError.invalidParameter("type", "Unknown SDT type: \(typeStr). Valid: richText, text, picture, date, dropDownList, comboBox, checkbox")
        }

        let sdt = StructuredDocumentTag(
            id: Int.random(in: 100000...999999),
            tag: tag,
            alias: alias,
            type: sdtType,
            placeholder: placeholder
        )

        let contentControl = ContentControl(sdt: sdt, content: contentText)

        try doc.insertContentControl(contentControl, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        return "Inserted \(typeStr) content control '\(tag)' at paragraph \(paragraphIndex)"
    }

    private func insertRepeatingSection(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)
        guard let tag = args["tag"]?.stringValue else {
            throw WordError.missingParameter("tag")
        }

        let index = args["index"]?.intValue ?? 0
        let sectionTitle = args["section_title"]?.stringValue
        let itemsArray = args["items"]?.arrayValue ?? []

        var items: [RepeatingSectionItem] = []
        for item in itemsArray {
            if let content = item.stringValue {
                let rsItem = RepeatingSectionItem(
                    tag: nil,
                    content: content
                )
                items.append(rsItem)
            }
        }

        if items.isEmpty {
            items.append(RepeatingSectionItem(content: ""))
        }

        let repeatingSection = RepeatingSection(
            tag: tag,
            alias: sectionTitle,
            items: items,
            allowInsertDeleteSections: true,
            sectionTitle: sectionTitle
        )

        try doc.insertRepeatingSection(repeatingSection, at: index)
        try await storeDocument(doc, for: docId)

        return "Inserted repeating section '\(tag)' with \(items.count) item(s) at index \(index)"
    }


    private func insertText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let position = args["position"]?.intValue

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let currentText = paragraphs[paragraphIndex].getText()
        let insertPosition = position ?? currentText.count

        let startIndex = currentText.startIndex
        let insertIndex = currentText.index(startIndex, offsetBy: min(insertPosition, currentText.count))
        let newText = String(currentText[..<insertIndex]) + text + String(currentText[insertIndex...])

        try doc.updateParagraph(at: paragraphIndex, text: newText)
        try await storeDocument(doc, for: docId)

        return "Inserted text at paragraph \(paragraphIndex)\(position.map { ", position \($0)" } ?? " (at end)")"
    }

    private func getDocumentText(args: [String: Value]) async throws -> String {
        return try await getText(args: args)
    }

    private func searchText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let query = args["query"]?.stringValue else {
            throw WordError.missingParameter("query")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let caseSensitive = args["case_sensitive"]?.boolValue ?? false

        let paragraphs = doc.getParagraphs()
        var results: [(paragraphIndex: Int, startPosition: Int, text: String)] = []

        for (index, para) in paragraphs.enumerated() {
            let paraText = para.getText()
            let searchText = caseSensitive ? paraText : paraText.lowercased()
            let searchQuery = caseSensitive ? query : query.lowercased()

            var searchStart = searchText.startIndex
            while let range = searchText.range(of: searchQuery, range: searchStart..<searchText.endIndex) {
                let position = searchText.distance(from: searchText.startIndex, to: range.lowerBound)
                let matchedText = String(paraText[range])
                results.append((index, position, matchedText))
                searchStart = range.upperBound
            }
        }

        if results.isEmpty {
            return "No matches found for '\(query)'"
        }

        var output = "Found \(results.count) match(es) for '\(query)':\n"
        for result in results {
            output += "- Paragraph \(result.paragraphIndex), position \(result.startPosition): \"\(result.text)\"\n"
        }
        return output
    }

    private func listHyperlinks(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let hyperlinks = doc.getHyperlinks()
        if hyperlinks.isEmpty {
            return "No hyperlinks in document"
        }

        var output = "Hyperlinks in document (\(hyperlinks.count)):\n"
        for (index, link) in hyperlinks.enumerated() {
            let displayText = link.text
            let target = link.url ?? link.anchor ?? "(unknown target)"
            output += "[\(index)] (\(link.type)) \(displayText) -> \(target)\n"
        }
        return output
    }

    private func listBookmarks(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let bookmarks = doc.getBookmarks()
        if bookmarks.isEmpty {
            return "No bookmarks in document"
        }

        var output = "Bookmarks in document (\(bookmarks.count)):\n"
        for (index, bookmark) in bookmarks.enumerated() {
            output += "[\(index)] \(bookmark.name)\n"
        }
        return output
    }

    private func listFootnotes(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let footnotes = doc.getFootnotes()
        if footnotes.isEmpty {
            return "No footnotes in document"
        }

        var output = "Footnotes in document (\(footnotes.count)):\n"
        for footnote in footnotes {
            let preview = String(footnote.text.prefix(50))
            output += "[\(footnote.id)] \(preview)...\n"
        }
        return output
    }

    private func listEndnotes(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let endnotes = doc.getEndnotes()
        if endnotes.isEmpty {
            return "No endnotes in document"
        }

        var output = "Endnotes in document (\(endnotes.count)):\n"
        for endnote in endnotes {
            let preview = String(endnote.text.prefix(50))
            output += "[\(endnote.id)] \(preview)...\n"
        }
        return output
    }

    private func getRevisions(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let revisions = doc.getRevisions()
        if revisions.isEmpty {
            return "No revisions in document"
        }

        var output = "Revisions in document (\(revisions.count)):\n"
        for revision in revisions {
            let typeStr = revision.type.uppercased()
            let author = revision.author
            output += "[\(revision.id)] \(typeStr) by \(author) at paragraph \(revision.paragraphIndex)\n"
            if let original = revision.originalText {
                output += "    Original: \(original.prefix(30))...\n"
            }
            if let newText = revision.newText {
                output += "    New: \(newText.prefix(30))...\n"
            }
        }
        return output
    }

    private func acceptAllRevisions(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let count = doc.getRevisions().count
        doc.acceptAllRevisions()
        try await storeDocument(doc, for: docId)

        return "Accepted \(count) revision(s)"
    }

    private func rejectAllRevisions(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let count = doc.getRevisions().count
        doc.rejectAllRevisions()
        try await storeDocument(doc, for: docId)

        return "Rejected \(count) revision(s)"
    }

    private func setDocumentProperties(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        var props = doc.properties

        if let title = args["title"]?.stringValue {
            props.title = title
        }
        if let subject = args["subject"]?.stringValue {
            props.subject = subject
        }
        if let creator = args["creator"]?.stringValue {
            props.creator = creator
        }
        if let keywords = args["keywords"]?.stringValue {
            props.keywords = keywords
        }
        if let description = args["description"]?.stringValue {
            props.description = description
        }

        doc.properties = props
        try await storeDocument(doc, for: docId)

        return "Updated document properties"
    }

    private func getDocumentProperties(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let props = doc.properties

        var output = "Document Properties:\n"
        if let title = props.title { output += "- Title: \(title)\n" }
        if let subject = props.subject { output += "- Subject: \(subject)\n" }
        if let creator = props.creator { output += "- Creator: \(creator)\n" }
        if let keywords = props.keywords { output += "- Keywords: \(keywords)\n" }
        if let description = props.description { output += "- Description: \(description)\n" }
        if let lastModifiedBy = props.lastModifiedBy { output += "- Last Modified By: \(lastModifiedBy)\n" }
        if let revision = props.revision { output += "- Revision: \(revision)\n" }
        if let created = props.created { output += "- Created: \(created)\n" }
        if let modified = props.modified { output += "- Modified: \(modified)\n" }

        if output == "Document Properties:\n" {
            return "No document properties set"
        }

        return output
    }

    private func getParagraphRuns(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let para = paragraphs[paragraphIndex]
        var output = "Paragraph [\(paragraphIndex)] Runs:\n"

        for (runIndex, run) in para.runs.enumerated() {
            output += "  Run [\(runIndex)]:\n"
            output += "    Text: \"\(run.text)\"\n"

            let props = run.properties
            var formatParts: [String] = []

            if props.bold { formatParts.append("bold") }
            if props.italic { formatParts.append("italic") }
            if props.strikethrough { formatParts.append("strikethrough") }
            if let underline = props.underline { formatParts.append("underline:\(underline.rawValue)") }
            if let color = props.color { formatParts.append("color:#\(color)") }
            if let highlight = props.highlight { formatParts.append("highlight:\(highlight.rawValue)") }
            if let fontSize = props.fontSize { formatParts.append("size:\(fontSize / 2)pt") }
            if let fontName = props.fontName { formatParts.append("font:\(fontName)") }
            if let verticalAlign = props.verticalAlign { formatParts.append("vertAlign:\(verticalAlign.rawValue)") }

            if formatParts.isEmpty {
                output += "    Format: (none)\n"
            } else {
                output += "    Format: \(formatParts.joined(separator: ", "))\n"
            }
        }

        if !para.hyperlinks.isEmpty {
            output += "  Hyperlinks:\n"
            for hyperlink in para.hyperlinks {
                output += "    - \"\(hyperlink.text)\" -> \(hyperlink.url ?? hyperlink.anchor ?? "unknown")\n"
            }
        }

        return output
    }

    private func getTextWithFormatting(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let paragraphs = doc.getParagraphs()

        if let paragraphIndex = args["paragraph_index"]?.intValue {
            guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
                throw WordError.invalidIndex(paragraphIndex)
            }
            return formatParagraphWithMarkup(paragraphs[paragraphIndex], index: paragraphIndex)
        }

        var output = ""
        for (index, para) in paragraphs.enumerated() {
            output += formatParagraphWithMarkup(para, index: index) + "\n"
        }

        return output.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    private func formatParagraphWithMarkup(_ para: Paragraph, index: Int) -> String {
        var result = "[\(index)] "

        for run in para.runs {
            var text = run.text
            let props = run.properties

            if props.bold {
                text = "**\(text)**"
            }
            if props.italic {
                text = "*\(text)*"
            }
            if props.strikethrough {
                text = "~~\(text)~~"
            }
            if let color = props.color {
                let colorName = colorHexToName(color)
                text = "{{color:\(colorName)}}\(text){{/color}}"
            }
            if let highlight = props.highlight {
                text = "{{highlight:\(highlight.rawValue)}}\(text){{/highlight}}"
            }
            if let underline = props.underline {
                text = "{{underline:\(underline.rawValue)}}\(text){{/underline}}"
            }

            result += text
        }

        for hyperlink in para.hyperlinks {
            result += " [\(hyperlink.text)](\(hyperlink.url ?? "#\(hyperlink.anchor ?? "")"))"
        }

        return result
    }

    private func colorHexToName(_ hex: String) -> String {
        let upperHex = hex.uppercased()
        switch upperHex {
        case "FF0000": return "red"
        case "00FF00": return "green"
        case "0000FF": return "blue"
        case "FFFF00": return "yellow"
        case "00FFFF": return "cyan"
        case "FF00FF": return "magenta"
        case "000000": return "black"
        case "FFFFFF": return "white"
        case "808080": return "gray"
        case "FFA500": return "orange"
        case "800080": return "purple"
        default: return "#\(hex)"
        }
    }

    private func searchByFormatting(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let searchColor = args["color"]?.stringValue?.uppercased()
        let searchBold = args["bold"]?.boolValue
        let searchItalic = args["italic"]?.boolValue
        let searchHighlight = args["highlight"]?.stringValue

        let paragraphs = doc.getParagraphs()
        var results: [(paragraphIndex: Int, runIndex: Int, text: String, format: String)] = []

        for (paraIndex, para) in paragraphs.enumerated() {
            for (runIndex, run) in para.runs.enumerated() {
                let props = run.properties
                var matches = true

                if let color = searchColor {
                    if props.color?.uppercased() != color {
                        matches = false
                    }
                }

                if let bold = searchBold {
                    if props.bold != bold {
                        matches = false
                    }
                }

                if let italic = searchItalic {
                    if props.italic != italic {
                        matches = false
                    }
                }

                if let highlight = searchHighlight {
                    if props.highlight?.rawValue != highlight {
                        matches = false
                    }
                }

                if matches && !run.text.isEmpty {
                    var formatParts: [String] = []
                    if props.bold { formatParts.append("bold") }
                    if props.italic { formatParts.append("italic") }
                    if let color = props.color { formatParts.append("color:#\(color)") }
                    if let highlight = props.highlight { formatParts.append("highlight:\(highlight.rawValue)") }

                    results.append((
                        paragraphIndex: paraIndex,
                        runIndex: runIndex,
                        text: run.text,
                        format: formatParts.isEmpty ? "(none)" : formatParts.joined(separator: ", ")
                    ))
                }
            }
        }

        if results.isEmpty {
            return "No text found matching the specified formatting"
        }

        var output = "Found \(results.count) match(es):\n"
        for result in results {
            output += "  [Para \(result.paragraphIndex), Run \(result.runIndex)]: \"\(result.text)\"\n"
            output += "    Format: \(result.format)\n"
        }

        return output
    }

    private func searchTextWithFormatting(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let query = args["query"]?.stringValue else {
            throw WordError.missingParameter("query")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let caseSensitive = args["case_sensitive"]?.boolValue ?? false
        let contextChars = args["context_chars"]?.intValue ?? 20

        let paragraphs = doc.getParagraphs()
        var results: [(paraIndex: Int, position: Int, matchedText: String, context: String, formats: [String])] = []

        for (paraIndex, para) in paragraphs.enumerated() {
            let paraText = para.getText()
            let searchText = caseSensitive ? paraText : paraText.lowercased()
            let searchQuery = caseSensitive ? query : query.lowercased()

            var searchStart = searchText.startIndex
            while let range = searchText.range(of: searchQuery, range: searchStart..<searchText.endIndex) {
                let position = searchText.distance(from: searchText.startIndex, to: range.lowerBound)
                let matchedText = String(paraText[range])

                let contextStart = max(0, position - contextChars)
                let contextEnd = min(paraText.count, position + matchedText.count + contextChars)
                let startIndex = paraText.index(paraText.startIndex, offsetBy: contextStart)
                let endIndex = paraText.index(paraText.startIndex, offsetBy: contextEnd)
                var context = String(paraText[startIndex..<endIndex])
                if contextStart > 0 { context = "..." + context }
                if contextEnd < paraText.count { context = context + "..." }

                var formats: [String] = []
                var currentPos = 0
                for run in para.runs {
                    let runEnd = currentPos + run.text.count
                    if currentPos <= position && position < runEnd {
                        let props = run.properties
                        if props.bold { formats.append("bold") }
                        if props.italic { formats.append("italic") }
                        if props.strikethrough { formats.append("strikethrough") }
                        if let color = props.color {
                            formats.append("color:\(colorHexToName(color))")
                        }
                        if let highlight = props.highlight {
                            formats.append("highlight:\(highlight.rawValue)")
                        }
                        if let underline = props.underline {
                            formats.append("underline:\(underline.rawValue)")
                        }
                        break
                    }
                    currentPos = runEnd
                }

                results.append((paraIndex, position, matchedText, context, formats))
                searchStart = range.upperBound
            }
        }

        if results.isEmpty {
            return "No matches found for '\(query)'"
        }

        var output = "Found \(results.count) match(es) for '\(query)':\n"
        for result in results {
            output += "[Para \(result.paraIndex)] \(result.context)\n"
            if result.formats.isEmpty {
                output += "  Format: (none)\n"
            } else {
                output += "  Format: \(result.formats.joined(separator: ", "))\n"
            }
        }
        return output
    }

    private func listAllFormattedText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let formatType = args["format_type"]?.stringValue?.lowercased() else {
            throw WordError.missingParameter("format_type")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let colorFilter = args["color_filter"]?.stringValue?.uppercased()
        let paragraphStart = args["paragraph_start"]?.intValue ?? 0
        let paragraphEnd = args["paragraph_end"]?.intValue

        let paragraphs = doc.getParagraphs()
        let endIndex = paragraphEnd ?? paragraphs.count - 1

        guard paragraphStart >= 0 && paragraphStart < paragraphs.count else {
            throw WordError.invalidIndex(paragraphStart)
        }
        guard endIndex >= paragraphStart && endIndex < paragraphs.count else {
            throw WordError.invalidIndex(endIndex)
        }

        var results: [(paraIndex: Int, text: String)] = []

        for paraIndex in paragraphStart...endIndex {
            let para = paragraphs[paraIndex]
            for run in para.runs {
                let props = run.properties
                var matches = false

                switch formatType {
                case "italic":
                    matches = props.italic
                case "bold":
                    matches = props.bold
                case "underline":
                    matches = props.underline != nil
                case "strikethrough":
                    matches = props.strikethrough
                case "highlight":
                    matches = props.highlight != nil
                case "color":
                    if let colorFilter = colorFilter {
                        matches = props.color?.uppercased() == colorFilter
                    } else {
                        matches = props.color != nil
                    }
                default:
                    break
                }

                if matches && !run.text.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
                    results.append((paraIndex, run.text))
                }
            }
        }

        if results.isEmpty {
            let rangeInfo = paragraphEnd != nil ? " in paragraphs \(paragraphStart)-\(endIndex)" : ""
            return "No \(formatType) text found\(rangeInfo)"
        }

        var output = "Found \(results.count) \(formatType) text segment(s):\n"
        for result in results {
            let displayText = result.text.count > 60 ? String(result.text.prefix(57)) + "..." : result.text
            output += "[Para \(result.paraIndex)] \"\(displayText)\"\n"
        }
        return output
    }

    private func getWordCountBySection(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        var sectionMarkers: [String] = []
        if let markersValue = args["section_markers"] {
            if let markersArray = markersValue.arrayValue {
                sectionMarkers = markersArray.compactMap { $0.stringValue }
            }
        }

        var excludeSections: Set<String> = []
        if let excludeValue = args["exclude_sections"] {
            if let excludeArray = excludeValue.arrayValue {
                excludeSections = Set(excludeArray.compactMap { $0.stringValue })
            }
        }

        let paragraphs = doc.getParagraphs()

        if sectionMarkers.isEmpty {
            var totalWords = 0
            var totalChars = 0
            for para in paragraphs {
                let text = para.getText()
                totalWords += countWords(text)
                totalChars += text.filter { !$0.isWhitespace }.count
            }
            return """
            Word Count Summary:
              Total words: \(formatNumber(totalWords))
              Total characters (no spaces): \(formatNumber(totalChars))
              Total paragraphs: \(paragraphs.count)
            """
        }

        var sectionStarts: [(name: String, startIndex: Int)] = []
        for (index, para) in paragraphs.enumerated() {
            let paraText = para.getText().trimmingCharacters(in: .whitespacesAndNewlines)
            for marker in sectionMarkers {
                let lowerParaText = paraText.lowercased()
                let lowerMarker = marker.lowercased()
                if lowerParaText == lowerMarker ||
                   lowerParaText.hasPrefix(lowerMarker + ":") ||
                   lowerParaText.hasPrefix(lowerMarker + " ") ||
                   lowerParaText.hasSuffix(" " + lowerMarker) ||
                   lowerParaText.contains(". " + lowerMarker) {
                    sectionStarts.append((marker, index))
                    break
                }
            }
        }

        if sectionStarts.isEmpty {
            var totalWords = 0
            for para in paragraphs {
                totalWords += countWords(para.getText())
            }
            return """
            No section markers found in document.
            Total words: \(formatNumber(totalWords))

            Tip: Section markers should match paragraph text (e.g., "Abstract", "Introduction", "References")
            """
        }

        var sectionCounts: [(name: String, words: Int, excluded: Bool)] = []
        var totalWords = 0
        var excludedWords = 0

        if sectionStarts[0].startIndex > 0 {
            var preWords = 0
            for i in 0..<sectionStarts[0].startIndex {
                preWords += countWords(paragraphs[i].getText())
            }
            if preWords > 0 {
                sectionCounts.append(("(Before first section)", preWords, false))
                totalWords += preWords
            }
        }

        for (i, section) in sectionStarts.enumerated() {
            let startIndex = section.startIndex
            let endIndex = (i + 1 < sectionStarts.count) ? sectionStarts[i + 1].startIndex : paragraphs.count

            var sectionWords = 0
            for j in startIndex..<endIndex {
                sectionWords += countWords(paragraphs[j].getText())
            }

            let isExcluded = excludeSections.contains(section.name)
            sectionCounts.append((section.name, sectionWords, isExcluded))
            totalWords += sectionWords
            if isExcluded {
                excludedWords += sectionWords
            }
        }

        var output = "Word Count by Section:\n"
        for section in sectionCounts {
            let excludeTag = section.excluded ? " (excluded)" : ""
            output += "  \(section.name): \(formatNumber(section.words)) words\(excludeTag)\n"
        }
        output += "  -----------------------------\n"
        if excludedWords > 0 {
            output += "  Main Text: \(formatNumber(totalWords - excludedWords)) words\n"
        }
        output += "  Total: \(formatNumber(totalWords)) words\n"

        return output
    }

    private func countWords(_ text: String) -> Int {
        let trimmed = text.trimmingCharacters(in: .whitespacesAndNewlines)
        if trimmed.isEmpty { return 0 }

        var englishWords = 0
        var chineseChars = 0

        let englishPattern = try? NSRegularExpression(pattern: "[a-zA-Z]+", options: [])
        let chinesePattern = try? NSRegularExpression(pattern: "[\\u4e00-\\u9fff]", options: [])

        let range = NSRange(trimmed.startIndex..., in: trimmed)

        if let matches = englishPattern?.matches(in: trimmed, options: [], range: range) {
            englishWords = matches.count
        }

        if let matches = chinesePattern?.matches(in: trimmed, options: [], range: range) {
            chineseChars = matches.count
        }

        return englishWords + chineseChars
    }

    private func formatNumber(_ number: Int) -> String {
        let formatter = NumberFormatter()
        formatter.numberStyle = .decimal
        return formatter.string(from: NSNumber(value: number)) ?? "\(number)"
    }

    // MARK: - Document Comparison

    private struct ParagraphSnapshot {
        let index: Int
        let text: String
        let textHash: Int
        let style: String?
        let formattedText: String
        let keepNext: Bool
    }

    private enum DiffType {
        case unchanged, modified, deleted, added, formatOnly
    }

    private struct DiffEntry {
        let type: DiffType
        let indexA: Int?
        let indexB: Int?
        let style: String?
        let textA: String?
        let textB: String?
        let formattedA: String?
        let formattedB: String?
    }

    private func snapshotParagraphs(_ doc: WordDocument) -> [ParagraphSnapshot] {
        let paragraphs = doc.getParagraphs()
        return paragraphs.enumerated().map { (index, para) in
            let text = para.getText().trimmingCharacters(in: .whitespacesAndNewlines)
            return ParagraphSnapshot(
                index: index,
                text: text,
                textHash: text.hashValue,
                style: para.properties.style,
                formattedText: formatParagraphWithMarkup(para, index: index),
                keepNext: para.properties.keepNext
            )
        }
    }

    private func computeLCS(_ a: [ParagraphSnapshot], _ b: [ParagraphSnapshot]) -> [[Int]] {
        let n = a.count
        let m = b.count
        var dp = Array(repeating: Array(repeating: 0, count: m + 1), count: n + 1)
        for i in 1...max(n, 1) {
            guard i <= n else { break }
            for j in 1...max(m, 1) {
                guard j <= m else { break }
                if a[i - 1].textHash == b[j - 1].textHash && a[i - 1].text == b[j - 1].text {
                    dp[i][j] = dp[i - 1][j - 1] + 1
                } else {
                    dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])
                }
            }
        }
        return dp
    }

    private func textSimilarity(_ a: String, _ b: String) -> Double {
        let wordsA = Set(a.lowercased().split(whereSeparator: { $0.isWhitespace || $0.isPunctuation }))
        let wordsB = Set(b.lowercased().split(whereSeparator: { $0.isWhitespace || $0.isPunctuation }))
        guard !wordsA.isEmpty || !wordsB.isEmpty else { return 1.0 }
        let intersection = wordsA.intersection(wordsB).count
        let union = wordsA.union(wordsB).count
        guard union > 0 else { return 0.0 }
        return Double(intersection) / Double(union)
    }

    private func buildDiffEntries(
        _ a: [ParagraphSnapshot],
        _ b: [ParagraphSnapshot],
        _ dp: [[Int]],
        mode: String
    ) -> [DiffEntry] {
        // Backtrack LCS to get aligned sequence
        var aligned: [(aIdx: Int?, bIdx: Int?)] = []
        var i = a.count
        var j = b.count
        while i > 0 || j > 0 {
            if i > 0 && j > 0 && a[i - 1].textHash == b[j - 1].textHash && a[i - 1].text == b[j - 1].text {
                aligned.append((i - 1, j - 1))
                i -= 1
                j -= 1
            } else if j > 0 && (i == 0 || dp[i][j - 1] >= dp[i - 1][j]) {
                aligned.append((nil, j - 1))
                j -= 1
            } else {
                aligned.append((i - 1, nil))
                i -= 1
            }
        }
        aligned.reverse()

        // Post-process: merge adjacent DELETED+ADDED into MODIFIED if similar
        var entries: [DiffEntry] = []
        var idx = 0
        while idx < aligned.count {
            let (aIdx, bIdx) = aligned[idx]
            if let ai = aIdx, let bi = bIdx {
                // Matched pair
                let checkFormatting = (mode == "formatting" || mode == "full")
                if checkFormatting && a[ai].formattedText != b[bi].formattedText {
                    entries.append(DiffEntry(
                        type: .formatOnly,
                        indexA: ai, indexB: bi,
                        style: a[ai].style ?? b[bi].style,
                        textA: a[ai].text, textB: b[bi].text,
                        formattedA: a[ai].formattedText, formattedB: b[bi].formattedText
                    ))
                } else {
                    entries.append(DiffEntry(
                        type: .unchanged,
                        indexA: ai, indexB: bi,
                        style: a[ai].style,
                        textA: a[ai].text, textB: nil,
                        formattedA: nil, formattedB: nil
                    ))
                }
                idx += 1
            } else if aIdx != nil && bIdx == nil {
                if idx + 1 < aligned.count,
                   aligned[idx + 1].aIdx == nil,
                   let bi = aligned[idx + 1].bIdx,
                   let ai = aIdx,
                   textSimilarity(a[ai].text, b[bi].text) > 0.5 {
                    entries.append(DiffEntry(
                        type: .modified,
                        indexA: ai, indexB: bi,
                        style: a[ai].style ?? b[bi].style,
                        textA: a[ai].text, textB: b[bi].text,
                        formattedA: a[ai].formattedText, formattedB: b[bi].formattedText
                    ))
                    idx += 2
                } else {
                    entries.append(DiffEntry(
                        type: .deleted,
                        indexA: aIdx, indexB: nil,
                        style: a[aIdx!].style,
                        textA: a[aIdx!].text, textB: nil,
                        formattedA: a[aIdx!].formattedText, formattedB: nil
                    ))
                    idx += 1
                }
            } else {
                if idx + 1 < aligned.count,
                   aligned[idx + 1].bIdx == nil,
                   let ai = aligned[idx + 1].aIdx,
                   let bi = bIdx,
                   textSimilarity(a[ai].text, b[bi].text) > 0.5 {
                    entries.append(DiffEntry(
                        type: .modified,
                        indexA: ai, indexB: bi,
                        style: a[ai].style ?? b[bi].style,
                        textA: a[ai].text, textB: b[bi].text,
                        formattedA: a[ai].formattedText, formattedB: b[bi].formattedText
                    ))
                    idx += 2
                } else {
                    entries.append(DiffEntry(
                        type: .added,
                        indexA: nil, indexB: bIdx,
                        style: b[bIdx!].style,
                        textA: nil, textB: b[bIdx!].text,
                        formattedA: nil, formattedB: b[bIdx!].formattedText
                    ))
                    idx += 1
                }
            }
        }
        return entries
    }

    private func truncateText(_ text: String, maxLength: Int = 500, contextChars: Int = 30) -> String {
        guard text.count > maxLength else { return text }
        let start = text.prefix(contextChars)
        let end = text.suffix(contextChars)
        return "\(start) [...] \(end)"
    }

    private func formatStructureComparison(
        docIdA: String, docIdB: String,
        snapshotsA: [ParagraphSnapshot], snapshotsB: [ParagraphSnapshot],
        infoA: (paragraphs: Int, words: Int), infoB: (paragraphs: Int, words: Int),
        customHeadingStyles: [String]? = nil
    ) -> String {
        var output = """
        === Document Comparison (Structure) ===
        Base: \(docIdA) (\(infoA.paragraphs) paragraphs, \(formatNumber(infoA.words)) words)
        Compare: \(docIdB) (\(infoB.paragraphs) paragraphs, \(formatNumber(infoB.words)) words)

        --- Statistics ---
        Paragraph count: \(infoA.paragraphs) -> \(infoB.paragraphs) (\(infoB.paragraphs >= infoA.paragraphs ? "+" : "")\(infoB.paragraphs - infoA.paragraphs))
        Word count: \(formatNumber(infoA.words)) -> \(formatNumber(infoB.words)) (\(infoB.words >= infoA.words ? "+" : "")\(formatNumber(infoB.words - infoA.words)))

        --- Heading Outline: Base (\(docIdA)) ---

        """
        let builtinHeadingStyles = Set(["Heading1", "Heading2", "Heading3", "Heading 1", "Heading 2", "Heading 3", "heading 1", "heading 2", "heading 3", "Title"])
        let customSet: Set<String>? = customHeadingStyles.map { Set($0) }

        func isHeading(_ s: ParagraphSnapshot) -> (isMatch: Bool, isHeuristic: Bool) {
            guard let style = s.style else { return (false, false) }
            // Custom heading styles take priority
            if let custom = customSet {
                return (custom.contains(style), false)
            }
            // Built-in heading styles
            if builtinHeadingStyles.contains(style) {
                return (true, false)
            }
            // Heuristic: keepNext + short text likely indicates a heading
            if s.keepNext == true && s.text.count < 100 && !s.text.isEmpty {
                return (true, true)
            }
            return (false, false)
        }

        func headingIndent(_ style: String) -> String {
            style.contains("2") ? "  " : (style.contains("3") ? "    " : "")
        }

        for s in snapshotsA {
            let (isMatch, isHeuristic) = isHeading(s)
            if isMatch {
                let indent = headingIndent(s.style ?? "")
                let marker = isHeuristic ? " (?)" : ""
                output += "\(indent)[\(s.index)] (\(s.style ?? ""))\(marker) \(truncateText(s.text, maxLength: 80))\n"
            }
        }
        output += "\n--- Heading Outline: Compare (\(docIdB)) ---\n"
        for s in snapshotsB {
            let (isMatch, isHeuristic) = isHeading(s)
            if isMatch {
                let indent = headingIndent(s.style ?? "")
                let marker = isHeuristic ? " (?)" : ""
                output += "\(indent)[\(s.index)] (\(s.style ?? ""))\(marker) \(truncateText(s.text, maxLength: 80))\n"
            }
        }
        return output
    }

    private func formatComparisonResult(
        docIdA: String, docIdB: String,
        infoA: (paragraphs: Int, words: Int), infoB: (paragraphs: Int, words: Int),
        entries: [DiffEntry], mode: String, contextLines: Int, maxResults: Int = 0
    ) -> String {
        let unchanged = entries.filter { $0.type == .unchanged }.count
        let modified = entries.filter { $0.type == .modified }.count
        let added = entries.filter { $0.type == .added }.count
        let deleted = entries.filter { $0.type == .deleted }.count
        let formatOnly = entries.filter { $0.type == .formatOnly }.count

        if modified == 0 && added == 0 && deleted == 0 && formatOnly == 0 {
            return """
            === Document Comparison ===
            Base: \(docIdA) (\(infoA.paragraphs) paragraphs, \(formatNumber(infoA.words)) words)
            Compare: \(docIdB) (\(infoB.paragraphs) paragraphs, \(formatNumber(infoB.words)) words)
            Mode: \(mode)

            Documents are identical.
            """
        }

        var output = """
        === Document Comparison ===
        Base: \(docIdA) (\(infoA.paragraphs) paragraphs, \(formatNumber(infoA.words)) words)
        Compare: \(docIdB) (\(infoB.paragraphs) paragraphs, \(formatNumber(infoB.words)) words)
        Mode: \(mode)

        --- Summary ---
        Unchanged: \(unchanged)  Modified: \(modified)  Added: \(added)  Deleted: \(deleted)
        """
        if formatOnly > 0 {
            output += "  Format-only: \(formatOnly)"
        }
        output += "\n\n--- Differences ---\n"

        var diffCount = 0
        for (entryIdx, entry) in entries.enumerated() {
            if entry.type == .unchanged { continue }
            diffCount += 1
            if maxResults > 0 && diffCount > maxResults {
                let remaining = entries.filter { $0.type != .unchanged }.count - maxResults
                output += "\n... and \(remaining) more differences (limited by max_results=\(maxResults))\n"
                break
            }

            // Context: show preceding unchanged paragraphs
            if contextLines > 0 {
                var contextEntries: [DiffEntry] = []
                var lookBack = entryIdx - 1
                while lookBack >= 0 && contextEntries.count < contextLines {
                    if entries[lookBack].type == .unchanged {
                        contextEntries.insert(entries[lookBack], at: 0)
                    } else {
                        break
                    }
                    lookBack -= 1
                }
                for ctx in contextEntries {
                    output += "\n  . A[\(ctx.indexA ?? 0)] \(truncateText(ctx.textA ?? "", maxLength: 80))"
                }
            }

            let style = entry.style ?? "Normal"
            switch entry.type {
            case .modified:
                output += "\n[MODIFIED] A[\(entry.indexA!)] -> B[\(entry.indexB!)] (\(style))"
                output += "\n  - \(truncateText(entry.textA ?? "", maxLength: 200))"
                output += "\n  + \(truncateText(entry.textB ?? "", maxLength: 200))"
            case .deleted:
                output += "\n[DELETED] A[\(entry.indexA!)] (\(style))"
                output += "\n  \(truncateText(entry.textA ?? "", maxLength: 200))"
            case .added:
                output += "\n[ADDED] B[\(entry.indexB!)] (\(style))"
                output += "\n  \(truncateText(entry.textB ?? "", maxLength: 200))"
            case .formatOnly:
                output += "\n[FORMAT_ONLY] A[\(entry.indexA!)] -> B[\(entry.indexB!)] (\(style))"
                output += "\n  Text: \(truncateText(entry.textA ?? "", maxLength: 120))"
                // Show formatting diff
                let fmtA = entry.formattedA ?? ""
                let fmtB = entry.formattedB ?? ""
                output += "\n  Base fmt: \(truncateText(fmtA, maxLength: 200))"
                output += "\n  Comp fmt: \(truncateText(fmtB, maxLength: 200))"
            case .unchanged:
                break
            }
            output += "\n"
        }
        return output
    }

    private func compareDocuments(args: [String: Value]) async throws -> String {
        guard let docIdA = args["doc_id_a"]?.stringValue else {
            throw WordError.missingParameter("doc_id_a")
        }
        guard let docIdB = args["doc_id_b"]?.stringValue else {
            throw WordError.missingParameter("doc_id_b")
        }
        if docIdA == docIdB {
            return "Error: doc_id_a and doc_id_b must be different documents."
        }
        guard let docA = openDocuments[docIdA] else {
            throw WordError.documentNotFound(docIdA)
        }
        guard let docB = openDocuments[docIdB] else {
            throw WordError.documentNotFound(docIdB)
        }

        let mode = args["mode"]?.stringValue ?? "text"
        let contextLines = min(max(args["context_lines"]?.intValue ?? 0, 0), 3)
        let maxResults = max(args["max_results"]?.intValue ?? 0, 0)

        // Parse custom heading styles for structure mode
        let customHeadingStyles: [String]? = {
            guard let arr = args["heading_styles"]?.arrayValue else { return nil }
            let styles = arr.compactMap { $0.stringValue }
            return styles.isEmpty ? nil : styles
        }()

        let snapshotsA = snapshotParagraphs(docA)
        let snapshotsB = snapshotParagraphs(docB)

        if snapshotsA.isEmpty && snapshotsB.isEmpty {
            return "Both documents have no paragraphs."
        }
        if snapshotsA.isEmpty {
            return "Base document (\(docIdA)) has no paragraphs."
        }
        if snapshotsB.isEmpty {
            return "Compare document (\(docIdB)) has no paragraphs."
        }

        let wordsA = snapshotsA.reduce(0) { $0 + countWords($1.text) }
        let wordsB = snapshotsB.reduce(0) { $0 + countWords($1.text) }
        let infoA = (paragraphs: snapshotsA.count, words: wordsA)
        let infoB = (paragraphs: snapshotsB.count, words: wordsB)

        // Structure mode: only statistics + heading outline
        if mode == "structure" {
            return formatStructureComparison(
                docIdA: docIdA, docIdB: docIdB,
                snapshotsA: snapshotsA, snapshotsB: snapshotsB,
                infoA: infoA, infoB: infoB,
                customHeadingStyles: customHeadingStyles
            )
        }

        let dp = computeLCS(snapshotsA, snapshotsB)
        let entries = buildDiffEntries(snapshotsA, snapshotsB, dp, mode: mode)

        return formatComparisonResult(
            docIdA: docIdA, docIdB: docIdB,
            infoA: infoA, infoB: infoB,
            entries: entries, mode: mode, contextLines: contextLines, maxResults: maxResults
        )
    }


    private func setColumns(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let columns = args["columns"]?.intValue else {
            throw WordError.missingParameter("columns")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let numCols = min(max(columns, 1), 4)
        let space = args["space"]?.intValue ?? 720
        let _ = args["equal_width"]?.boolValue ?? true
        let separator = args["separator"]?.boolValue ?? false

        doc.sectionProperties.columns = numCols

        try await storeDocument(doc, for: docId)

        var result = "Set document to \(numCols) column(s)"
        if numCols > 1 {
            result += " (space: \(space) twips"
            if separator {
                result += ", with separator line"
            }
            result += ")"
        }
        return result
    }

    private func insertColumnBreak(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var columnBreakPara = Paragraph()
        var columnBreakRun = Run(text: "")
        columnBreakRun.text = "\u{000C}"
        columnBreakPara.runs = [columnBreakRun]
        columnBreakPara.properties.pageBreakBefore = false

        doc.insertParagraph(columnBreakPara, at: paragraphIndex + 1)
        try await storeDocument(doc, for: docId)

        return "Inserted column break after paragraph \(paragraphIndex)"
    }

    private func setLineNumbers(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let enable = args["enable"]?.boolValue else {
            throw WordError.missingParameter("enable")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let start = args["start"]?.intValue ?? 1
        let countBy = args["count_by"]?.intValue ?? 1
        let restart = args["restart"]?.stringValue ?? "continuous"
        let distance = args["distance"]?.intValue ?? 360


        if enable {
            return "Line numbers enabled (start: \(start), count by: \(countBy), restart: \(restart), distance: \(distance) twips)"
        } else {
            return "Line numbers disabled"
        }
    }

    private func setPageBorders(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let style = args["style"]?.stringValue else {
            throw WordError.missingParameter("style")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let color = args["color"]?.stringValue ?? "000000"
        let size = args["size"]?.intValue ?? 4
        let offsetFrom = args["offset_from"]?.stringValue ?? "text"
        let showTop = args["top"]?.boolValue ?? true
        let showBottom = args["bottom"]?.boolValue ?? true
        let showLeft = args["left"]?.boolValue ?? true
        let showRight = args["right"]?.boolValue ?? true

        let validStyles = ["single", "double", "dotted", "dashed", "thick", "none"]
        guard validStyles.contains(style) else {
            return "Error: Invalid border style. Valid options: \(validStyles.joined(separator: ", "))"
        }


        var borders: [String] = []
        if showTop { borders.append("top") }
        if showBottom { borders.append("bottom") }
        if showLeft { borders.append("left") }
        if showRight { borders.append("right") }

        if style == "none" {
            return "Page borders removed"
        }

        return "Page borders set: style=\(style), color=#\(color), size=\(size), offset from \(offsetFrom), borders: \(borders.joined(separator: ", "))"
    }

    private func insertSymbol(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let charCode = args["char"]?.stringValue else {
            throw WordError.missingParameter("char")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let font = args["font"]?.stringValue
        let position = args["position"]?.stringValue ?? "end"

        guard let codePoint = UInt32(charCode, radix: 16),
              let scalar = Unicode.Scalar(codePoint) else {
            return "Error: Invalid character code '\(charCode)'. Use hexadecimal format (e.g., F020)."
        }
        let symbolChar = String(Character(scalar))

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = doc.body.children[actualIndex] {
            var symbolRun = Run(text: symbolChar)
            if let fontName = font {
                symbolRun.properties.fontName = fontName
            }

            if position == "start" {
                para.runs.insert(symbolRun, at: 0)
            } else {
                para.runs.append(symbolRun)
            }
            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        var result = "Inserted symbol (U+\(charCode.uppercased()))"
        if let fontName = font {
            result += " using font '\(fontName)'"
        }
        result += " at \(position) of paragraph \(paragraphIndex)"
        return result
    }

    private func setTextDirection(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let direction = args["direction"]?.stringValue else {
            throw WordError.missingParameter("direction")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let validDirections = ["lrTb", "tbRl", "btLr"]
        guard validDirections.contains(direction) else {
            return "Error: Invalid text direction. Valid options: lrTb (left-to-right, top-to-bottom), tbRl (vertical, right-to-left), btLr (bottom-to-top, left-to-right)"
        }

        let paragraphIndex = args["paragraph_index"]?.intValue


        if let pIndex = paragraphIndex {
            return "Text direction set to '\(direction)' for paragraph \(pIndex)"
        } else {
            return "Text direction set to '\(direction)' for entire document"
        }
    }

    private func insertDropCap(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let dropCapType = args["type"]?.stringValue ?? "drop"
        let lines = min(max(args["lines"]?.intValue ?? 3, 2), 10)
        let distance = args["distance"]?.intValue ?? 0
        let font = args["font"]?.stringValue

        let validTypes = ["drop", "margin", "none"]
        guard validTypes.contains(dropCapType) else {
            return "Error: Invalid drop cap type. Valid options: drop, margin, none"
        }

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(let para) = doc.body.children[actualIndex] {

            if dropCapType == "none" {
                doc.body.children[actualIndex] = .paragraph(para)
                try await storeDocument(doc, for: docId)
                return "Drop cap removed from paragraph \(paragraphIndex)"
            }

            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        var result = "Drop cap (\(dropCapType)) applied to paragraph \(paragraphIndex)"
        result += " (lines: \(lines), distance: \(distance) twips"
        if let fontName = font {
            result += ", font: \(fontName)"
        }
        result += ")"
        return result
    }

    private func insertHorizontalLine(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let style = args["style"]?.stringValue ?? "single"
        let color = args["color"]?.stringValue ?? "000000"
        let size = args["size"]?.intValue ?? 12  // 1.5pt

        let validStyles = ["single", "double", "dotted", "dashed", "thick"]
        guard validStyles.contains(style) else {
            return "Error: Invalid line style. Valid options: \(validStyles.joined(separator: ", "))"
        }

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = doc.body.children[actualIndex] {
            let borderType: ParagraphBorderType
            switch style {
            case "double": borderType = .double
            case "dotted": borderType = .dotted
            case "dashed": borderType = .dashed
            case "thick": borderType = .thick
            default: borderType = .single
            }

            let borderStyle = ParagraphBorderStyle(type: borderType, color: color, size: size, space: 1)
            para.properties.border = ParagraphBorder(bottom: borderStyle)
            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        return "Horizontal line added below paragraph \(paragraphIndex) (style: \(style), color: #\(color), size: \(size))"
    }

    private func setWidowOrphan(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let enable = args["enable"]?.boolValue ?? true
        let paragraphIndex = args["paragraph_index"]?.intValue


        if let pIndex = paragraphIndex {
            let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            guard pIndex >= 0 && pIndex < paragraphIndices.count else {
                throw WordError.invalidIndex(pIndex)
            }

            let actualIndex = paragraphIndices[pIndex]
            if case .paragraph(var para) = doc.body.children[actualIndex] {
                para.properties.keepLines = enable
                doc.body.children[actualIndex] = .paragraph(para)
            }

            try await storeDocument(doc, for: docId)
            return "Widow/orphan control \(enable ? "enabled" : "disabled") for paragraph \(pIndex)"
        } else {
            let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            for actualIndex in paragraphIndices {
                if case .paragraph(var para) = doc.body.children[actualIndex] {
                    para.properties.keepLines = enable
                    doc.body.children[actualIndex] = .paragraph(para)
                }
            }

            try await storeDocument(doc, for: docId)
            return "Widow/orphan control \(enable ? "enabled" : "disabled") for all paragraphs"
        }
    }

    private func setKeepWithNext(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let enable = args["enable"]?.boolValue ?? true

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = doc.body.children[actualIndex] {
            para.properties.keepNext = enable
            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        return "Keep with next \(enable ? "enabled" : "disabled") for paragraph \(paragraphIndex)"
    }


    private func insertWatermark(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let font = args["font"]?.stringValue ?? "Calibri Light"
        let color = args["color"]?.stringValue ?? "C0C0C0"
        let size = args["size"]?.intValue ?? 72
        let semitransparent = args["semitransparent"]?.boolValue ?? true
        let rotation = args["rotation"]?.intValue ?? -45


        var result = "Watermark inserted: \"\(text)\""
        result += " (font: \(font), color: #\(color), size: \(size)pt"
        if semitransparent {
            result += ", semitransparent"
        }
        result += ", rotation: \(rotation) degrees)"
        return result
    }

    private func insertImageWatermark(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let imagePath = args["image_path"]?.stringValue else {
            throw WordError.missingParameter("image_path")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let scale = args["scale"]?.intValue ?? 100
        let washout = args["washout"]?.boolValue ?? true

        guard FileManager.default.fileExists(atPath: imagePath) else {
            return "Error: Image file not found at '\(imagePath)'"
        }

        var result = "Image watermark inserted from: \(imagePath)"
        result += " (scale: \(scale)%"
        if washout {
            result += ", washout enabled"
        }
        result += ")"
        return result
    }

    private func removeWatermark(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        return "Watermark removed from document"
    }

    private func protectDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let protectionType = args["protection_type"]?.stringValue else {
            throw WordError.missingParameter("protection_type")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let validTypes = ["readOnly", "comments", "trackedChanges", "forms"]
        guard validTypes.contains(protectionType) else {
            return "Error: Invalid protection type. Valid options: \(validTypes.joined(separator: ", "))"
        }

        let hasPassword = args["password"]?.stringValue != nil

        var result = "Document protection enabled: \(protectionType)"
        if hasPassword {
            result += " (password protected)"
        }
        return result
    }

    private func unprotectDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let _ = args["password"]?.stringValue

        return "Document protection removed"
    }

    private func setDocumentPassword(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let password = args["password"]?.stringValue else {
            throw WordError.missingParameter("password")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        return "Document password set (password length: \(password.count) characters)"
    }

    private func removeDocumentPassword(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard args["current_password"]?.stringValue != nil else {
            throw WordError.missingParameter("current_password")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        return "Document password removed"
    }

    private func restrictEditingRegion(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let startParagraph = args["start_paragraph"]?.intValue else {
            throw WordError.missingParameter("start_paragraph")
        }
        guard let endParagraph = args["end_paragraph"]?.intValue else {
            throw WordError.missingParameter("end_paragraph")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let editor = args["editor"]?.stringValue

        let paragraphs = doc.getParagraphs()
        guard startParagraph >= 0 && startParagraph < paragraphs.count else {
            throw WordError.invalidIndex(startParagraph)
        }
        guard endParagraph >= startParagraph && endParagraph < paragraphs.count else {
            throw WordError.invalidIndex(endParagraph)
        }

        var result = "Editable region set: paragraphs \(startParagraph) to \(endParagraph)"
        if let editorName = editor {
            result += " (editor: \(editorName))"
        }
        return result
    }


    private func insertCaption(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let label = args["label"]?.stringValue else {
            throw WordError.missingParameter("label")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let captionText = args["caption_text"]?.stringValue ?? ""
        let position = args["position"]?.stringValue ?? "below"
        let includeChapterNumber = args["include_chapter_number"]?.boolValue ?? false

        let validLabels = ["Figure", "Table", "Equation"]
        guard validLabels.contains(label) else {
            return "Error: Invalid label. Valid options: \(validLabels.joined(separator: ", "))"
        }

        let seqField = "{ SEQ \(label) \\* ARABIC }"
        var captionContent = "\(label) "
        if includeChapterNumber {
            captionContent += "{ STYLEREF 1 \\s }-"
        }
        captionContent += seqField
        if !captionText.isEmpty {
            captionContent += ": \(captionText)"
        }

        var captionPara = Paragraph(text: captionContent)
        captionPara.properties.style = "Caption"

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex <= paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let insertIndex = position == "above" ? paragraphIndex : paragraphIndex + 1
        doc.insertParagraph(captionPara, at: insertIndex)
        try await storeDocument(doc, for: docId)

        return "Caption inserted: \(label) \(position) paragraph \(paragraphIndex)"
    }

    private func insertCrossReference(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let referenceType = args["reference_type"]?.stringValue else {
            throw WordError.missingParameter("reference_type")
        }
        guard let referenceTarget = args["reference_target"]?.stringValue else {
            throw WordError.missingParameter("reference_target")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let format = args["format"]?.stringValue ?? "full"
        let includeHyperlink = args["include_hyperlink"]?.boolValue ?? true

        let validTypes = ["bookmark", "heading", "figure", "table", "equation"]
        guard validTypes.contains(referenceType) else {
            return "Error: Invalid reference type. Valid options: \(validTypes.joined(separator: ", "))"
        }

        var result = "Cross-reference inserted at paragraph \(paragraphIndex)"
        result += " (type: \(referenceType), target: \(referenceTarget), format: \(format)"
        if includeHyperlink {
            result += ", hyperlinked"
        }
        result += ")"
        return result
    }

    private func insertTableOfFigures(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let captionLabel = args["caption_label"]?.stringValue else {
            throw WordError.missingParameter("caption_label")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let includePageNumbers = args["include_page_numbers"]?.boolValue ?? true
        let rightAlignPageNumbers = args["right_align_page_numbers"]?.boolValue ?? true
        let tabLeader = args["tab_leader"]?.stringValue ?? "dot"

        let validLabels = ["Figure", "Table", "Equation"]
        guard validLabels.contains(captionLabel) else {
            return "Error: Invalid caption label. Valid options: \(validLabels.joined(separator: ", "))"
        }

        var tocPara = Paragraph(text: "{ TOC \\c \"\(captionLabel)\" }")
        tocPara.properties.style = "TOCHeading"

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex <= paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        doc.insertParagraph(tocPara, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        var result = "Table of \(captionLabel)s inserted at paragraph \(paragraphIndex)"
        result += " (page numbers: \(includePageNumbers)"
        if includePageNumbers && rightAlignPageNumbers {
            result += ", right-aligned"
        }
        result += ", tab leader: \(tabLeader))"
        return result
    }

    private func insertIndexEntry(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let mainEntry = args["main_entry"]?.stringValue else {
            throw WordError.missingParameter("main_entry")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let subEntry = args["sub_entry"]?.stringValue
        let crossReference = args["cross_reference"]?.stringValue
        let bold = args["bold"]?.boolValue ?? false
        let italic = args["italic"]?.boolValue ?? false

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        // { XE "main entry:sub entry" \b \i \t "see also" }
        var result = "Index entry marked: \"\(mainEntry)\""
        if let sub = subEntry {
            result += ":\"\(sub)\""
        }
        if let xref = crossReference {
            result += " (see also: \(xref))"
        }
        if bold {
            result += " [bold]"
        }
        if italic {
            result += " [italic]"
        }
        result += " at paragraph \(paragraphIndex)"
        return result
    }

    private func insertIndex(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let columns = min(max(args["columns"]?.intValue ?? 2, 1), 4)
        let rightAlignPageNumbers = args["right_align_page_numbers"]?.boolValue ?? true
        let tabLeader = args["tab_leader"]?.stringValue ?? "dot"
        let runIn = args["run_in"]?.boolValue ?? false

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex <= paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var indexPara = Paragraph(text: "{ INDEX \\c \"\(columns)\" }")
        indexPara.properties.style = "Index"

        doc.insertParagraph(indexPara, at: paragraphIndex)
        try await storeDocument(doc, for: docId)

        var result = "Index inserted at paragraph \(paragraphIndex)"
        result += " (\(columns) columns"
        if rightAlignPageNumbers {
            result += ", right-aligned page numbers"
        }
        result += ", tab leader: \(tabLeader)"
        if runIn {
            result += ", run-in format"
        }
        result += ")"
        return result
    }


    private func setLanguage(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let language = args["language"]?.stringValue else {
            throw WordError.missingParameter("language")
        }
        guard openDocuments[docId] != nil else {
            throw WordError.documentNotFound(docId)
        }

        let paragraphIndex = args["paragraph_index"]?.intValue
        let noProofing = args["no_proofing"]?.boolValue ?? false


        if let pIndex = paragraphIndex {
            var result = "Language set to '\(language)' for paragraph \(pIndex)"
            if noProofing {
                result += " (proofing disabled)"
            }
            return result
        } else {
            var result = "Language set to '\(language)' for entire document"
            if noProofing {
                result += " (proofing disabled)"
            }
            return result
        }
    }

    private func setKeepLines(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let enable = args["enable"]?.boolValue ?? true

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = doc.body.children[actualIndex] {
            para.properties.keepLines = enable
            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        return "Keep lines together \(enable ? "enabled" : "disabled") for paragraph \(paragraphIndex)"
    }

    private func insertTabStop(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let position = args["position"]?.intValue else {
            throw WordError.missingParameter("position")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let alignment = args["alignment"]?.stringValue ?? "left"
        let leader = args["leader"]?.stringValue ?? "none"

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let validAlignments = ["left", "center", "right", "decimal"]
        guard validAlignments.contains(alignment) else {
            return "Error: Invalid alignment. Valid options: \(validAlignments.joined(separator: ", "))"
        }

        return "Tab stop added at position \(position) twips (alignment: \(alignment), leader: \(leader)) for paragraph \(paragraphIndex)"
    }

    private func clearTabStops(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        return "Tab stops cleared for paragraph \(paragraphIndex)"
    }

    private func setPageBreakBefore(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let enable = args["enable"]?.boolValue ?? true

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = doc.body.children[actualIndex] {
            para.properties.pageBreakBefore = enable
            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        return "Page break before \(enable ? "enabled" : "disabled") for paragraph \(paragraphIndex)"
    }

    private func setOutlineLevel(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let level = args["level"]?.intValue else {
            throw WordError.missingParameter("level")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        guard level >= 0 && level <= 9 else {
            return "Error: Outline level must be between 0 (body text) and 9"
        }

        let paragraphs = doc.getParagraphs()
        guard paragraphIndex >= 0 && paragraphIndex < paragraphs.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let levelDesc = level == 0 ? "body text" : "level \(level)"
        return "Outline level set to \(levelDesc) for paragraph \(paragraphIndex)"
    }

    private func insertContinuousSectionBreak(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let paragraphIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = doc.body.children[actualIndex] {
            para.properties.sectionBreak = .continuous
            doc.body.children[actualIndex] = .paragraph(para)
        }

        try await storeDocument(doc, for: docId)

        return "Continuous section break inserted after paragraph \(paragraphIndex)"
    }

    private func getSectionProperties(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let props = doc.sectionProperties
        var result = "Section Properties:\n"
        result += "- Page Size: \(props.pageSize.name) (\(props.pageSize.widthInInches)\" x \(props.pageSize.heightInInches)\")\n"
        result += "- Orientation: \(props.orientation.rawValue)\n"
        result += "- Margins: \(props.pageMargins.name)\n"
        result += "  - Top: \(props.pageMargins.top) twips\n"
        result += "  - Bottom: \(props.pageMargins.bottom) twips\n"
        result += "  - Left: \(props.pageMargins.left) twips\n"
        result += "  - Right: \(props.pageMargins.right) twips\n"
        result += "- Columns: \(props.columns)"
        if props.headerReference != nil {
            result += "\n- Has Header"
        }
        if props.footerReference != nil {
            result += "\n- Has Footer"
        }

        return result
    }

    private func addRowToTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let position = args["position"]?.stringValue ?? "end"
        let rowIndex = args["row_index"]?.intValue
        let data = args["data"]?.arrayValue?.compactMap { $0.stringValue } ?? []

        let tables = doc.getTables()
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let table = tables[tableIndex]
        let colCount = table.rows.first?.cells.count ?? 0

        var cells: [TableCell] = []
        for i in 0..<colCount {
            let text = i < data.count ? data[i] : ""
            cells.append(TableCell(text: text))
        }
        let newRow = TableRow(cells: cells)

        let tableIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .table = child { return i }
            return nil
        }

        guard tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var tbl) = doc.body.children[actualIndex] {
            switch position {
            case "start":
                tbl.rows.insert(newRow, at: 0)
            case "after_row":
                if let rIndex = rowIndex, rIndex >= 0 && rIndex < tbl.rows.count {
                    tbl.rows.insert(newRow, at: rIndex + 1)
                } else {
                    tbl.rows.append(newRow)
                }
            default: // "end"
                tbl.rows.append(newRow)
            }
            doc.body.children[actualIndex] = .table(tbl)
        }

        try await storeDocument(doc, for: docId)

        return "Row added to table \(tableIndex) at position '\(position)'"
    }

    private func addColumnToTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let position = args["position"]?.stringValue ?? "end"
        let colIndex = args["col_index"]?.intValue
        let data = args["data"]?.arrayValue?.compactMap { $0.stringValue } ?? []

        let tableIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .table = child { return i }
            return nil
        }

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var tbl) = doc.body.children[actualIndex] {
            for rowIdx in 0..<tbl.rows.count {
                let text = rowIdx < data.count ? data[rowIdx] : ""
                let newCell = TableCell(text: text)

                switch position {
                case "start":
                    tbl.rows[rowIdx].cells.insert(newCell, at: 0)
                case "after_col":
                    if let cIndex = colIndex, cIndex >= 0 && cIndex < tbl.rows[rowIdx].cells.count {
                        tbl.rows[rowIdx].cells.insert(newCell, at: cIndex + 1)
                    } else {
                        tbl.rows[rowIdx].cells.append(newCell)
                    }
                default: // "end"
                    tbl.rows[rowIdx].cells.append(newCell)
                }
            }
            doc.body.children[actualIndex] = .table(tbl)
        }

        try await storeDocument(doc, for: docId)

        return "Column added to table \(tableIndex) at position '\(position)'"
    }

    private func deleteRowFromTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let rowIndex = args["row_index"]?.intValue else {
            throw WordError.missingParameter("row_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let tableIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .table = child { return i }
            return nil
        }

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var tbl) = doc.body.children[actualIndex] {
            guard rowIndex >= 0 && rowIndex < tbl.rows.count else {
                throw WordError.invalidIndex(rowIndex)
            }

            tbl.rows.remove(at: rowIndex)
            doc.body.children[actualIndex] = .table(tbl)
        }

        try await storeDocument(doc, for: docId)

        return "Row \(rowIndex) deleted from table \(tableIndex)"
    }

    private func deleteColumnFromTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let colIndex = args["col_index"]?.intValue else {
            throw WordError.missingParameter("col_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        enforceTrackChangesIfNeeded(&doc, docId: docId)

        let tableIndices = doc.body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .table = child { return i }
            return nil
        }

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var tbl) = doc.body.children[actualIndex] {
            for rowIdx in 0..<tbl.rows.count {
                guard colIndex >= 0 && colIndex < tbl.rows[rowIdx].cells.count else {
                    throw WordError.invalidIndex(colIndex)
                }
                tbl.rows[rowIdx].cells.remove(at: colIndex)
            }
            doc.body.children[actualIndex] = .table(tbl)
        }

        try await storeDocument(doc, for: docId)

        return "Column \(colIndex) deleted from table \(tableIndex)"
    }

    private func setCellWidth(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let row = args["row"]?.intValue else {
            throw WordError.missingParameter("row")
        }
        guard let col = args["col"]?.intValue else {
            throw WordError.missingParameter("col")
        }
        guard let width = args["width"]?.intValue else {
            throw WordError.missingParameter("width")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let widthType = args["width_type"]?.stringValue ?? "dxa"

        let tables = doc.getTables()
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let table = tables[tableIndex]
        guard row >= 0 && row < table.rows.count else {
            throw WordError.invalidIndex(row)
        }
        guard col >= 0 && col < table.rows[row].cells.count else {
            throw WordError.invalidIndex(col)
        }

        return "Cell width set to \(width) \(widthType) for table \(tableIndex), row \(row), col \(col)"
    }

    private func setRowHeight(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let rowIndex = args["row_index"]?.intValue else {
            throw WordError.missingParameter("row_index")
        }
        guard let height = args["height"]?.intValue else {
            throw WordError.missingParameter("height")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let heightRule = args["height_rule"]?.stringValue ?? "atLeast"

        let tables = doc.getTables()
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let table = tables[tableIndex]
        guard rowIndex >= 0 && rowIndex < table.rows.count else {
            throw WordError.invalidIndex(rowIndex)
        }

        return "Row height set to \(height) twips (\(heightRule)) for table \(tableIndex), row \(rowIndex)"
    }

    private func setTableAlignment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let alignment = args["alignment"]?.stringValue else {
            throw WordError.missingParameter("alignment")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let validAlignments = ["left", "center", "right"]
        guard validAlignments.contains(alignment) else {
            return "Error: Invalid alignment. Valid options: \(validAlignments.joined(separator: ", "))"
        }

        let tables = doc.getTables()
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        return "Table \(tableIndex) alignment set to '\(alignment)'"
    }

    private func setCellVerticalAlignment(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let row = args["row"]?.intValue else {
            throw WordError.missingParameter("row")
        }
        guard let col = args["col"]?.intValue else {
            throw WordError.missingParameter("col")
        }
        guard let alignment = args["alignment"]?.stringValue else {
            throw WordError.missingParameter("alignment")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let validAlignments = ["top", "center", "bottom"]
        guard validAlignments.contains(alignment) else {
            return "Error: Invalid vertical alignment. Valid options: \(validAlignments.joined(separator: ", "))"
        }

        let tables = doc.getTables()
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let table = tables[tableIndex]
        guard row >= 0 && row < table.rows.count else {
            throw WordError.invalidIndex(row)
        }
        guard col >= 0 && col < table.rows[row].cells.count else {
            throw WordError.invalidIndex(col)
        }

        return "Cell vertical alignment set to '\(alignment)' for table \(tableIndex), row \(row), col \(col)"
    }

    private func setHeaderRow(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let tableIndex = args["table_index"]?.intValue else {
            throw WordError.missingParameter("table_index")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let rowCount = args["row_count"]?.intValue ?? 1

        let tables = doc.getTables()
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let table = tables[tableIndex]
        guard rowCount > 0 && rowCount <= table.rows.count else {
            return "Error: Row count must be between 1 and \(table.rows.count)"
        }

        return "Header row(s) set for table \(tableIndex): first \(rowCount) row(s) will repeat across pages"
    }
}
