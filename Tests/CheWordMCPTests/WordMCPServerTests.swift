import XCTest
import MCP
import OOXMLSwift
@testable import CheWordMCP

final class WordMCPServerTests: XCTestCase {
    private func tempURL(_ suffix: String = UUID().uuidString) -> URL {
        FileManager.default.temporaryDirectory.appendingPathComponent("cheword-\(suffix).docx")
    }

    private func writeDocument(text: String, to url: URL) throws {
        var document = WordDocument()
        document.appendParagraph(Paragraph(text: text))
        try DocxWriter.write(document, to: url)
    }

    private func writeDocument(_ document: WordDocument, to url: URL) throws {
        try DocxWriter.write(document, to: url)
    }

    private func readDocumentText(from url: URL) throws -> String {
        try DocxReader.read(from: url).getText()
    }

    private func resultText(_ result: CallTool.Result) -> String {
        guard let first = result.content.first else { return "" }
        switch first {
        case .text(let text):
            return text
        default:
            return ""
        }
    }

    func testCreateDocumentEnablesTrackChangesByDefault() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(server.isTrackChangesEnabledForTesting("doc"), true)
    }

    func testDirtyDocumentCannotCloseWithoutSave() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Hello from test")
            ]
        )

        let closeResult = await server.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(closeResult.isError, true)
        XCTAssertTrue(resultText(closeResult).contains("unsaved changes"))
        XCTAssertTrue(server.isDocumentDirtyForTesting("doc"))
    }

    func testDuplicateDocIdIsRejectedBeforeOverwrite() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Unsaved draft")
            ]
        )

        let duplicateCreate = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(duplicateCreate.isError, true)
        XCTAssertTrue(resultText(duplicateCreate).contains("Document already open"))
        XCTAssertTrue(server.isDocumentDirtyForTesting("doc"))
    }

    func testNewDocumentRequiresExplicitPathForFirstSave() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Hello")
            ]
        )

        let saveResult = await server.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(saveResult.isError, true)
        XCTAssertTrue(resultText(saveResult).contains("No path was provided"))
        XCTAssertTrue(server.isDocumentDirtyForTesting("doc"))
    }

    func testSaveDocumentWithoutPathUsesOriginalOpenedPath() async throws {
        let url = tempURL("save-fallback")
        try writeDocument(text: "Before", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("After")
            ]
        )

        let saveResult = await server.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(saveResult.isError, nil)
        XCTAssertTrue(resultText(saveResult).contains("original path"))
        XCTAssertTrue(try readDocumentText(from: url).contains("After"))
        XCTAssertFalse(server.isDocumentDirtyForTesting("doc"))
    }

    func testAutosavePersistsMutationsImmediatelyWhenEnabled() async throws {
        let url = tempURL("autosave")
        try writeDocument(text: "Start", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )

        let insertResult = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Autosaved line")
            ]
        )

        XCTAssertEqual(insertResult.isError, nil)
        XCTAssertTrue(try readDocumentText(from: url).contains("Autosaved line"))
        XCTAssertFalse(server.isDocumentDirtyForTesting("doc"))
    }

    func testGetDocumentSessionStateReportsDirtyAndFinalizeReadiness() async throws {
        let url = tempURL("session-state")
        try writeDocument(text: "State doc", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Dirty edit")
            ]
        )

        let stateResult = await server.invokeToolForTesting(
            name: "get_document_session_state",
            arguments: ["doc_id": .string("doc")]
        )

        let text = resultText(stateResult)
        XCTAssertEqual(stateResult.isError, nil)
        XCTAssertTrue(text.contains("Dirty: true"))
        XCTAssertTrue(text.contains("Track changes enabled: true"))
        XCTAssertTrue(text.contains("Save without explicit path available: true"))
        XCTAssertTrue(text.contains("Close without save allowed: false"))
        XCTAssertTrue(text.contains("Finalize without explicit path available: true"))
    }

    func testGetDocumentSessionStateForNewDocumentRequiresExplicitSavePath() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )

        let stateResult = await server.invokeToolForTesting(
            name: "get_document_session_state",
            arguments: ["doc_id": .string("doc")]
        )

        let text = resultText(stateResult)
        XCTAssertEqual(stateResult.isError, nil)
        XCTAssertTrue(text.contains("Original path: (none)"))
        XCTAssertTrue(text.contains("Save without explicit path available: false"))
        XCTAssertTrue(text.contains("Finalize without explicit path available: false"))
    }

    func testShutdownFlushPersistsDirtyOpenedDocuments() async throws {
        let url = tempURL("flush")
        try writeDocument(text: "Original", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Needs flush")
            ]
        )

        XCTAssertTrue(server.isDocumentDirtyForTesting("doc"))

        await server.flushDirtyDocumentsForTesting()

        XCTAssertFalse(server.isDocumentDirtyForTesting("doc"))
        XCTAssertTrue(try readDocumentText(from: url).contains("Needs flush"))
    }

    func testShutdownFlushSkipsDirtyNewDocumentWithoutKnownPath() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Unsaved new doc")
            ]
        )

        await server.flushDirtyDocumentsForTesting()

        XCTAssertTrue(server.isDocumentDirtyForTesting("doc"))
    }

    func testFinalizeDocumentSavesAndClosesUsingOriginalPath() async throws {
        let url = tempURL("finalize-opened")
        try writeDocument(text: "Before finalize", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Finalize me")
            ]
        )

        let finalizeResult = await server.invokeToolForTesting(
            name: "finalize_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(finalizeResult.isError, nil)
        XCTAssertTrue(resultText(finalizeResult).contains("Finalized document"))
        XCTAssertNil(server.isTrackChangesEnabledForTesting("doc"))
        XCTAssertTrue(try readDocumentText(from: url).contains("Finalize me"))
    }

    func testFinalizeDocumentRequiresExplicitPathForNewDocument() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "text": .string("Needs first path")
            ]
        )

        let finalizeResult = await server.invokeToolForTesting(
            name: "finalize_document",
            arguments: ["doc_id": .string("doc")]
        )

        XCTAssertEqual(finalizeResult.isError, true)
        XCTAssertTrue(resultText(finalizeResult).contains("No path was provided"))
        XCTAssertTrue(server.isDocumentDirtyForTesting("doc"))
    }

    func testInsertTextPreservesUnaffectedRuns() async throws {
        let url = tempURL("insert-inline")
        let document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello "),
                TestFixtures.makeRun("world", bold: true)
            ])
        ])
        try writeDocument(document, to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )

        let insertResult = await server.invokeToolForTesting(
            name: "insert_text",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "position": .int(6),
                "text": .string("brave ")
            ]
        )

        XCTAssertEqual(insertResult.isError, nil)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "brave ", "world"])
        XCTAssertFalse(runs[0].properties.bold)
        XCTAssertFalse(runs[1].properties.bold)
        XCTAssertTrue(runs[2].properties.bold)
    }

    func testReplaceTextRangeToolPreservesFormattingOutsideTarget() async throws {
        let url = tempURL("replace-range")
        let document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello "),
                TestFixtures.makeRun("world", bold: true),
                TestFixtures.makeRun("!")
            ])
        ])
        try writeDocument(document, to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )

        let replaceResult = await server.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "replacement": .string("team")
            ]
        )

        XCTAssertEqual(replaceResult.isError, nil)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "team", "!"])
        XCTAssertFalse(runs[0].properties.bold)
        XCTAssertTrue(runs[1].properties.bold)
        XCTAssertFalse(runs[2].properties.bold)
    }

    func testFormatTextRangeToolOnlyFormatsTargetSpan() async throws {
        let url = tempURL("format-range")
        try writeDocument(text: "Hello world", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )

        let formatResult = await server.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "bold": .bool(true),
                "color": .string("FF0000")
            ]
        )

        XCTAssertEqual(formatResult.isError, nil)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "world"])
        XCTAssertFalse(runs[0].properties.bold)
        XCTAssertTrue(runs[1].properties.bold)
        XCTAssertEqual(runs[1].properties.color, "FF0000")
    }
}
