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
}
