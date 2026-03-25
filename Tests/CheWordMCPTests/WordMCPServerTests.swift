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

    func testFormatTextRangeToolCanSetHighlight() async throws {
        let url = tempURL("format-range-highlight")
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
                "highlight": .string("yellow")
            ]
        )

        XCTAssertNil(formatResult.isError)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "world"])
        XCTAssertNil(runs[0].properties.highlight)
        XCTAssertEqual(runs[1].properties.highlight, .yellow)
    }

    func testFormatTextRangeToolCanClearHighlight() async throws {
        let url = tempURL("format-range-clear-highlight")
        let document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello "),
                TestFixtures.makeRun("world", highlight: .yellow)
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

        let formatResult = await server.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "highlight": .string("none")
            ]
        )

        XCTAssertNil(formatResult.isError)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "world"])
        XCTAssertNil(runs[1].properties.highlight)
    }

    func testFormatTextCanClearHighlightAcrossParagraph() async throws {
        let url = tempURL("format-paragraph-clear-highlight")
        let document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello", highlight: .yellow),
                TestFixtures.makeRun(" world", highlight: .yellow)
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

        let formatResult = await server.invokeToolForTesting(
            name: "format_text",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "highlight": .string("none")
            ]
        )

        XCTAssertNil(formatResult.isError)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertTrue(runs.allSatisfy { $0.properties.highlight == nil })
    }

    func testFormatTextRangeRejectsUnsupportedHighlightValue() async throws {
        let url = tempURL("format-range-bad-highlight")
        try writeDocument(text: "Hello world", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )

        let formatResult = await server.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "highlight": .string("purple")
            ]
        )

        XCTAssertEqual(formatResult.isError, true)
        XCTAssertTrue(resultText(formatResult).contains("unsupported highlight 'purple'"))
    }

    func testReplyToCommentAcceptsCanonicalParameterNames() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("doc")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("doc"), "text": .string("Comment target")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_comment",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "author": .string("Reviewer"),
                "text": .string("Needs work")
            ]
        )

        let replyResult = await server.invokeToolForTesting(
            name: "reply_to_comment",
            arguments: [
                "doc_id": .string("doc"),
                "parent_comment_id": .int(1),
                "text": .string("Handled"),
                "author": .string("Author")
            ]
        )

        XCTAssertNil(replyResult.isError)

        let commentsResult = await server.invokeToolForTesting(
            name: "list_comments",
            arguments: ["doc_id": .string("doc")]
        )
        let commentsText = resultText(commentsResult)
        XCTAssertTrue(commentsText.contains("Handled"))
        XCTAssertTrue(commentsText.contains("Author"))
    }

    func testReplyToCommentSupportsLegacyAliasParameters() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("doc")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("doc"), "text": .string("Comment target")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_comment",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "author": .string("Reviewer"),
                "text": .string("Needs work")
            ]
        )

        let replyResult = await server.invokeToolForTesting(
            name: "reply_to_comment",
            arguments: [
                "doc_id": .string("doc"),
                "comment_id": .int(1),
                "reply_text": .string("Legacy handled")
            ]
        )

        XCTAssertNil(replyResult.isError)
        XCTAssertTrue(resultText(replyResult).contains("reply ID"))
        XCTAssertTrue(resultText(replyResult).contains("by Author"))
    }

    func testFormatTextRangeRejectsEmptyParagraphOutOfRange() async throws {
        let url = tempURL("empty-para")
        var document = WordDocument()
        document.appendParagraph(Paragraph(text: ""))
        try writeDocument(document, to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )

        let formatResult = await server.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(0),
                "end": .int(50),
                "bold": .bool(true)
            ]
        )

        XCTAssertEqual(formatResult.isError, true)
        XCTAssertTrue(resultText(formatResult).contains("Invalid text range start=0 end=50 for visible paragraph length 0"))
    }

    func testGetRevisionsReturnsNativeRevisionsAfterEdit() async throws {
        let url = tempURL("revision-open")
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
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0),
                "text": .string("Changed")
            ]
        )

        let revisionsResult = await server.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("doc")]
        )

        let revisionsText = resultText(revisionsResult)
        XCTAssertNil(revisionsResult.isError)
        XCTAssertTrue(revisionsText.contains("Original"))
        XCTAssertTrue(revisionsText.contains("Revisions in document"))
        XCTAssertTrue(revisionsText.contains("Changed"))
    }

    func testAcceptRevisionWithSinglePendingRevisionKeepsDocumentState() async throws {
        let url = tempURL("accept-single")
        try writeDocument(text: "Original", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0),
                "text": .string("Accepted")
            ]
        )

        let acceptResult = await server.invokeToolForTesting(
            name: "accept_revision",
            arguments: [
                "doc_id": .string("doc"),
                "revision_id": .int(1)
            ]
        )

        XCTAssertNil(acceptResult.isError)
        XCTAssertEqual(try readDocumentText(from: url), "Accepted")

        let revisionsResult = await server.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("No revisions"))
    }

    func testRejectRevisionWithSinglePendingRevisionRestoresPriorDocumentState() async throws {
        let url = tempURL("reject-single")
        try writeDocument(text: "Original", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0),
                "text": .string("Rejected")
            ]
        )

        let rejectResult = await server.invokeToolForTesting(
            name: "reject_revision",
            arguments: [
                "doc_id": .string("doc"),
                "revision_id": .int(1)
            ]
        )

        XCTAssertNil(rejectResult.isError)
        XCTAssertEqual(try readDocumentText(from: url), "Original")

        let revisionsResult = await server.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("No revisions"))
    }

    func testAcceptAndRejectAllUseNativeRevisions() async throws {
        let acceptURL = tempURL("accept-all")
        try writeDocument(text: "Base", to: acceptURL)

        let acceptingServer = await WordMCPServer()
        _ = await acceptingServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(acceptURL.path),
                "doc_id": .string("accept"),
                "autosave": .bool(true)
            ]
        )
        _ = await acceptingServer.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("accept"),
                "index": .int(0),
                "text": .string("Base updated")
            ]
        )
        _ = await acceptingServer.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("accept"),
                "text": .string("Second line")
            ]
        )

        let acceptAllResult = await acceptingServer.invokeToolForTesting(
            name: "accept_all_revisions",
            arguments: ["doc_id": .string("accept")]
        )

        XCTAssertNil(acceptAllResult.isError)
        XCTAssertTrue(resultText(acceptAllResult).contains("Accepted 3 revision"))
        XCTAssertTrue(try readDocumentText(from: acceptURL).contains("Second line"))

        let rejectURL = tempURL("reject-all")
        try writeDocument(text: "Base", to: rejectURL)

        let rejectingServer = await WordMCPServer()
        _ = await rejectingServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(rejectURL.path),
                "doc_id": .string("reject"),
                "autosave": .bool(true)
            ]
        )
        _ = await rejectingServer.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("reject"),
                "index": .int(0),
                "text": .string("Base updated")
            ]
        )
        _ = await rejectingServer.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("reject"),
                "text": .string("Second line")
            ]
        )

        let rejectAllResult = await rejectingServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reject")]
        )

        XCTAssertNil(rejectAllResult.isError)
        XCTAssertTrue(resultText(rejectAllResult).contains("Rejected 3 revision"))
        XCTAssertEqual(try readDocumentText(from: rejectURL), "Base")
    }

    func testNativeRevisionsPersistAfterSaveAndReopen() async throws {
        let url = tempURL("native-revisions")
        try writeDocument(text: "Original", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0),
                "text": .string("Changed")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened")
            ]
        )

        let revisionsResult = await secondServer.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("reopened")]
        )

        XCTAssertNil(revisionsResult.isError)
        XCTAssertTrue(resultText(revisionsResult).contains("Revisions in document"))
        XCTAssertTrue(resultText(revisionsResult).contains("Original"))
        XCTAssertTrue(resultText(revisionsResult).contains("Changed"))
    }

    func testRejectAllRevisionsAfterReopenRestoresOriginalText() async throws {
        let url = tempURL("reject-after-reopen")
        try writeDocument(text: "Original", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0),
                "text": .string("Changed")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened"),
                "autosave": .bool(true)
            ]
        )

        let rejectResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reopened")]
        )

        XCTAssertNil(rejectResult.isError)
        XCTAssertEqual(try readDocumentText(from: url), "Original")
    }

    func testDeleteParagraphHidesTrackedDeletedParagraphFromListing() async throws {
        let url = tempURL("delete-hidden")
        var document = WordDocument()
        document.appendParagraph(Paragraph(text: "First"))
        document.appendParagraph(Paragraph(text: "Second"))
        try writeDocument(document, to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )

        let deleteResult = await server.invokeToolForTesting(
            name: "delete_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0)
            ]
        )
        XCTAssertNil(deleteResult.isError)

        let paragraphsResult = await server.invokeToolForTesting(
            name: "get_paragraphs",
            arguments: ["doc_id": .string("doc")]
        )
        let text = resultText(paragraphsResult)
        XCTAssertTrue(text.contains("[0] (Normal) Second"))
        XCTAssertFalse(text.contains("[1]"))
        XCTAssertFalse(text.contains("First"))
    }

    func testTrackedDeleteStaysHiddenAfterReopenWhileRevisionRemainsVisible() async throws {
        let url = tempURL("delete-reopen-hidden")
        var document = WordDocument()
        document.appendParagraph(Paragraph(text: "First"))
        document.appendParagraph(Paragraph(text: "Second"))
        try writeDocument(document, to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "delete_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0)
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened")
            ]
        )

        let paragraphsResult = await secondServer.invokeToolForTesting(
            name: "get_paragraphs",
            arguments: ["doc_id": .string("reopened")]
        )
        let paragraphText = resultText(paragraphsResult)
        XCTAssertTrue(paragraphText.contains("[0] (Normal) Second"))
        XCTAssertFalse(paragraphText.contains("[1]"))
        XCTAssertFalse(paragraphText.contains("First"))

        let revisionsResult = await secondServer.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("First"))
    }

    func testRejectAndAcceptAllRevisionsHandleTrackedDeletedParagraphWithoutArtifacts() async throws {
        let rejectURL = tempURL("delete-reject-all")
        var rejectDocument = WordDocument()
        rejectDocument.appendParagraph(Paragraph(text: "First"))
        rejectDocument.appendParagraph(Paragraph(text: "Second"))
        try writeDocument(rejectDocument, to: rejectURL)

        let rejectServer = await WordMCPServer()
        _ = await rejectServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(rejectURL.path),
                "doc_id": .string("reject"),
                "autosave": .bool(true)
            ]
        )
        _ = await rejectServer.invokeToolForTesting(
            name: "delete_paragraph",
            arguments: [
                "doc_id": .string("reject"),
                "index": .int(0)
            ]
        )

        let rejectAllResult = await rejectServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reject")]
        )
        XCTAssertNil(rejectAllResult.isError)

        let rejected = try DocxReader.read(from: rejectURL)
        XCTAssertEqual(rejected.getParagraphs().map { $0.getText() }, ["First", "Second"])

        let acceptURL = tempURL("delete-accept-all")
        var acceptDocument = WordDocument()
        acceptDocument.appendParagraph(Paragraph(text: "First"))
        acceptDocument.appendParagraph(Paragraph(text: "Second"))
        try writeDocument(acceptDocument, to: acceptURL)

        let acceptServer = await WordMCPServer()
        _ = await acceptServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(acceptURL.path),
                "doc_id": .string("accept"),
                "autosave": .bool(true)
            ]
        )
        _ = await acceptServer.invokeToolForTesting(
            name: "delete_paragraph",
            arguments: [
                "doc_id": .string("accept"),
                "index": .int(0)
            ]
        )

        let acceptAllResult = await acceptServer.invokeToolForTesting(
            name: "accept_all_revisions",
            arguments: ["doc_id": .string("accept")]
        )
        XCTAssertNil(acceptAllResult.isError)

        let accepted = try DocxReader.read(from: acceptURL)
        XCTAssertEqual(accepted.getParagraphs().map { $0.getText() }, ["Second"])
    }

    func testListCommentsShowsReplyParentIdsInsteadOfParaMinusOne() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("doc")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("doc"), "text": .string("Comment target")]
        )
        _ = await server.invokeToolForTesting(
            name: "insert_comment",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "author": .string("Reviewer"),
                "text": .string("Needs work")
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "reply_to_comment",
            arguments: [
                "doc_id": .string("doc"),
                "parent_comment_id": .int(1),
                "text": .string("Handled"),
                "author": .string("Author")
            ]
        )

        let commentsResult = await server.invokeToolForTesting(
            name: "list_comments",
            arguments: ["doc_id": .string("doc")]
        )

        let commentsText = resultText(commentsResult)
        XCTAssertTrue(commentsText.contains("reply to 1"))
        XCTAssertFalse(commentsText.contains("para -1"))
    }

    func testReplaceTextRangePersistsNativeRevisionAfterReopenAndRejectAllRestoresText() async throws {
        let url = tempURL("replace-native-reopen")
        try writeDocument(text: "Hello world", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "replacement": .string("team")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened"),
                "autosave": .bool(true)
            ]
        )
        let revisionsResult = await secondServer.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("team"))

        let rejectResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertNil(rejectResult.isError)
        XCTAssertEqual(try readDocumentText(from: url), "Hello world")
    }

    func testReplaceTextRangeTwiceInSameParagraphRejectAllRestoresOriginalText() async throws {
        let url = tempURL("replace-twice-native-reopen")
        try writeDocument(text: "Hello world", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "replacement": .string("team")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(0),
                "end": .int(5),
                "replacement": .string("Hi")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened"),
                "autosave": .bool(true)
            ]
        )

        let rejectResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reopened")]
        )

        XCTAssertNil(rejectResult.isError)
        XCTAssertEqual(try readDocumentText(from: url), "Hello world")
    }

    func testFormatTextRangePersistsNativeRevisionAfterReopenAndRejectAllRestoresFormatting() async throws {
        let url = tempURL("format-native-reopen")
        try writeDocument(text: "Hello world", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
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
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened"),
                "autosave": .bool(true)
            ]
        )
        let revisionsResult = await secondServer.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("RPRCHANGE"))

        let rejectResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertNil(rejectResult.isError)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello world"])
        XCTAssertFalse(runs[0].properties.bold)
        XCTAssertNil(runs[0].properties.color)
    }

    func testFormatTextRangeAfterTrackedReplaceRejectAllRestoresOriginalState() async throws {
        let url = tempURL("format-after-replace-native-reopen")
        try writeDocument(text: "Hello world", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "replacement": .string("team")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(10),
                "bold": .bool(true),
                "color": .string("FF0000")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened"),
                "autosave": .bool(true)
            ]
        )

        let rejectResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reopened")]
        )

        XCTAssertNil(rejectResult.isError)

        let saved = try DocxReader.read(from: url)
        let runs = saved.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text).joined(), "Hello world")
        XCTAssertTrue(runs.allSatisfy { !$0.properties.bold })
        XCTAssertTrue(runs.allSatisfy { $0.properties.color == nil })
    }

    func testFormatTextRangeHighlightSetAndClearPersistsAcrossSaveReopen() async throws {
        let url = tempURL("format-highlight-native-reopen")
        let document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello "),
                TestFixtures.makeRun("world", highlight: .yellow)
            ])
        ])
        try writeDocument(document, to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(11),
                "highlight": .string("none")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("doc")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("reopened"),
                "autosave": .bool(true)
            ]
        )
        let revisionsResult = await secondServer.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("RPRCHANGE"))

        let savedBeforeReject = try DocxReader.read(from: url)
        XCTAssertNil(savedBeforeReject.getParagraphs()[0].runs[1].properties.highlight)

        let rejectResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("reopened")]
        )
        XCTAssertNil(rejectResult.isError)

        let savedAfterReject = try DocxReader.read(from: url)
        XCTAssertEqual(savedAfterReject.getParagraphs()[0].runs[1].properties.highlight, .yellow)
    }

    func testExistingNativeRevisionsRemainVisibleAfterNewEditAndRejectAllRestoresBaseline() async throws {
        let url = tempURL("mixed-native-revisions")
        try writeDocument(text: "Baseline", to: url)

        let firstServer = await WordMCPServer()
        _ = await firstServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("first")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("first"),
                "index": .int(0),
                "text": .string("First edit")
            ]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("first")]
        )
        _ = await firstServer.invokeToolForTesting(
            name: "close_document",
            arguments: ["doc_id": .string("first")]
        )

        let secondServer = await WordMCPServer()
        _ = await secondServer.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("second"),
                "autosave": .bool(true)
            ]
        )
        _ = await secondServer.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: [
                "doc_id": .string("second"),
                "text": .string("Second edit")
            ]
        )

        let revisionsResult = await secondServer.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("second")]
        )
        let revisionsText = resultText(revisionsResult)
        XCTAssertTrue(revisionsText.contains("Baseline"))
        XCTAssertTrue(revisionsText.contains("First edit"))
        XCTAssertTrue(revisionsText.contains("Second edit"))

        let rejectAllResult = await secondServer.invokeToolForTesting(
            name: "reject_all_revisions",
            arguments: ["doc_id": .string("second")]
        )
        XCTAssertNil(rejectAllResult.isError)
        XCTAssertEqual(try readDocumentText(from: url), "Baseline")
    }

    func testReplaceTextRejectsEmptyFindString() async throws {
        let url = tempURL("replace-empty-find")
        try writeDocument(text: "Hello world", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc")
            ]
        )

        let replaceResult = await server.invokeToolForTesting(
            name: "replace_text",
            arguments: [
                "doc_id": .string("doc"),
                "find": .string(""),
                "replace": .string("ignored")
            ]
        )

        XCTAssertEqual(replaceResult.isError, true)
        XCTAssertTrue(resultText(replaceResult).contains("find"))
    }

    func testReplaceTextUsesVisibleTrackedTextAcrossMultipleEditsWithoutReopen() async throws {
        let url = tempURL("replace-visible-tracked")
        try writeDocument(text: "Alpha Beta Gamma", to: url)

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(url.path),
                "doc_id": .string("doc"),
                "autosave": .bool(true)
            ]
        )

        let firstReplace = await server.invokeToolForTesting(
            name: "replace_text",
            arguments: [
                "doc_id": .string("doc"),
                "find": .string("Beta"),
                "replace": .string("LongerBeta")
            ]
        )
        XCTAssertNil(firstReplace.isError)

        let secondReplace = await server.invokeToolForTesting(
            name: "replace_text",
            arguments: [
                "doc_id": .string("doc"),
                "find": .string("Gamma"),
                "replace": .string("Done")
            ]
        )
        XCTAssertNil(secondReplace.isError)

        let saved = try DocxReader.read(from: url)
        let savedText = saved.getParagraphs()[0].getText()
        XCTAssertTrue(savedText.contains("LongerBeta"))
        XCTAssertTrue(savedText.contains("Done"))
        XCTAssertFalse(savedText.contains("Gamma"))
    }

    func testReplaceTextKeepsTypographicMatchingLiteral() async throws {
        let url = tempURL("replace-literal-dash")
        try writeDocument(text: "between 35 kDa – 100 kDa", to: url)

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
            name: "replace_text",
            arguments: [
                "doc_id": .string("doc"),
                "find": .string("35 kDa - 100 kDa"),
                "replace": .string("range updated")
            ]
        )

        XCTAssertNil(replaceResult.isError)
        XCTAssertTrue(resultText(replaceResult).contains("Replaced 0 occurrence"))
        XCTAssertEqual(try readDocumentText(from: url), "between 35 kDa – 100 kDa")
    }

    func testReplaceTextStillWorksWhenParagraphContainsHyperlink() async throws {
        let url = tempURL("replace-hyperlink")
        var document = WordDocument()
        let hyperlink = Hyperlink.external(
            id: "link-1",
            text: "docs",
            url: "https://example.com",
            relationshipId: "rId1"
        )
        var paragraph = Paragraph(runs: [Run(text: "See ")])
        paragraph.hyperlinks = [hyperlink]
        document.appendParagraph(paragraph)
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
            name: "replace_text",
            arguments: [
                "doc_id": .string("doc"),
                "find": .string("See"),
                "replace": .string("Read")
            ]
        )

        XCTAssertNil(replaceResult.isError)
        XCTAssertTrue(try readDocumentText(from: url).contains("Read"))
    }
}
