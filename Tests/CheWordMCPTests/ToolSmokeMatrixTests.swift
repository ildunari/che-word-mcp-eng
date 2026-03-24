import XCTest
import MCP
import OOXMLSwift
@testable import CheWordMCP

final class ToolSmokeMatrixTests: XCTestCase {
    private enum CoverageBucket: String {
        case lifecycle
        case documentContent
        case paragraphFormatting
        case styleManagement
        case tables
        case images
        case linksCommentsRevisions
        case pageSectionLayout
        case fieldsReferences
        case protectionLanguageSearchExportCompare
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

    private func tempURL(_ suffix: String = UUID().uuidString, ext: String = "docx") -> URL {
        FileManager.default.temporaryDirectory.appendingPathComponent("cheword-smoke-\(suffix).\(ext)")
    }

    private func bucket(for toolName: String) -> CoverageBucket? {
        switch toolName {
        case "create_document", "open_document", "save_document", "close_document", "finalize_document",
             "list_open_documents", "get_document_session_state", "get_document_info":
            return .lifecycle
        case "get_text", "get_paragraphs", "insert_paragraph", "update_paragraph", "delete_paragraph",
             "replace_text", "insert_text", "get_document_text", "insert_bullet_list", "insert_numbered_list",
             "set_list_level":
            return .documentContent
        case "format_text", "format_text_range", "replace_text_range", "set_paragraph_format", "apply_style",
             "get_paragraph_runs", "get_text_with_formatting", "search_by_formatting", "search_text",
             "search_text_with_formatting", "list_all_formatted_text", "get_word_count_by_section":
            return .paragraphFormatting
        case "list_styles", "create_style", "update_style", "delete_style":
            return .styleManagement
        case "insert_table", "get_tables", "update_cell", "delete_table", "merge_cells", "set_table_style",
             "add_column_to_table", "add_row_to_table", "delete_column_from_table", "delete_row_from_table",
             "set_header_row", "set_table_alignment", "set_row_height", "set_cell_width", "set_cell_vertical_alignment":
            return .tables
        case "insert_image", "insert_image_from_path", "insert_floating_image", "update_image", "delete_image",
             "list_images", "export_image", "export_all_images", "set_image_style", "insert_image_watermark":
            return .images
        case "insert_hyperlink", "insert_internal_link", "update_hyperlink", "delete_hyperlink", "list_hyperlinks",
             "insert_bookmark", "delete_bookmark", "list_bookmarks", "insert_comment", "update_comment",
             "delete_comment", "list_comments", "reply_to_comment", "resolve_comment", "enable_track_changes",
             "disable_track_changes", "get_revisions", "accept_revision", "reject_revision",
             "accept_all_revisions", "reject_all_revisions", "insert_footnote", "delete_footnote",
             "list_footnotes", "insert_endnote", "delete_endnote", "list_endnotes":
            return .linksCommentsRevisions
        case "set_page_size", "set_page_margins", "set_page_orientation", "insert_page_break", "insert_section_break",
             "insert_continuous_section_break", "add_header", "update_header", "add_footer", "update_footer",
             "insert_page_number", "set_columns", "insert_column_break", "set_line_numbers", "set_page_borders",
             "set_text_direction", "insert_drop_cap", "insert_horizontal_line", "set_paragraph_border",
             "set_paragraph_shading", "set_character_spacing", "set_text_effect", "set_keep_lines",
             "set_keep_with_next", "set_outline_level", "set_page_break_before", "set_widow_orphan",
             "get_section_properties":
            return .pageSectionLayout
        case "insert_toc", "insert_text_field", "insert_checkbox", "insert_dropdown", "insert_equation",
             "insert_if_field", "insert_calculation_field", "insert_date_field", "insert_page_field",
             "insert_merge_field", "insert_sequence_field", "insert_content_control", "insert_repeating_section",
             "insert_caption", "insert_cross_reference", "insert_index", "insert_index_entry", "insert_table_of_figures",
             "insert_tab_stop", "clear_tab_stops", "insert_symbol", "insert_watermark":
            return .fieldsReferences
        case "set_document_properties", "get_document_properties", "compare_documents", "export_text", "export_markdown",
             "set_language", "protect_document", "unprotect_document",
             "set_document_password", "remove_document_password", "restrict_editing_region", "remove_watermark":
            return .protectionLanguageSearchExportCompare
        default:
            return nil
        }
    }

    func testEveryToolIsAssignedCoverageBucket() async {
        let server = await WordMCPServer()
        let uncategorized = server.toolsForTesting().map(\.name).filter { bucket(for: $0) == nil }
        XCTAssertEqual(uncategorized, [], "Uncategorized tools: \(uncategorized)")
    }

    func testLifecycleSmokeWorkflow() async throws {
        let server = await WordMCPServer()
        let docURL = tempURL("lifecycle")

        let createResult = await server.invokeToolForTesting(
            name: "create_document",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertNil(createResult.isError)

        let insertResult = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("doc"), "text": .string("Smoke test")]
        )
        XCTAssertNil(insertResult.isError)

        let saveResult = await server.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc"), "path": .string(docURL.path)]
        )
        XCTAssertNil(saveResult.isError)

        let listResult = await server.invokeToolForTesting(name: "list_open_documents")
        XCTAssertTrue(resultText(listResult).contains("doc"))

        let finalizeResult = await server.invokeToolForTesting(
            name: "finalize_document",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertNil(finalizeResult.isError)
    }

    func testParagraphAndInlineEditingSmokeWorkflow() async throws {
        let server = await WordMCPServer()
        let docURL = tempURL("inline")

        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("doc")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("doc"), "text": .string("Hello world")]
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
        XCTAssertNil(replaceResult.isError)

        let formatResult = await server.invokeToolForTesting(
            name: "format_text_range",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "start": .int(6),
                "end": .int(10),
                "bold": .bool(true)
            ]
        )
        XCTAssertNil(formatResult.isError)

        _ = await server.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("doc"), "path": .string(docURL.path)]
        )

        let saved = try DocxReader.read(from: docURL)
        XCTAssertEqual(saved.getParagraphs()[0].getText(), "Hello team")
        XCTAssertEqual(saved.getParagraphs()[0].runs.count, 2)
    }

    func testTableFamilySmokeWorkflow() async {
        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("doc")])

        let insertResult = await server.invokeToolForTesting(
            name: "insert_table",
            arguments: [
                "doc_id": .string("doc"),
                "rows": .int(2),
                "cols": .int(2),
                "data": .array([.string("A1"), .string("B1"), .string("A2"), .string("B2")])
            ]
        )
        XCTAssertNil(insertResult.isError)

        let updateResult = await server.invokeToolForTesting(
            name: "update_cell",
            arguments: [
                "doc_id": .string("doc"),
                "table_index": .int(0),
                "row": .int(1),
                "col": .int(1),
                "text": .string("Updated")
            ]
        )
        XCTAssertNil(updateResult.isError)

        let listResult = await server.invokeToolForTesting(
            name: "get_tables",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertTrue(resultText(listResult).contains("Table"))
    }

    func testCommentsAndRevisionsSmokeWorkflow() async throws {
        let server = await WordMCPServer()
        let docURL = tempURL("comments-revisions")

        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("doc")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("doc"), "text": .string("Draft paragraph")]
        )

        let commentResult = await server.invokeToolForTesting(
            name: "insert_comment",
            arguments: [
                "doc_id": .string("doc"),
                "paragraph_index": .int(0),
                "author": .string("Reviewer"),
                "text": .string("Please revise")
            ]
        )
        XCTAssertNil(commentResult.isError)

        let replyResult = await server.invokeToolForTesting(
            name: "reply_to_comment",
            arguments: [
                "doc_id": .string("doc"),
                "parent_comment_id": .int(1),
                "text": .string("Updated"),
                "author": .string("Author")
            ]
        )
        XCTAssertNil(replyResult.isError)

        let updateResult = await server.invokeToolForTesting(
            name: "update_paragraph",
            arguments: [
                "doc_id": .string("doc"),
                "index": .int(0),
                "text": .string("Final paragraph")
            ]
        )
        XCTAssertNil(updateResult.isError)

        let revisionsResult = await server.invokeToolForTesting(
            name: "get_revisions",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertTrue(resultText(revisionsResult).contains("Revisions in document"))

        let acceptAllResult = await server.invokeToolForTesting(
            name: "accept_all_revisions",
            arguments: ["doc_id": .string("doc")]
        )
        XCTAssertNil(acceptAllResult.isError)

        let saveResult = await server.invokeToolForTesting(
            name: "save_document",
            arguments: [
                "doc_id": .string("doc"),
                "path": .string(docURL.path)
            ]
        )
        XCTAssertNil(saveResult.isError)
        XCTAssertTrue(try DocxReader.read(from: docURL).getText().contains("Final paragraph"))
    }

    func testExportAndCompareSmokeWorkflow() async throws {
        let server = await WordMCPServer()
        let urlA = tempURL("compare-a")
        let urlB = tempURL("compare-b")
        let txtURL = tempURL("compare", ext: "txt")

        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("a")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("a"), "text": .string("Alpha")]
        )
        _ = await server.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("a"), "path": .string(urlA.path)]
        )

        _ = await server.invokeToolForTesting(name: "create_document", arguments: ["doc_id": .string("b")])
        _ = await server.invokeToolForTesting(
            name: "insert_paragraph",
            arguments: ["doc_id": .string("b"), "text": .string("Beta")]
        )
        _ = await server.invokeToolForTesting(
            name: "save_document",
            arguments: ["doc_id": .string("b"), "path": .string(urlB.path)]
        )

        let compareResult = await server.invokeToolForTesting(
            name: "compare_documents",
            arguments: [
                "doc_id_a": .string("a"),
                "doc_id_b": .string("b"),
                "mode": .string("text")
            ]
        )
        XCTAssertNil(compareResult.isError)

        let exportResult = await server.invokeToolForTesting(
            name: "export_text",
            arguments: [
                "doc_id": .string("b"),
                "path": .string(txtURL.path)
            ]
        )
        XCTAssertNil(exportResult.isError)
        XCTAssertTrue(FileManager.default.fileExists(atPath: txtURL.path))
    }
}
