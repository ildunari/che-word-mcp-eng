import Foundation
import XCTest
import MCP
import OOXMLSwift
@testable import CheWordMCP

final class ReportTemplateIntegrationTests: XCTestCase {
    private let templateURL = URL(fileURLWithPath: "/Users/Kosta/LocalDev/che-word-mcp-eng/Tests/Generic-report-template.docx")

    private func copiedTemplateURL(_ suffix: String = UUID().uuidString) throws -> URL {
        let destination = FileManager.default.temporaryDirectory.appendingPathComponent("generic-report-\(suffix).docx")
        if FileManager.default.fileExists(atPath: destination.path) {
            try FileManager.default.removeItem(at: destination)
        }
        try FileManager.default.copyItem(at: templateURL, to: destination)
        return destination
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

    private func parsedValue(label: String, in text: String) -> Int? {
        let pattern = "\(NSRegularExpression.escapedPattern(for: label)): ([0-9]+)"
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let range = NSRange(text.startIndex..., in: text)
        guard let match = regex.firstMatch(in: text, range: range),
              let valueRange = Range(match.range(at: 1), in: text) else {
            return nil
        }
        return Int(text[valueRange])
    }

    private func firstEditableMultiRunParagraph(in document: WordDocument) -> (index: Int, paragraph: Paragraph)? {
        for (index, paragraph) in document.getParagraphs().enumerated() {
            let editableRuns = paragraph.runs.filter { run in
                !run.text.isEmpty && run.rawXML == nil && run.properties.rawXML == nil && run.drawing == nil
            }
            if paragraph.hyperlinks.isEmpty && editableRuns.count >= 2 {
                return (index, paragraph)
            }
        }
        return nil
    }

    func testOpenLongReportFixtureAndInspectCounts() async throws {
        let workingCopy = try copiedTemplateURL("counts")
        let server = await WordMCPServer()

        let openResult = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(workingCopy.path),
                "doc_id": .string("report")
            ]
        )
        XCTAssertNil(openResult.isError)

        let infoResult = await server.invokeToolForTesting(
            name: "get_document_info",
            arguments: ["doc_id": .string("report")]
        )
        XCTAssertNil(infoResult.isError)

        let text = resultText(infoResult)
        XCTAssertGreaterThan(parsedValue(label: "Paragraphs", in: text) ?? 0, 100)
        XCTAssertGreaterThan(parsedValue(label: "Tables", in: text) ?? 0, 0)
        XCTAssertGreaterThan(parsedValue(label: "Words", in: text) ?? 0, 100)
    }

    func testInspectionToolsWorkOnTemplate() async throws {
        let workingCopy = try copiedTemplateURL("inspect")
        let server = await WordMCPServer()

        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(workingCopy.path),
                "doc_id": .string("report")
            ]
        )

        let paragraphsResult = await server.invokeToolForTesting(
            name: "get_paragraphs",
            arguments: ["doc_id": .string("report")]
        )
        XCTAssertNil(paragraphsResult.isError)
        XCTAssertTrue(resultText(paragraphsResult).contains("Paragraphs:"))

        let paragraphRunsResult = await server.invokeToolForTesting(
            name: "get_paragraph_runs",
            arguments: [
                "doc_id": .string("report"),
                "paragraph_index": .int(0)
            ]
        )
        XCTAssertNil(paragraphRunsResult.isError)
        XCTAssertTrue(resultText(paragraphRunsResult).contains("Paragraph [0] Runs"))

        let searchResult = await server.invokeToolForTesting(
            name: "search_text_with_formatting",
            arguments: [
                "doc_id": .string("report"),
                "query": .string("the")
            ]
        )
        XCTAssertNil(searchResult.isError)

        for toolName in ["list_styles", "get_tables", "list_hyperlinks", "list_footnotes", "list_endnotes"] {
            let result = await server.invokeToolForTesting(
                name: toolName,
                arguments: ["doc_id": .string("report")]
            )
            XCTAssertNil(result.isError, "\(toolName) failed with \(resultText(result))")
        }
    }

    func testInlineRangeEditPreservesRunsOnTemplateCopy() async throws {
        let workingCopy = try copiedTemplateURL("inline")
        let original = try DocxReader.read(from: workingCopy)
        guard let target = firstEditableMultiRunParagraph(in: original) else {
            throw XCTSkip("Template did not contain a plain multi-run paragraph suitable for inline editing")
        }

        let editableRuns = target.paragraph.runs.filter { !$0.text.isEmpty }
        let firstRunLength = editableRuns[0].text.count
        let secondRunLength = editableRuns[1].text.count

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(workingCopy.path),
                "doc_id": .string("report"),
                "autosave": .bool(true)
            ]
        )

        let editResult = await server.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("report"),
                "paragraph_index": .int(target.index),
                "start": .int(firstRunLength),
                "end": .int(firstRunLength + secondRunLength),
                "replacement": .string("Edited")
            ]
        )
        XCTAssertNil(editResult.isError)

        let updated = try DocxReader.read(from: workingCopy)
        let updatedRuns = updated.getParagraphs()[target.index].runs
        XCTAssertEqual(updatedRuns.first?.text, editableRuns[0].text)
        XCTAssertEqual(updatedRuns[1].text, "Edited")
        XCTAssertGreaterThanOrEqual(updatedRuns.count, 2)
    }

    func testExportAndCompareWorkflowOnTemplateCopy() async throws {
        let baselineCopy = try copiedTemplateURL("baseline")
        let mutatedCopy = try copiedTemplateURL("mutated")
        let exportTextURL = FileManager.default.temporaryDirectory.appendingPathComponent("generic-report-export.txt")
        let exportMarkdownURL = FileManager.default.temporaryDirectory.appendingPathComponent("generic-report-export.md")

        let baselineDocument = try DocxReader.read(from: baselineCopy)
        guard let target = firstEditableMultiRunParagraph(in: baselineDocument) else {
            throw XCTSkip("Template did not contain a plain multi-run paragraph suitable for compare/export workflow")
        }

        let editableRuns = target.paragraph.runs.filter { !$0.text.isEmpty }
        let firstRunLength = editableRuns[0].text.count

        let server = await WordMCPServer()
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(baselineCopy.path),
                "doc_id": .string("baseline")
            ]
        )
        _ = await server.invokeToolForTesting(
            name: "open_document",
            arguments: [
                "path": .string(mutatedCopy.path),
                "doc_id": .string("mutated"),
                "autosave": .bool(true)
            ]
        )

        let mutateResult = await server.invokeToolForTesting(
            name: "replace_text_range",
            arguments: [
                "doc_id": .string("mutated"),
                "paragraph_index": .int(target.index),
                "start": .int(firstRunLength),
                "end": .int(firstRunLength + min(5, editableRuns[1].text.count)),
                "replacement": .string("Delta")
            ]
        )
        XCTAssertNil(mutateResult.isError)

        let compareResult = await server.invokeToolForTesting(
            name: "compare_documents",
            arguments: [
                "doc_id_a": .string("baseline"),
                "doc_id_b": .string("mutated"),
                "mode": .string("text")
            ]
        )
        XCTAssertNil(compareResult.isError)
        XCTAssertTrue(resultText(compareResult).contains("baseline"))

        let exportTextResult = await server.invokeToolForTesting(
            name: "export_text",
            arguments: [
                "doc_id": .string("mutated"),
                "path": .string(exportTextURL.path)
            ]
        )
        XCTAssertNil(exportTextResult.isError)
        XCTAssertTrue(FileManager.default.fileExists(atPath: exportTextURL.path))

        let exportMarkdownResult = await server.invokeToolForTesting(
            name: "export_markdown",
            arguments: [
                "source_path": .string(mutatedCopy.path),
                "path": .string(exportMarkdownURL.path)
            ]
        )
        XCTAssertNil(exportMarkdownResult.isError)
        XCTAssertTrue(FileManager.default.fileExists(atPath: exportMarkdownURL.path))
    }
}
