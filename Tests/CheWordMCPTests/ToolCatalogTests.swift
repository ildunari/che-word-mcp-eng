import XCTest
import MCP
@testable import CheWordMCP

final class ToolCatalogTests: XCTestCase {
    private let expectedToolCount = 149

    private let exactReadOnlyNames: Set<String> = [
        "compare_documents",
        "export_markdown",
        "export_text",
        "get_document_properties",
        "get_document_session_state",
        "get_paragraph_runs",
        "get_revisions",
        "get_section_properties",
        "get_tables",
        "get_text_with_formatting",
        "get_word_count_by_section",
    ]

    private func readOnlyToolNames(from tools: [Tool]) -> Set<String> {
        Set(tools.map(\.name).filter { name in
            name.hasPrefix("get_")
                || name.hasPrefix("list_")
                || name.hasPrefix("search_")
                || exactReadOnlyNames.contains(name)
        })
    }

    private func schemaHasShape(_ schema: Value) -> Bool {
        guard let object = schema.objectValue else {
            return false
        }

        guard !object.isEmpty else {
            return false
        }

        return object["type"] != nil || object["properties"] != nil || object["required"] != nil
    }

    func testToolNamesAreUnique() async {
        let server = await WordMCPServer()
        let names = server.toolsForTesting().map(\.name)

        XCTAssertEqual(Set(names).count, names.count)
    }

    func testToolCountMatchesExpectedCatalog() async {
        let server = await WordMCPServer()

        XCTAssertEqual(server.toolsForTesting().count, expectedToolCount)
    }

    func testAllToolsHaveDescriptionsAndSchemas() async {
        let server = await WordMCPServer()
        let tools = server.toolsForTesting()

        for tool in tools {
            XCTAssertNotNil(tool.description, "Missing description for \(tool.name)")
            XCTAssertFalse(tool.description?.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty ?? true, "Empty description for \(tool.name)")
            XCTAssertTrue(schemaHasShape(tool.inputSchema), "Schema missing object shape for \(tool.name)")
        }
    }

    func testAllToolsHaveTitles() async {
        let server = await WordMCPServer()
        let tools = server.toolsForTesting()

        for tool in tools {
            XCTAssertFalse(tool.annotations.title?.isEmpty ?? true, "Missing title for \(tool.name)")
        }
    }

    func testReadOnlyFamiliesAreAnnotatedReadOnly() async throws {
        let server = await WordMCPServer()
        let tools = server.toolsForTesting()
        let toolsByName = Dictionary(uniqueKeysWithValues: tools.map { ($0.name, $0) })

        for name in readOnlyToolNames(from: tools).sorted() {
            let tool = try XCTUnwrap(toolsByName[name], "Missing tool \(name)")
            XCTAssertEqual(tool.annotations.readOnlyHint, true, "Expected readOnlyHint for \(name)")
            XCTAssertNotEqual(tool.annotations.destructiveHint, true, "Read-only tool \(name) should not be marked destructive")
        }
    }
}
