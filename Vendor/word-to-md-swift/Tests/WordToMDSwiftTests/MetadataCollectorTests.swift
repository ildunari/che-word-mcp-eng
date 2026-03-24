import XCTest
import OOXMLSwift
@testable import WordToMDSwift

final class MetadataCollectorTests: XCTestCase {

    // MARK: - Helpers

    private func makeYAML(from collector: MetadataCollector) throws -> String {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("meta-test-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        let url = tempDir.appendingPathComponent("test.meta.yaml")
        try collector.writeYAML(to: url)
        return try String(contentsOf: url, encoding: .utf8)
    }

    // MARK: - Bug Fix: characterSpacing

    func testCharacterSpacingIsWrittenToYAML() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(spacing: 20)
        let run = Run(text: "spaced", properties: props)
        doc.appendParagraph(Paragraph(runs: [run]))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("characterSpacing"), "characterSpacing should appear in YAML output")
        XCTAssertTrue(yaml.contains("spacing: 20"))
    }

    func testCharacterSpacingWithPositionAndKern() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(spacing: 10, position: 5, kern: 16)
        let run = Run(text: "text", properties: props)
        doc.appendParagraph(Paragraph(runs: [run]))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("spacing: 10"))
        XCTAssertTrue(yaml.contains("position: 5"))
        XCTAssertTrue(yaml.contains("kern: 16"))
    }

    // MARK: - Table Metadata

    func testTableMetadataCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var tableProps = TableProperties()
        tableProps.width = 9000
        tableProps.widthType = .dxa
        tableProps.alignment = .center
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "A")])
        ], properties: tableProps)
        doc.body.children = [.table(table)]
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("tables:"))
        XCTAssertTrue(yaml.contains("width: 9000"))
        XCTAssertTrue(yaml.contains("alignment: center"))
    }

    func testTableHeaderRowMetadata() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var headerRowProps = TableRowProperties()
        headerRowProps.isHeader = true
        var tableProps = TableProperties()
        tableProps.width = 9000
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "Header")], properties: headerRowProps),
            TableRow(cells: [TableCell(text: "Data")])
        ], properties: tableProps)
        doc.body.children = [.table(table)]
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("isHeader: true"))
    }

    // MARK: - Comment Content

    func testCommentContentCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        let comment = Comment(id: 1, author: "Alice", text: "Review this", paragraphIndex: 0)
        doc.comments.addComment(comment)
        doc.appendParagraph(Paragraph(text: "Text"))
        collector.collectDocument(doc)

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("comments:"))
        XCTAssertTrue(yaml.contains("author: \"Alice\""))
        XCTAssertTrue(yaml.contains("text: \"Review this\""))
    }

    func testCommentReplyCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        doc.comments.addComment(Comment(id: 1, author: "Alice", text: "Review", paragraphIndex: 0))
        _ = doc.comments.addReply(to: 1, author: "Bob", text: "Done")
        doc.appendParagraph(Paragraph(text: "Text"))
        collector.collectDocument(doc)

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("parentId: 1"))
    }

    func testCommentDoneStatus() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        doc.comments.addComment(Comment(id: 1, author: "Alice", text: "Check", paragraphIndex: 0))
        doc.comments.markAsDone(1)
        doc.appendParagraph(Paragraph(text: "Text"))
        collector.collectDocument(doc)

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("done: true"))
    }

    // MARK: - Numbering Definitions

    func testNumberingDefinitionCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        doc.appendParagraph(Paragraph(text: "Item"))
        collector.collectDocument(doc)

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("numbering:"))
        XCTAssertTrue(yaml.contains("abstractNumId: 0"))
        XCTAssertTrue(yaml.contains("numFmt: bullet"))
    }

    func testNumberedListDefinitionCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        _ = doc.numbering.createNumberedList()
        doc.appendParagraph(Paragraph(text: "Item"))
        collector.collectDocument(doc)

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("numFmt: decimal"))
        XCTAssertTrue(yaml.contains("numFmt: lowerLetter"))
    }

    // MARK: - Document Properties Enhancement

    func testKeywordsCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        doc.properties.keywords = "swift, docx, converter"
        doc.appendParagraph(Paragraph(text: "Text"))
        collector.collectDocument(doc)

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("keywords: \"swift, docx, converter\""))
    }

    // MARK: - Paragraph Advanced Properties

    func testKeepNextCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = ParagraphProperties()
        props.keepNext = true
        doc.appendParagraph(Paragraph(text: "Keep with next", properties: props))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("keepNext: true"))
    }

    func testKeepLinesCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = ParagraphProperties()
        props.keepLines = true
        doc.appendParagraph(Paragraph(text: "Keep lines", properties: props))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("keepLines: true"))
    }

    func testPageBreakBeforeCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = ParagraphProperties()
        props.pageBreakBefore = true
        doc.appendParagraph(Paragraph(text: "New page", properties: props))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("pageBreakBefore: true"))
    }

    func testParagraphBorderCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = ParagraphProperties()
        props.border = ParagraphBorder.all(ParagraphBorderStyle(type: .single, color: "FF0000", size: 8))
        doc.appendParagraph(Paragraph(text: "Bordered", properties: props))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("border:"))
        XCTAssertTrue(yaml.contains("type: single"))
        XCTAssertTrue(yaml.contains("color: \"FF0000\""))
    }

    func testParagraphShadingCollection() throws {
        var collector = MetadataCollector()

        var doc = WordDocument()
        var props = ParagraphProperties()
        props.shading = ParagraphShading(fill: "FFFF00")
        doc.appendParagraph(Paragraph(text: "Shaded", properties: props))
        collector.collectDocument(doc)

        for (index, child) in doc.body.children.enumerated() {
            collector.collectElement(child, index: index)
        }

        let yaml = try makeYAML(from: collector)
        XCTAssertTrue(yaml.contains("shading:"))
        XCTAssertTrue(yaml.contains("fill: \"FFFF00\""))
    }
}
