import XCTest
@testable import OOXMLSwift

/// Targeted tests for DocxReader parsing logic.
/// Uses Write→Read pipeline to exercise specific XML parsing paths.
final class DocxReaderXMLTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxReaderXMLTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    private func roundTrip(_ document: WordDocument) throws -> WordDocument {
        let docxURL = tempDir.appendingPathComponent("test-\(UUID().uuidString).docx")
        try DocxWriter.write(document, to: docxURL)
        return try DocxReader.read(from: docxURL)
    }

    // MARK: - Empty / Minimal Documents

    func testReadEmptyDocument() throws {
        let doc = WordDocument()
        let result = try roundTrip(doc)
        // Empty body should still be valid
        XCTAssertTrue(result.body.children.isEmpty)
    }

    func testReadFileNotFound() throws {
        let fakeURL = tempDir.appendingPathComponent("nonexistent.docx")
        XCTAssertThrowsError(try DocxReader.read(from: fakeURL))
    }

    // MARK: - Styles Parsing

    func testReadDefaultStyles() throws {
        let doc = WordDocument()
        let result = try roundTrip(doc)
        // Default styles should be present
        XCTAssertFalse(result.styles.isEmpty)
    }

    func testReadCustomStyles() throws {
        var doc = WordDocument()
        doc.styles = [
            Style(id: "Heading1", name: "heading 1", type: .paragraph),
            Style(id: "Normal", name: "Normal", type: .paragraph)
        ]
        doc.body.children = [.paragraph(Paragraph(text: "Text"))]

        let result = try roundTrip(doc)
        let styleIds = result.styles.map { $0.id }
        XCTAssertTrue(styleIds.contains("Heading1"))
        XCTAssertTrue(styleIds.contains("Normal"))
    }

    func testReadStyleBasedOn() throws {
        var doc = WordDocument()
        var style = Style(id: "CustomHeading", name: "Custom Heading", type: .paragraph)
        style.basedOn = "Heading1"
        doc.styles = [
            Style(id: "Heading1", name: "heading 1", type: .paragraph),
            style
        ]
        doc.body.children = [.paragraph(Paragraph(text: "Text"))]

        let result = try roundTrip(doc)
        let customStyle = result.styles.first { $0.id == "CustomHeading" }
        XCTAssertEqual(customStyle?.basedOn, "Heading1")
    }

    func testReadStyleIsDefault() throws {
        var doc = WordDocument()
        var style = Style(id: "Normal", name: "Normal", type: .paragraph)
        style.isDefault = true
        doc.styles = [style]
        doc.body.children = [.paragraph(Paragraph(text: "Text"))]

        let result = try roundTrip(doc)
        let normalStyle = result.styles.first { $0.id == "Normal" }
        XCTAssertEqual(normalStyle?.isDefault, true)
    }

    func testReadStyleWithParagraphProperties() throws {
        var doc = WordDocument()
        var style = Style(id: "CenteredStyle", name: "Centered", type: .paragraph)
        var pProps = ParagraphProperties()
        pProps.alignment = .center
        style.paragraphProperties = pProps
        doc.styles = [style]
        doc.body.children = [.paragraph(Paragraph(text: "Text"))]

        let result = try roundTrip(doc)
        let readStyle = result.styles.first { $0.id == "CenteredStyle" }
        XCTAssertEqual(readStyle?.paragraphProperties?.alignment, .center)
    }

    func testReadStyleWithRunProperties() throws {
        var doc = WordDocument()
        var style = Style(id: "BoldStyle", name: "Bold", type: .character)
        var rProps = RunProperties()
        rProps.bold = true
        rProps.fontSize = 28
        style.runProperties = rProps
        doc.styles = [style]
        doc.body.children = [.paragraph(Paragraph(text: "Text"))]

        let result = try roundTrip(doc)
        let readStyle = result.styles.first { $0.id == "BoldStyle" }
        XCTAssertEqual(readStyle?.runProperties?.bold, true)
        XCTAssertEqual(readStyle?.runProperties?.fontSize, 28)
    }

    // MARK: - Semantic Detection

    func testReadHeadingSemantic() throws {
        var doc = WordDocument()
        doc.styles = Style.defaultStyles
        var props = ParagraphProperties()
        props.style = "Heading1"
        doc.body.children = [.paragraph(Paragraph(text: "Title", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.semantic, SemanticAnnotation.heading(1))
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadHeading2Semantic() throws {
        var doc = WordDocument()
        doc.styles = Style.defaultStyles
        var props = ParagraphProperties()
        props.style = "Heading2"
        doc.body.children = [.paragraph(Paragraph(text: "Subtitle", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.semantic, SemanticAnnotation.heading(2))
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadParagraphSemantic() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Normal text"))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.semantic, SemanticAnnotation.paragraph)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadNestedBulletSemantic() throws {
        var doc = WordDocument()
        let numId = doc.numbering.createBulletList()
        var propsL0 = ParagraphProperties()
        propsL0.numbering = NumberingInfo(numId: numId, level: 0)
        var propsL1 = ParagraphProperties()
        propsL1.numbering = NumberingInfo(numId: numId, level: 1)
        doc.body.children = [
            .paragraph(Paragraph(text: "L0", properties: propsL0)),
            .paragraph(Paragraph(text: "L1", properties: propsL1))
        ]

        let result = try roundTrip(doc)

        if case .paragraph(let p0) = result.body.children[0] {
            XCTAssertEqual(p0.semantic, SemanticAnnotation.bulletItem(level: 0))
        } else { XCTFail("Expected paragraph at [0]") }

        if case .paragraph(let p1) = result.body.children[1] {
            XCTAssertEqual(p1.semantic, SemanticAnnotation.bulletItem(level: 1))
        } else { XCTFail("Expected paragraph at [1]") }
    }

    // MARK: - Numbering Parsing

    func testReadNumberingAbstractNums() throws {
        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        _ = doc.numbering.createNumberedList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.numbering.abstractNums.count, 2)
    }

    func testReadNumberingLevels() throws {
        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let result = try roundTrip(doc)
        let abstractNum = result.numbering.abstractNums.first
        XCTAssertNotNil(abstractNum)
        // Bullet lists typically have multiple levels
        XCTAssertGreaterThan(abstractNum?.levels.count ?? 0, 0)
    }

    func testReadNumberingBulletFormat() throws {
        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let result = try roundTrip(doc)
        let firstLevel = result.numbering.abstractNums.first?.levels.first
        XCTAssertEqual(firstLevel?.numFmt, .bullet)
    }

    func testReadNumberingDecimalFormat() throws {
        var doc = WordDocument()
        _ = doc.numbering.createNumberedList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let result = try roundTrip(doc)
        // Find the numbered list (second abstract num)
        let numberedAbstract = result.numbering.abstractNums.first { abstractNum in
            abstractNum.levels.first?.numFmt == .decimal
        }
        XCTAssertNotNil(numberedAbstract)
    }

    func testReadNumberingNums() throws {
        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let result = try roundTrip(doc)
        XCTAssertGreaterThan(result.numbering.nums.count, 0)
        // Each Num should reference a valid abstractNumId
        for num in result.numbering.nums {
            let hasAbstract = result.numbering.abstractNums.contains { $0.abstractNumId == num.abstractNumId }
            XCTAssertTrue(hasAbstract, "Num \(num.numId) references missing abstractNumId \(num.abstractNumId)")
        }
    }

    func testReadNumberingLevelIndent() throws {
        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let result = try roundTrip(doc)
        let levels = result.numbering.abstractNums.first?.levels ?? []
        // Each successive level should have increasing indent
        if levels.count >= 2 {
            XCTAssertGreaterThan(levels[1].indent, levels[0].indent)
        }
    }

    // MARK: - Table Cell Properties

    func testReadTableCellWidth() throws {
        var doc = WordDocument()
        var cellProps = TableCellProperties()
        cellProps.width = 4500
        cellProps.widthType = .dxa
        var cell = TableCell(text: "Cell")
        cell.properties = cellProps
        let table = Table(rows: [TableRow(cells: [cell])])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows[0].cells[0].properties.width, 4500)
            XCTAssertEqual(readTable.rows[0].cells[0].properties.widthType, .dxa)
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableCellShading() throws {
        var doc = WordDocument()
        var cellProps = TableCellProperties()
        cellProps.shading = CellShading(fill: "FFFF00")
        var cell = TableCell(text: "Shaded")
        cell.properties = cellProps
        let table = Table(rows: [TableRow(cells: [cell])])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows[0].cells[0].properties.shading?.fill, "FFFF00")
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableCellVerticalAlignment() throws {
        var doc = WordDocument()
        var cellProps = TableCellProperties()
        cellProps.verticalAlignment = .center
        var cell = TableCell(text: "Center")
        cell.properties = cellProps
        let table = Table(rows: [TableRow(cells: [cell])])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows[0].cells[0].properties.verticalAlignment, .center)
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableCellGridSpan() throws {
        var doc = WordDocument()
        var cellProps = TableCellProperties()
        cellProps.gridSpan = 2
        var cell = TableCell(text: "Merged")
        cell.properties = cellProps
        let table = Table(rows: [
            TableRow(cells: [cell, TableCell(text: "Normal")])
        ])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows[0].cells[0].properties.gridSpan, 2)
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableCantSplit() throws {
        var doc = WordDocument()
        var rowProps = TableRowProperties()
        rowProps.cantSplit = true
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "NoSplit")], properties: rowProps)
        ])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertTrue(readTable.rows[0].properties.cantSplit)
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadMultipleTables() throws {
        var doc = WordDocument()
        let table1 = Table(rows: [TableRow(cells: [TableCell(text: "T1")])])
        let table2 = Table(rows: [TableRow(cells: [TableCell(text: "T2")])])
        doc.body.children = [
            .table(table1),
            .paragraph(Paragraph(text: "Between")),
            .table(table2)
        ]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.children.count, 3)

        if case .table(let t1) = result.body.children[0] {
            XCTAssertEqual(t1.rows[0].cells[0].paragraphs[0].getText(), "T1")
        } else { XCTFail("Expected table at [0]") }

        if case .table(let t2) = result.body.children[2] {
            XCTAssertEqual(t2.rows[0].cells[0].paragraphs[0].getText(), "T2")
        } else { XCTFail("Expected table at [2]") }
    }

    // MARK: - Table in body.tables

    func testReadBodyTablesPopulated() throws {
        var doc = WordDocument()
        let table = Table(rows: [TableRow(cells: [TableCell(text: "Data")])])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.tables.count, 1)
    }

    // MARK: - Core Properties (Detailed)

    func testReadCorePropertiesDescription() throws {
        var doc = WordDocument()
        doc.properties.description = "A test document"
        doc.body.children = [.paragraph(Paragraph(text: "Content"))]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.properties.description, "A test document")
    }

    func testReadCorePropertiesLastModifiedBy() throws {
        var doc = WordDocument()
        doc.properties.lastModifiedBy = "Editor"
        doc.body.children = [.paragraph(Paragraph(text: "Content"))]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.properties.lastModifiedBy, "Editor")
    }

    func testReadCorePropertiesCreatedDate() throws {
        var doc = WordDocument()
        let formatter = ISO8601DateFormatter()
        doc.properties.created = formatter.date(from: "2024-01-15T10:30:00Z")
        doc.body.children = [.paragraph(Paragraph(text: "Content"))]

        let result = try roundTrip(doc)
        XCTAssertNotNil(result.properties.created)
        // Compare with second precision (ISO8601 round-trip)
        if let created = result.properties.created {
            let expected = formatter.date(from: "2024-01-15T10:30:00Z")!
            XCTAssertEqual(created.timeIntervalSince1970, expected.timeIntervalSince1970, accuracy: 1.0)
        }
    }

    func testReadCorePropertiesModifiedDate() throws {
        var doc = WordDocument()
        let formatter = ISO8601DateFormatter()
        doc.properties.modified = formatter.date(from: "2024-06-20T15:45:00Z")
        doc.body.children = [.paragraph(Paragraph(text: "Content"))]

        let result = try roundTrip(doc)
        XCTAssertNotNil(result.properties.modified)
        if let modified = result.properties.modified {
            let expected = formatter.date(from: "2024-06-20T15:45:00Z")!
            XCTAssertEqual(modified.timeIntervalSince1970, expected.timeIntervalSince1970, accuracy: 1.0)
        }
    }

    // MARK: - Comment anchoring

    func testReadCommentParagraphIndex() throws {
        var doc = WordDocument()
        let comment = Comment(id: 1, author: "Bob", text: "Note", paragraphIndex: 1)
        doc.comments.addComment(comment)
        var para0 = Paragraph(text: "First")
        var para1 = Paragraph(text: "Second")
        para1.commentIds = [1]
        doc.body.children = [.paragraph(para0), .paragraph(para1)]

        let result = try roundTrip(doc)
        let readComment = result.comments.comments.first { $0.id == 1 }
        XCTAssertEqual(readComment?.paragraphIndex, 1)
    }

    // MARK: - Paragraph Style Reference

    func testReadParagraphStyleReference() throws {
        var doc = WordDocument()
        doc.styles = Style.defaultStyles
        var props = ParagraphProperties()
        props.style = "Heading1"
        doc.body.children = [.paragraph(Paragraph(text: "H1", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.properties.style, "Heading1")
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - Spacing LineRule

    func testReadSpacingLineRule() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.spacing = Spacing(line: 360, lineRule: .auto)
        doc.body.children = [.paragraph(Paragraph(text: "Text", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.properties.spacing?.lineRule, .auto)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - Indentation Hanging

    func testReadIndentationHanging() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.indentation = Indentation(left: 720, hanging: 360)
        doc.body.children = [.paragraph(Paragraph(text: "Hanging", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.properties.indentation?.left, 720)
            XCTAssertEqual(para.properties.indentation?.hanging, 360)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - XML Escaping Round-trip

    func testReadXMLSpecialCharacters() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Brackets <angle> & ampersands \"quotes\""))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.getText(), "Brackets <angle> & ampersands \"quotes\"")
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadCorePropertiesXMLEscaping() throws {
        var doc = WordDocument()
        doc.properties.title = "Title with <angle> & \"special\" chars"
        doc.body.children = [.paragraph(Paragraph(text: "Content"))]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.properties.title, "Title with <angle> & \"special\" chars")
    }

    // MARK: - Table with Multi-Paragraph Cells

    func testReadTableCellMultipleParagraphs() throws {
        var doc = WordDocument()
        var cell = TableCell()
        cell.paragraphs = [
            Paragraph(text: "Line 1"),
            Paragraph(text: "Line 2")
        ]
        let table = Table(rows: [TableRow(cells: [cell])])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows[0].cells[0].paragraphs.count, 2)
            XCTAssertEqual(readTable.rows[0].cells[0].paragraphs[0].getText(), "Line 1")
            XCTAssertEqual(readTable.rows[0].cells[0].paragraphs[1].getText(), "Line 2")
        } else {
            XCTFail("Expected table")
        }
    }

    // MARK: - Combined Formatting

    func testReadCombinedRunFormatting() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.bold = true
        props.italic = true
        props.underline = .single
        props.fontName = "Times New Roman"
        props.fontSize = 32
        props.color = "0000FF"
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Fully styled", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            let rp = para.runs[0].properties
            XCTAssertTrue(rp.bold)
            XCTAssertTrue(rp.italic)
            XCTAssertEqual(rp.underline, .single)
            XCTAssertEqual(rp.fontName, "Times New Roman")
            XCTAssertEqual(rp.fontSize, 32)
            XCTAssertEqual(rp.color, "0000FF")
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - Large Document

    func testReadLargeDocument() throws {
        var doc = WordDocument()
        for i in 0..<50 {
            doc.body.children.append(.paragraph(Paragraph(text: "Paragraph \(i)")))
        }

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.children.count, 50)

        // Spot check first and last
        if case .paragraph(let first) = result.body.children[0] {
            XCTAssertEqual(first.getText(), "Paragraph 0")
        }
        if case .paragraph(let last) = result.body.children[49] {
            XCTAssertEqual(last.getText(), "Paragraph 49")
        }
    }
}
