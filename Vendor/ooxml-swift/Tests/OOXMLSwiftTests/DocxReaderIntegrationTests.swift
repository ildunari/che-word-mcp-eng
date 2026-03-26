import XCTest
@testable import OOXMLSwift

/// Write→Read→Compare integration tests for DocxReader.
/// Each test creates a WordDocument, writes it to .docx via DocxWriter,
/// reads it back via DocxReader, and verifies the parsed output matches.
final class DocxReaderIntegrationTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxReaderIntTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    /// Helper: write → read round-trip
    private func roundTrip(_ document: WordDocument) throws -> WordDocument {
        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(document, to: docxURL)
        return try DocxReader.read(from: docxURL)
    }

    // MARK: - Basic Paragraph

    func testReadSimpleParagraph() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Hello World"))]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.children.count, 1)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs.count, 1)
            XCTAssertEqual(para.runs[0].text, "Hello World")
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadMultipleParagraphs() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "First")),
            .paragraph(Paragraph(text: "Second")),
            .paragraph(Paragraph(text: "Third"))
        ]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.children.count, 3)

        let texts = result.body.children.compactMap { child -> String? in
            if case .paragraph(let p) = child { return p.getText() }
            return nil
        }
        XCTAssertEqual(texts, ["First", "Second", "Third"])
    }

    func testReadMultipleRuns() throws {
        var doc = WordDocument()
        let runs = [Run(text: "Hello "), Run(text: "World")]
        doc.body.children = [.paragraph(Paragraph(runs: runs))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            // Runs might merge but text should be preserved
            let fullText = para.runs.map { $0.text }.joined()
            XCTAssertEqual(fullText, "Hello World")
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testGetParagraphsHidesTrackedDeletedParagraphShells() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "First"))
        doc.appendParagraph(Paragraph(text: "Second"))
        doc.enableTrackChanges(author: "Test")

        try doc.deleteParagraph(at: 0)

        XCTAssertEqual(doc.getParagraphs().map { $0.getText() }, ["Second"])
    }

    // MARK: - Paragraph Properties

    func testReadParagraphAlignment() throws {
        var doc = WordDocument()
        for alignment in [Alignment.center, .right, .both] {
            var props = ParagraphProperties()
            props.alignment = alignment
            doc.body.children.append(.paragraph(Paragraph(text: "\(alignment)", properties: props)))
        }

        let result = try roundTrip(doc)

        let alignments: [Alignment?] = result.body.children.compactMap { child in
            if case .paragraph(let p) = child { return p.properties.alignment }
            return nil
        }
        XCTAssertEqual(alignments, [.center, .right, .both])
    }

    func testReadParagraphSpacing() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.spacing = Spacing(before: 240, after: 120, line: 360)
        doc.body.children = [.paragraph(Paragraph(text: "Spaced", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.properties.spacing?.before, 240)
            XCTAssertEqual(para.properties.spacing?.after, 120)
            XCTAssertEqual(para.properties.spacing?.line, 360)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadParagraphIndentation() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.indentation = Indentation(left: 720, right: 360, firstLine: 360)
        doc.body.children = [.paragraph(Paragraph(text: "Indented", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.properties.indentation?.left, 720)
            XCTAssertEqual(para.properties.indentation?.right, 360)
            XCTAssertEqual(para.properties.indentation?.firstLine, 360)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadKeepNext() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.keepNext = true
        doc.body.children = [.paragraph(Paragraph(text: "Keep", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertTrue(para.properties.keepNext)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadKeepLines() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.keepLines = true
        doc.body.children = [.paragraph(Paragraph(text: "Lines", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertTrue(para.properties.keepLines)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadPageBreakBefore() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.pageBreakBefore = true
        doc.body.children = [.paragraph(Paragraph(text: "New page", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertTrue(para.properties.pageBreakBefore)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - Run Properties

    func testReadBoldItalic() throws {
        var doc = WordDocument()
        var boldProps = RunProperties()
        boldProps.bold = true
        var italicProps = RunProperties()
        italicProps.italic = true
        var bothProps = RunProperties()
        bothProps.bold = true
        bothProps.italic = true

        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Bold", properties: boldProps),
            Run(text: "Italic", properties: italicProps),
            Run(text: "Both", properties: bothProps)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs.count, 3)
            XCTAssertTrue(para.runs[0].properties.bold)
            XCTAssertFalse(para.runs[0].properties.italic)
            XCTAssertFalse(para.runs[1].properties.bold)
            XCTAssertTrue(para.runs[1].properties.italic)
            XCTAssertTrue(para.runs[2].properties.bold)
            XCTAssertTrue(para.runs[2].properties.italic)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadUnderline() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.underline = .single
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Underlined", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs[0].properties.underline, .single)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadStrikethrough() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.strikethrough = true
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Struck", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertTrue(para.runs[0].properties.strikethrough)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadSmallCapsAndAllCaps() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.smallCaps = true
        props.allCaps = true
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Caps", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertTrue(para.runs[0].properties.smallCaps)
            XCTAssertTrue(para.runs[0].properties.allCaps)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadFontAndSize() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.fontName = "Arial"
        props.fontSize = 24
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Styled", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs[0].properties.fontName, "Arial")
            XCTAssertEqual(para.runs[0].properties.fontSize, 24)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadColor() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.color = "FF0000"
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Red", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs[0].properties.color, "FF0000")
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadHighlight() throws {
        var doc = WordDocument()
        var props = RunProperties()
        props.highlight = .yellow
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "Highlighted", properties: props)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs[0].properties.highlight, .yellow)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadVerticalAlign() throws {
        var doc = WordDocument()
        var superProps = RunProperties()
        superProps.verticalAlign = .superscript
        var subProps = RunProperties()
        subProps.verticalAlign = .subscript
        doc.body.children = [.paragraph(Paragraph(runs: [
            Run(text: "super", properties: superProps),
            Run(text: "sub", properties: subProps)
        ]))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertEqual(para.runs[0].properties.verticalAlign, .superscript)
            XCTAssertEqual(para.runs[1].properties.verticalAlign, .subscript)
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - Table

    func testReadBasicTable() throws {
        var doc = WordDocument()
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")]),
            TableRow(cells: [TableCell(text: "C"), TableCell(text: "D")])
        ])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.children.count, 1)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows.count, 2)
            XCTAssertEqual(readTable.rows[0].cells.count, 2)
            XCTAssertEqual(readTable.rows[0].cells[0].paragraphs[0].getText(), "A")
            XCTAssertEqual(readTable.rows[0].cells[1].paragraphs[0].getText(), "B")
            XCTAssertEqual(readTable.rows[1].cells[0].paragraphs[0].getText(), "C")
            XCTAssertEqual(readTable.rows[1].cells[1].paragraphs[0].getText(), "D")
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableProperties() throws {
        var doc = WordDocument()
        var tableProps = TableProperties()
        tableProps.width = 9000
        tableProps.widthType = .dxa
        tableProps.alignment = .center
        tableProps.layout = .fixed
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "Cell")])
        ], properties: tableProps)
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.properties.width, 9000)
            XCTAssertEqual(readTable.properties.widthType, .dxa)
            XCTAssertEqual(readTable.properties.alignment, .center)
            XCTAssertEqual(readTable.properties.layout, .fixed)
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableRowHeader() throws {
        var doc = WordDocument()
        var headerRowProps = TableRowProperties()
        headerRowProps.isHeader = true
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "Header")], properties: headerRowProps),
            TableRow(cells: [TableCell(text: "Data")])
        ])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertTrue(readTable.rows[0].properties.isHeader)
            XCTAssertFalse(readTable.rows[1].properties.isHeader)
        } else {
            XCTFail("Expected table")
        }
    }

    func testReadTableRowHeight() throws {
        var doc = WordDocument()
        var rowProps = TableRowProperties()
        rowProps.height = 500
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "Tall")], properties: rowProps)
        ])
        doc.body.children = [.table(table)]

        let result = try roundTrip(doc)

        if case .table(let readTable) = result.body.children[0] {
            XCTAssertEqual(readTable.rows[0].properties.height, 500)
        } else {
            XCTFail("Expected table")
        }
    }

    // MARK: - Numbering

    func testReadBulletList() throws {
        var doc = WordDocument()
        let numId = doc.numbering.createBulletList()
        var props = ParagraphProperties()
        props.numbering = NumberingInfo(numId: numId, level: 0)
        doc.body.children = [.paragraph(Paragraph(text: "Bullet item", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertNotNil(para.properties.numbering)
            XCTAssertEqual(para.properties.numbering?.level, 0)
            // Verify semantic detection marks it as bullet
            XCTAssertEqual(para.semantic, SemanticAnnotation.bulletItem(level: 0))
        } else {
            XCTFail("Expected paragraph")
        }
    }

    func testReadOrderedList() throws {
        var doc = WordDocument()
        let numId = doc.numbering.createNumberedList()
        var props = ParagraphProperties()
        props.numbering = NumberingInfo(numId: numId, level: 0)
        doc.body.children = [.paragraph(Paragraph(text: "Item 1", properties: props))]

        let result = try roundTrip(doc)

        if case .paragraph(let para) = result.body.children[0] {
            XCTAssertNotNil(para.properties.numbering)
            // Verify semantic detection marks it as numbered
            XCTAssertEqual(para.semantic, SemanticAnnotation.numberedItem(level: 0))
        } else {
            XCTFail("Expected paragraph")
        }
    }

    // MARK: - Comments

    func testReadComments() throws {
        var doc = WordDocument()
        let comment = Comment(id: 1, author: "Alice", text: "Review this", paragraphIndex: 0)
        doc.comments.addComment(comment)
        var para = Paragraph(text: "Text with comment")
        para.commentIds = [1]
        doc.body.children = [.paragraph(para)]

        let result = try roundTrip(doc)

        XCTAssertEqual(result.comments.comments.count, 1)
        let readComment = result.comments.comments[0]
        XCTAssertEqual(readComment.id, 1)
        XCTAssertEqual(readComment.author, "Alice")
        XCTAssertTrue(readComment.text.contains("Review this"))
        XCTAssertEqual(readComment.paragraphIndex, 0)
    }

    // MARK: - Core Properties

    func testReadCoreProperties() throws {
        var doc = WordDocument()
        doc.properties.title = "Test Title"
        doc.properties.creator = "Test Author"
        doc.properties.subject = "Test Subject"
        doc.properties.keywords = "swift, test"
        doc.body.children = [.paragraph(Paragraph(text: "Content"))]

        let result = try roundTrip(doc)

        XCTAssertEqual(result.properties.title, "Test Title")
        XCTAssertEqual(result.properties.creator, "Test Author")
        XCTAssertEqual(result.properties.subject, "Test Subject")
        XCTAssertEqual(result.properties.keywords, "swift, test")
    }

    // MARK: - Mixed Content

    func testReadMixedParagraphsAndTables() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Before table")),
            .table(Table(rows: [
                TableRow(cells: [TableCell(text: "Cell")])
            ])),
            .paragraph(Paragraph(text: "After table"))
        ]

        let result = try roundTrip(doc)
        XCTAssertEqual(result.body.children.count, 3)

        if case .paragraph(let p1) = result.body.children[0] {
            XCTAssertEqual(p1.getText(), "Before table")
        } else { XCTFail("Expected paragraph at [0]") }

        if case .table(let t) = result.body.children[1] {
            XCTAssertEqual(t.rows.count, 1)
        } else { XCTFail("Expected table at [1]") }

        if case .paragraph(let p2) = result.body.children[2] {
            XCTAssertEqual(p2.getText(), "After table")
        } else { XCTFail("Expected paragraph at [2]") }
    }
}
