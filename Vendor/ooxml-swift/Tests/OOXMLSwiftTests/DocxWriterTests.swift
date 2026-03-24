import XCTest
@testable import OOXMLSwift

final class DocxWriterTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxWriterTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    /// Helper: write a WordDocument to .docx, unzip, and return the extracted directory
    private func writeAndUnzip(_ document: WordDocument) throws -> URL {
        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(document, to: docxURL)

        let extractDir = tempDir.appendingPathComponent("extracted")
        try ZipHelper.unzip(docxURL, to: extractDir)
        return extractDir
    }

    /// Helper: read XML file content from extracted docx
    private func readXML(from extractDir: URL, path: String) throws -> String {
        let fileURL = extractDir.appendingPathComponent(path)
        return try String(contentsOf: fileURL, encoding: .utf8)
    }

    // MARK: - Basic Structure

    func testWriteCreatesValidZip() throws {
        let doc = WordDocument()
        let docxURL = tempDir.appendingPathComponent("basic.docx")
        try DocxWriter.write(doc, to: docxURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: docxURL.path))

        // Verify it's a valid ZIP by checking the magic bytes
        let data = try Data(contentsOf: docxURL)
        XCTAssertGreaterThan(data.count, 4)
        XCTAssertEqual(data[0], 0x50)  // 'P'
        XCTAssertEqual(data[1], 0x4B)  // 'K'
    }

    func testWriteContainsRequiredFiles() throws {
        let doc = WordDocument()
        let extractDir = try writeAndUnzip(doc)

        // Check required files exist
        let requiredFiles = [
            "[Content_Types].xml",
            "_rels/.rels",
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "word/styles.xml",
            "word/settings.xml",
            "word/fontTable.xml",
            "docProps/core.xml",
            "docProps/app.xml"
        ]

        for file in requiredFiles {
            let fileURL = extractDir.appendingPathComponent(file)
            XCTAssertTrue(FileManager.default.fileExists(atPath: fileURL.path), "Missing required file: \(file)")
        }
    }

    // MARK: - Document Content

    func testWriteParagraphContent() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Hello World"))]
        let extractDir = try writeAndUnzip(doc)

        let documentXML = try readXML(from: extractDir, path: "word/document.xml")
        XCTAssertTrue(documentXML.contains("Hello World"))
        XCTAssertTrue(documentXML.contains("<w:body>"))
    }

    func testWriteFormattedParagraph() throws {
        var doc = WordDocument()
        var props = ParagraphProperties()
        props.alignment = .center
        doc.body.children = [.paragraph(Paragraph(text: "Centered", properties: props))]
        let extractDir = try writeAndUnzip(doc)

        let documentXML = try readXML(from: extractDir, path: "word/document.xml")
        XCTAssertTrue(documentXML.contains("<w:jc w:val=\"center\"/>"))
    }

    func testWriteTable() throws {
        var doc = WordDocument()
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")]),
            TableRow(cells: [TableCell(text: "C"), TableCell(text: "D")])
        ])
        doc.body.children = [.table(table)]
        let extractDir = try writeAndUnzip(doc)

        let documentXML = try readXML(from: extractDir, path: "word/document.xml")
        XCTAssertTrue(documentXML.contains("<w:tbl>"))
        XCTAssertTrue(documentXML.contains("A"))
        XCTAssertTrue(documentXML.contains("D"))
    }

    // MARK: - Numbering

    func testWriteWithNumbering() throws {
        var doc = WordDocument()
        let numId = doc.numbering.createBulletList()
        var props = ParagraphProperties()
        props.numbering = NumberingInfo(numId: numId, level: 0)
        doc.body.children = [.paragraph(Paragraph(text: "Bullet item", properties: props))]

        let extractDir = try writeAndUnzip(doc)

        // numbering.xml should exist
        let numberingURL = extractDir.appendingPathComponent("word/numbering.xml")
        XCTAssertTrue(FileManager.default.fileExists(atPath: numberingURL.path))

        let numberingXML = try readXML(from: extractDir, path: "word/numbering.xml")
        XCTAssertTrue(numberingXML.contains("<w:numbering"))
        XCTAssertTrue(numberingXML.contains("<w:abstractNum"))
    }

    // MARK: - Comments

    func testWriteWithComments() throws {
        var doc = WordDocument()
        var comment = Comment(id: 1, author: "Test", text: "A comment", paragraphIndex: 0)
        comment.done = false
        doc.comments.addComment(comment)
        var para = Paragraph(text: "Commented text")
        para.commentIds = [1]
        doc.body.children = [.paragraph(para)]

        let extractDir = try writeAndUnzip(doc)

        let commentsURL = extractDir.appendingPathComponent("word/comments.xml")
        XCTAssertTrue(FileManager.default.fileExists(atPath: commentsURL.path))

        let commentsXML = try readXML(from: extractDir, path: "word/comments.xml")
        XCTAssertTrue(commentsXML.contains("A comment"))
        XCTAssertTrue(commentsXML.contains("w:author=\"Test\""))
    }

    // MARK: - Footnotes

    func testWriteWithFootnotes() throws {
        var doc = WordDocument()
        doc.footnotes.footnotes.append(Footnote(id: 1, text: "A footnote", paragraphIndex: 0))
        var para = Paragraph(text: "Text with footnote")
        para.footnoteIds = [1]
        doc.body.children = [.paragraph(para)]

        let extractDir = try writeAndUnzip(doc)

        let footnotesURL = extractDir.appendingPathComponent("word/footnotes.xml")
        XCTAssertTrue(FileManager.default.fileExists(atPath: footnotesURL.path))

        let footnotesXML = try readXML(from: extractDir, path: "word/footnotes.xml")
        XCTAssertTrue(footnotesXML.contains("A footnote"))
    }

    // MARK: - Core Properties

    func testWriteCoreProperties() throws {
        var doc = WordDocument()
        doc.properties.title = "Test Title"
        doc.properties.creator = "Test Author"
        doc.properties.subject = "Test Subject"

        let extractDir = try writeAndUnzip(doc)

        let coreXML = try readXML(from: extractDir, path: "docProps/core.xml")
        XCTAssertTrue(coreXML.contains("Test Title"))
        XCTAssertTrue(coreXML.contains("Test Author"))
        XCTAssertTrue(coreXML.contains("Test Subject"))
    }

    // MARK: - Content Types

    func testContentTypesIncludesNumbering() throws {
        var doc = WordDocument()
        _ = doc.numbering.createBulletList()
        doc.body.children = [.paragraph(Paragraph(text: "Item"))]

        let extractDir = try writeAndUnzip(doc)
        let contentTypes = try readXML(from: extractDir, path: "[Content_Types].xml")
        XCTAssertTrue(contentTypes.contains("numbering.xml"))
    }

    func testContentTypesIncludesComments() throws {
        var doc = WordDocument()
        doc.comments.addComment(Comment(id: 1, author: "A", text: "B", paragraphIndex: 0))
        doc.body.children = [.paragraph(Paragraph(text: "C"))]

        let extractDir = try writeAndUnzip(doc)
        let contentTypes = try readXML(from: extractDir, path: "[Content_Types].xml")
        XCTAssertTrue(contentTypes.contains("comments.xml"))
    }

    // MARK: - Styles

    func testWriteDefaultStyles() throws {
        let doc = WordDocument()
        let extractDir = try writeAndUnzip(doc)

        let stylesURL = extractDir.appendingPathComponent("word/styles.xml")
        XCTAssertTrue(FileManager.default.fileExists(atPath: stylesURL.path))

        let stylesXML = try readXML(from: extractDir, path: "word/styles.xml")
        XCTAssertTrue(stylesXML.contains("<w:styles"))
    }

    // MARK: - XML Escaping

    func testWriteXMLEscaping() throws {
        var doc = WordDocument()
        doc.properties.title = "Title with <special> & \"chars\""
        doc.body.children = [.paragraph(Paragraph(text: "Text with <angle> brackets & ampersands"))]

        let extractDir = try writeAndUnzip(doc)

        let coreXML = try readXML(from: extractDir, path: "docProps/core.xml")
        XCTAssertTrue(coreXML.contains("&lt;special&gt;"))
        XCTAssertTrue(coreXML.contains("&amp;"))

        let documentXML = try readXML(from: extractDir, path: "word/document.xml")
        XCTAssertTrue(documentXML.contains("&lt;angle&gt;"))
    }
}

// MARK: - ZipHelper extension for test use

extension ZipHelper {
    /// Unzip to a specific directory (for testing)
    static func unzip(_ zipURL: URL, to destURL: URL) throws {
        try FileManager.default.createDirectory(at: destURL, withIntermediateDirectories: true)
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/ditto")
        process.arguments = ["-xk", zipURL.path, destURL.path]
        try process.run()
        process.waitUntilExit()

        guard process.terminationStatus == 0 else {
            throw NSError(domain: "ZipHelper", code: Int(process.terminationStatus),
                         userInfo: [NSLocalizedDescriptionKey: "Failed to unzip"])
        }
    }
}
