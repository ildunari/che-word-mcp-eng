import XCTest
import OOXMLSwift
import DocConverterSwift
@testable import WordToMDSwift

final class WordConverterTests: XCTestCase {
    let converter = WordConverter()

    // MARK: - Helpers

    /// 建立一個包含指定段落的 WordDocument
    private func makeDocument(paragraphs: [Paragraph]) -> WordDocument {
        var doc = WordDocument()
        for para in paragraphs {
            doc.appendParagraph(para)
        }
        return doc
    }

    /// 便利：單一段落文件
    private func makeDocument(paragraph: Paragraph) -> WordDocument {
        makeDocument(paragraphs: [paragraph])
    }

    /// 轉換為 Markdown 字串
    private func convert(
        _ doc: WordDocument,
        options: ConversionOptions = .default
    ) throws -> String {
        try converter.convertToString(document: doc, options: options)
    }

    // MARK: - 既有功能回歸測試

    func testBasicParagraph() throws {
        let doc = makeDocument(paragraph: Paragraph(text: "Hello world"))
        let md = try convert(doc)
        XCTAssertEqual(md.trimmingCharacters(in: .whitespacesAndNewlines), "Hello world")
    }

    func testHeading() throws {
        var para = Paragraph(text: "Title")
        para.properties.style = "Heading1"
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("# Title"))
    }

    func testHeadingLevel3() throws {
        var para = Paragraph(text: "Sub section")
        para.properties.style = "Heading 3"
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("### Sub section"))
    }

    func testBold() throws {
        let run = Run(text: "strong", properties: RunProperties(bold: true))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        let md = try convert(doc)
        XCTAssertTrue(md.contains("**strong**"))
    }

    func testItalic() throws {
        let run = Run(text: "emphasis", properties: RunProperties(italic: true))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        let md = try convert(doc)
        XCTAssertTrue(md.contains("_emphasis_"))
    }

    func testBoldItalic() throws {
        let run = Run(text: "both", properties: RunProperties(bold: true, italic: true))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        let md = try convert(doc)
        XCTAssertTrue(md.contains("***both***"))
    }

    func testStrikethrough() throws {
        let run = Run(text: "deleted", properties: RunProperties(strikethrough: true))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        let md = try convert(doc)
        XCTAssertTrue(md.contains("~~deleted~~"))
    }

    func testBulletList() throws {
        var para = Paragraph(text: "item one")
        para.properties.numbering = NumberingInfo(numId: 1, level: 0)

        var doc = WordDocument()
        // 設定 numbering 為 bullet
        var abstractNum = AbstractNum(abstractNumId: 0)
        abstractNum.levels = [Level(ilvl: 0, numFmt: .bullet, lvlText: "•", indent: 720)]
        doc.numbering.abstractNums = [abstractNum]
        doc.numbering.nums = [Num(numId: 1, abstractNumId: 0)]
        doc.appendParagraph(para)

        let md = try convert(doc)
        XCTAssertTrue(md.contains("- item one"))
    }

    func testNumberedList() throws {
        var para = Paragraph(text: "step one")
        para.properties.numbering = NumberingInfo(numId: 1, level: 0)

        var doc = WordDocument()
        var abstractNum = AbstractNum(abstractNumId: 0)
        abstractNum.levels = [Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720)]
        doc.numbering.abstractNums = [abstractNum]
        doc.numbering.nums = [Num(numId: 1, abstractNumId: 0)]
        doc.appendParagraph(para)

        let md = try convert(doc)
        XCTAssertTrue(md.contains("1. step one"))
    }

    // MARK: - Hyperlinks（tier_min = 1）

    func testHyperlinkExternal() throws {
        var para = Paragraph()
        para.runs = [Run(text: "before ")]
        para.hyperlinks = [
            Hyperlink(id: "h1", text: "Google", url: "https://google.com", relationshipId: "rId1")
        ]

        var doc = WordDocument()
        doc.hyperlinkReferences = [HyperlinkReference(relationshipId: "rId1", url: "https://google.com")]
        doc.appendParagraph(para)

        let md = try convert(doc)
        XCTAssertTrue(md.contains("[Google](https://google.com)"), "Got: \(md)")
    }

    func testHyperlinkInternal() throws {
        var para = Paragraph()
        para.runs = [Run(text: "see ")]
        para.hyperlinks = [
            Hyperlink(id: "h1", text: "section A", anchor: "sectionA")
        ]
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("[section A](#sectionA)"), "Got: \(md)")
    }

    // MARK: - Footnotes / Endnotes（tier_min = 1）

    func testFootnoteInParagraph() throws {
        var para = Paragraph(text: "Some text")
        para.footnoteIds = [1]

        var doc = WordDocument()
        doc.footnotes.footnotes = [Footnote(id: 1, text: "This is a footnote.", paragraphIndex: 0)]
        doc.appendParagraph(para)

        let md = try convert(doc)
        XCTAssertTrue(md.contains("[^1]"), "Missing footnote ref. Got: \(md)")
        XCTAssertTrue(md.contains("[^1]: This is a footnote."), "Missing footnote definition. Got: \(md)")
    }

    func testEndnoteMergedToFootnote() throws {
        var para = Paragraph(text: "Some text")
        para.endnoteIds = [1]

        var doc = WordDocument()
        doc.endnotes.endnotes = [Endnote(id: 1, text: "Endnote text.", paragraphIndex: 0)]
        doc.appendParagraph(para)

        let md = try convert(doc)
        // Endnotes 使用 "en" prefix 的 footnote 語法
        XCTAssertTrue(md.contains("[^en1]"), "Missing endnote ref. Got: \(md)")
        XCTAssertTrue(md.contains("[^en1]: Endnote text."), "Missing endnote definition. Got: \(md)")
    }

    // MARK: - Code（tier_min = 1）

    func testCodeBlock() throws {
        var para = Paragraph(text: "let x = 42")
        para.properties.style = "Code"
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("```"), "Missing code fence. Got: \(md)")
        XCTAssertTrue(md.contains("let x = 42"), "Missing code content. Got: \(md)")
    }

    func testCodeBlockSourceStyle() throws {
        var para = Paragraph(text: "print('hello')")
        para.properties.style = "Source Code"
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("```"), "Source style should trigger code block. Got: \(md)")
    }

    // MARK: - Blockquote（tier_min = 1）

    func testBlockquote() throws {
        var para = Paragraph(text: "Famous words")
        para.properties.style = "Quote"
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("> Famous words"), "Got: \(md)")
    }

    func testBlockquoteIntenseStyle() throws {
        var para = Paragraph(text: "Deep thought")
        para.properties.style = "Intense Quote"
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("> Deep thought"), "Got: \(md)")
    }

    // MARK: - Horizontal Rule（tier_min = 1）

    func testHorizontalRuleFromPageBreak() throws {
        var para = Paragraph()
        para.hasPageBreak = true
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("---"), "Page break should produce horizontal rule. Got: \(md)")
    }

    func testHorizontalRuleFromPageBreakBefore() throws {
        var para = Paragraph(text: "Next section")
        para.properties.pageBreakBefore = true
        let doc = makeDocument(paragraph: para)
        let md = try convert(doc)
        XCTAssertTrue(md.contains("---"), "pageBreakBefore should produce ---. Got: \(md)")
        XCTAssertTrue(md.contains("Next section"), "Text after page break should be preserved. Got: \(md)")
    }

    // MARK: - Images（reference tier_min = 1）

    func testImageInline() throws {
        let drawing = Drawing(
            type: .inline,
            width: 914400,
            height: 914400,
            imageId: "rId5",
            name: "photo",
            description: "A photo"
        )
        var run = Run(text: "")
        run.drawing = drawing

        var doc = WordDocument()
        doc.images = [ImageReference(id: "rId5", fileName: "image1.png", contentType: "image/png", data: Data())]
        doc.appendParagraph(Paragraph(runs: [run]))

        let md = try convert(doc)
        XCTAssertTrue(md.contains("![A photo](image1.png)"), "Got: \(md)")
    }

    func testImageWithoutDescription() throws {
        let drawing = Drawing(
            type: .inline,
            width: 914400,
            height: 914400,
            imageId: "rId5",
            name: "diagram",
            description: ""
        )
        var run = Run(text: "")
        run.drawing = drawing

        var doc = WordDocument()
        doc.images = [ImageReference(id: "rId5", fileName: "figure1.png", contentType: "image/png", data: Data())]
        doc.appendParagraph(Paragraph(runs: [run]))

        let md = try convert(doc)
        // 無 description 時使用 name 作為 alt
        XCTAssertTrue(md.contains("![diagram](figure1.png)"), "Got: \(md)")
    }

    // MARK: - Layer B HTML Extensions

    func testHTMLExtensionsDisabledByDefault() throws {
        let run = Run(text: "underlined", properties: RunProperties(underline: .single))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        let md = try convert(doc)
        // 預設不啟用 HTML extensions，underline 不應出現
        XCTAssertFalse(md.contains("<u>"), "HTML extensions should be disabled by default. Got: \(md)")
        XCTAssertTrue(md.contains("underlined"))
    }

    func testUnderlineWithHTMLExtensions() throws {
        let run = Run(text: "underlined", properties: RunProperties(underline: .single))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        var options = ConversionOptions.default
        options.useHTMLExtensions = true
        let md = try convert(doc, options: options)
        XCTAssertTrue(md.contains("<u>underlined</u>"), "Got: \(md)")
    }

    func testSuperscriptWithHTMLExtensions() throws {
        let run = Run(text: "2", properties: RunProperties(verticalAlign: .superscript))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        var options = ConversionOptions.default
        options.useHTMLExtensions = true
        let md = try convert(doc, options: options)
        XCTAssertTrue(md.contains("<sup>2</sup>"), "Got: \(md)")
    }

    func testSubscriptWithHTMLExtensions() throws {
        let run = Run(text: "2", properties: RunProperties(verticalAlign: .subscript))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        var options = ConversionOptions.default
        options.useHTMLExtensions = true
        let md = try convert(doc, options: options)
        XCTAssertTrue(md.contains("<sub>2</sub>"), "Got: \(md)")
    }

    func testHighlightWithHTMLExtensions() throws {
        let run = Run(text: "important", properties: RunProperties(highlight: .yellow))
        let doc = makeDocument(paragraph: Paragraph(runs: [run]))
        var options = ConversionOptions.default
        options.useHTMLExtensions = true
        let md = try convert(doc, options: options)
        XCTAssertTrue(md.contains("<mark>important</mark>"), "Got: \(md)")
    }

    // MARK: - Table

    func testBasicTable() throws {
        let table = Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [Paragraph(text: "Header 1")]),
                TableCell(paragraphs: [Paragraph(text: "Header 2")])
            ]),
            TableRow(cells: [
                TableCell(paragraphs: [Paragraph(text: "A")]),
                TableCell(paragraphs: [Paragraph(text: "B")])
            ])
        ])
        var doc = WordDocument()
        doc.body.children.append(.table(table))
        let md = try convert(doc)
        XCTAssertTrue(md.contains("| Header 1 | Header 2 |"), "Got: \(md)")
        XCTAssertTrue(md.contains("| A | B |"), "Got: \(md)")
    }

    // MARK: - Frontmatter

    func testFrontmatter() throws {
        var doc = WordDocument()
        doc.properties.title = "My Doc"
        doc.properties.creator = "Author"
        doc.properties.subject = "Test"
        doc.appendParagraph(Paragraph(text: "content"))

        var options = ConversionOptions.default
        options.includeFrontmatter = true
        let md = try convert(doc, options: options)
        XCTAssertTrue(md.contains("---"))
        XCTAssertTrue(md.contains("title: \"My Doc\""))
        XCTAssertTrue(md.contains("author: \"Author\""))
    }

    // MARK: - FidelityTier

    func testFidelityTierComparable() {
        XCTAssertTrue(FidelityTier.markdown < .markdownWithFigures)
        XCTAssertTrue(FidelityTier.markdownWithFigures < .marker)
        XCTAssertTrue(FidelityTier.markdown < .marker)
        XCTAssertFalse(FidelityTier.marker < .markdown)
    }

    func testDefaultFidelityIsMarkdown() {
        let options = ConversionOptions.default
        XCTAssertEqual(options.fidelity, .markdown)
        XCTAssertFalse(options.useHTMLExtensions)
        XCTAssertNil(options.figuresDirectory)
        XCTAssertNil(options.metadataOutput)
    }

    // MARK: - Mixed Content

    func testMixedRunsAndHyperlinks() throws {
        var para = Paragraph()
        para.runs = [
            Run(text: "Click "),
            Run(text: "here", properties: RunProperties(bold: true)),
            Run(text: " or "),
        ]
        para.hyperlinks = [
            Hyperlink(id: "h1", text: "this link", url: "https://example.com", relationshipId: "rId1")
        ]
        var doc = WordDocument()
        doc.hyperlinkReferences = [HyperlinkReference(relationshipId: "rId1", url: "https://example.com")]
        doc.appendParagraph(para)

        let md = try convert(doc)
        XCTAssertTrue(md.contains("Click **here** or [this link](https://example.com)"), "Got: \(md)")
    }

    func testEmptyParagraphSkipped() throws {
        let doc = makeDocument(paragraphs: [
            Paragraph(text: "first"),
            Paragraph(text: "   "),
            Paragraph(text: "second"),
        ])
        let md = try convert(doc)
        XCTAssertFalse(md.contains("\n\n\n\n"), "Empty paragraphs should be skipped")
    }

    // MARK: - Practical Mode: Heading Heuristic

    func testHeadingHeuristicInfersFromFontSize() throws {
        // 建立：3 段 body (fontSize 24 = 12pt)，1 段 heading (fontSize 36 = 18pt, bold)
        let bodyRun = Run(text: "Body paragraph", properties: RunProperties(fontSize: 24))
        let headingRun = Run(text: "Chapter Title", properties: RunProperties(bold: true, fontSize: 36))

        let doc = makeDocument(paragraphs: [
            Paragraph(runs: [headingRun]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
        ])

        // headingHeuristic 預設 true
        let md = try convert(doc)
        XCTAssertTrue(md.contains("# **Chapter Title**"), "Heading heuristic should infer H1. Got: \(md)")
    }

    func testHeadingHeuristicMultipleLevels() throws {
        let bodyRun = Run(text: "Normal text", properties: RunProperties(fontSize: 24))
        let h1Run = Run(text: "Main Title", properties: RunProperties(bold: true, fontSize: 48))
        let h2Run = Run(text: "Sub Title", properties: RunProperties(bold: true, fontSize: 36))

        let doc = makeDocument(paragraphs: [
            Paragraph(runs: [h1Run]),
            Paragraph(runs: [h2Run]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
        ])

        let md = try convert(doc)
        XCTAssertTrue(md.contains("# **Main Title**"), "Should infer H1. Got: \(md)")
        XCTAssertTrue(md.contains("## **Sub Title**"), "Should infer H2. Got: \(md)")
    }

    func testHeadingHeuristicDisabled() throws {
        let bodyRun = Run(text: "Body text", properties: RunProperties(fontSize: 24))
        let headingRun = Run(text: "Title", properties: RunProperties(bold: true, fontSize: 36))

        let doc = makeDocument(paragraphs: [
            Paragraph(runs: [headingRun]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
        ])

        var options = ConversionOptions.default
        options.headingHeuristic = false
        let md = try convert(doc, options: options)
        XCTAssertFalse(md.contains("# Title"), "Heuristic disabled, no heading prefix. Got: \(md)")
        XCTAssertTrue(md.contains("**Title**"), "Bold should still be applied. Got: \(md)")
    }

    func testHeadingHeuristicStyleBasedTakesPriority() throws {
        // 有 Heading style 的段落不應被 heuristic 覆蓋
        let bodyRun = Run(text: "Body text", properties: RunProperties(fontSize: 24))
        let headingRun = Run(text: "Styled Heading", properties: RunProperties(bold: true, fontSize: 36))
        var styledPara = Paragraph(runs: [headingRun])
        styledPara.properties.style = "Heading2"

        let doc = makeDocument(paragraphs: [
            styledPara,
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
            Paragraph(runs: [bodyRun]),
        ])

        let md = try convert(doc)
        // style-based 偵測到 Heading2 → ##
        XCTAssertTrue(md.contains("## **Styled Heading**"), "Style-based should take priority. Got: \(md)")
        // contains("# **Styled") 會子字串匹配到 "## **Styled"，改用 line-level 檢查
        let hasH1 = md.components(separatedBy: "\n").contains { $0 == "# **Styled Heading**" }
        XCTAssertFalse(hasH1, "Should be H2, not H1. Got: \(md)")
    }

    // MARK: - Practical Mode: FigureExtractor EMF→PNG

    func testFigureExtractorConvertsNonWebFriendly() throws {
        let tmpDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("word-to-md-test-\(UUID().uuidString)")
        let figDir = tmpDir.appendingPathComponent("figures")

        defer { try? FileManager.default.removeItem(at: tmpDir) }

        // 建立一個 1x1 紅色像素的 PNG 作為測試圖片
        let pngData = createTestPNGData()

        // 模擬 EMF 檔（實際上是 PNG 資料，但副檔名是 .emf）
        let imageRef = ImageReference(
            id: "rId10",
            fileName: "image1.emf",
            contentType: "image/x-emf",
            data: pngData
        )

        let drawing = Drawing(
            type: .inline, width: 914400, height: 914400,
            imageId: "rId10", name: "chart", description: "test chart"
        )
        var run = Run(text: "")
        run.drawing = drawing

        var doc = WordDocument()
        doc.images = [imageRef]
        doc.appendParagraph(Paragraph(runs: [run]))

        let options = ConversionOptions(
            fidelity: .markdownWithFigures,
            figuresDirectory: figDir
        )
        let md = try converter.convertToString(document: doc, options: options)

        // 應該輸出 .png 而非 .emf
        XCTAssertTrue(md.contains("figures/image1.png"), "EMF should be converted to PNG. Got: \(md)")
        XCTAssertFalse(md.contains("image1.emf"), "EMF extension should not appear. Got: \(md)")

        // 確認檔案存在
        let pngFile = figDir.appendingPathComponent("image1.png")
        XCTAssertTrue(FileManager.default.fileExists(atPath: pngFile.path), "PNG file should exist at \(pngFile.path)")
    }

    func testFigureExtractorPreservesWebFriendly() throws {
        let tmpDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("word-to-md-test-\(UUID().uuidString)")
        let figDir = tmpDir.appendingPathComponent("figures")

        defer { try? FileManager.default.removeItem(at: tmpDir) }

        let pngData = createTestPNGData()

        let imageRef = ImageReference(
            id: "rId11",
            fileName: "photo.png",
            contentType: "image/png",
            data: pngData
        )

        let drawing = Drawing(
            type: .inline, width: 914400, height: 914400,
            imageId: "rId11", name: "photo"
        )
        var run = Run(text: "")
        run.drawing = drawing

        var doc = WordDocument()
        doc.images = [imageRef]
        doc.appendParagraph(Paragraph(runs: [run]))

        let options = ConversionOptions(
            fidelity: .markdownWithFigures,
            figuresDirectory: figDir
        )
        let md = try converter.convertToString(document: doc, options: options)

        // PNG 應保持原始副檔名
        XCTAssertTrue(md.contains("figures/photo.png"), "PNG should keep extension. Got: \(md)")
    }

    /// 建立最小的有效 PNG 資料（1x1 紅色像素）
    private func createTestPNGData() -> Data {
        #if canImport(AppKit)
        let image = NSImage(size: NSSize(width: 1, height: 1))
        image.lockFocus()
        NSColor.red.set()
        NSRect(x: 0, y: 0, width: 1, height: 1).fill()
        image.unlockFocus()
        let tiff = image.tiffRepresentation!
        let bitmap = NSBitmapImageRep(data: tiff)!
        return bitmap.representation(using: .png, properties: [:])!
        #else
        fatalError("AppKit required for test")
        #endif
    }
}
