import XCTest
@testable import OOXMLSwift

final class ParagraphTests: XCTestCase {

    // MARK: - ParagraphProperties.toXML()

    func testEmptyProperties() {
        let props = ParagraphProperties()
        XCTAssertEqual(props.toXML(), "")
    }

    func testAlignmentCenter() {
        var props = ParagraphProperties()
        props.alignment = .center
        XCTAssertTrue(props.toXML().contains("<w:jc w:val=\"center\"/>"))
    }

    func testAlignmentRight() {
        var props = ParagraphProperties()
        props.alignment = .right
        XCTAssertTrue(props.toXML().contains("<w:jc w:val=\"right\"/>"))
    }

    func testAlignmentBoth() {
        var props = ParagraphProperties()
        props.alignment = .both
        XCTAssertTrue(props.toXML().contains("<w:jc w:val=\"both\"/>"))
    }

    func testAlignmentDistribute() {
        var props = ParagraphProperties()
        props.alignment = .distribute
        XCTAssertTrue(props.toXML().contains("<w:jc w:val=\"distribute\"/>"))
    }

    func testSpacingBefore() {
        var props = ParagraphProperties()
        props.spacing = Spacing(before: 240)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:before=\"240\""))
    }

    func testSpacingAfter() {
        var props = ParagraphProperties()
        props.spacing = Spacing(after: 120)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:after=\"120\""))
    }

    func testSpacingLine() {
        var props = ParagraphProperties()
        props.spacing = Spacing(line: 360, lineRule: .exact)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:line=\"360\""))
        XCTAssertTrue(xml.contains("w:lineRule=\"exact\""))
    }

    func testSpacingFull() {
        var props = ParagraphProperties()
        props.spacing = Spacing(before: 240, after: 120, line: 360, lineRule: .atLeast)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:before=\"240\""))
        XCTAssertTrue(xml.contains("w:after=\"120\""))
        XCTAssertTrue(xml.contains("w:line=\"360\""))
        XCTAssertTrue(xml.contains("w:lineRule=\"atLeast\""))
    }

    func testIndentationLeft() {
        var props = ParagraphProperties()
        props.indentation = Indentation(left: 720)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:left=\"720\""))
    }

    func testIndentationRight() {
        var props = ParagraphProperties()
        props.indentation = Indentation(right: 360)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:right=\"360\""))
    }

    func testIndentationFirstLine() {
        var props = ParagraphProperties()
        props.indentation = Indentation(firstLine: 480)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:firstLine=\"480\""))
    }

    func testIndentationHanging() {
        var props = ParagraphProperties()
        props.indentation = Indentation(hanging: 360)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:hanging=\"360\""))
    }

    func testKeepNext() {
        var props = ParagraphProperties()
        props.keepNext = true
        XCTAssertTrue(props.toXML().contains("<w:keepNext/>"))
    }

    func testKeepLines() {
        var props = ParagraphProperties()
        props.keepLines = true
        XCTAssertTrue(props.toXML().contains("<w:keepLines/>"))
    }

    func testPageBreakBefore() {
        var props = ParagraphProperties()
        props.pageBreakBefore = true
        XCTAssertTrue(props.toXML().contains("<w:pageBreakBefore/>"))
    }

    func testParagraphBorder() {
        var props = ParagraphProperties()
        props.border = ParagraphBorder.all(ParagraphBorderStyle(type: .single, color: "FF0000", size: 8, space: 2))
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:pBdr>"))
        XCTAssertTrue(xml.contains("w:val=\"single\""))
        XCTAssertTrue(xml.contains("w:color=\"FF0000\""))
        XCTAssertTrue(xml.contains("w:sz=\"8\""))
    }

    func testParagraphShading() {
        var props = ParagraphProperties()
        props.shading = ParagraphShading(fill: "FFFF00")
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:shd"))
        XCTAssertTrue(xml.contains("w:fill=\"FFFF00\""))
    }

    func testStyle() {
        var props = ParagraphProperties()
        props.style = "Heading1"
        XCTAssertTrue(props.toXML().contains("<w:pStyle w:val=\"Heading1\"/>"))
    }

    func testNumbering() {
        var props = ParagraphProperties()
        props.numbering = NumberingInfo(numId: 1, level: 2)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:numPr>"))
        XCTAssertTrue(xml.contains("<w:ilvl w:val=\"2\"/>"))
        XCTAssertTrue(xml.contains("<w:numId w:val=\"1\"/>"))
    }

    // MARK: - Paragraph.toXML()

    func testSimpleParagraph() {
        let para = Paragraph(text: "Hello World")
        let xml = para.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:p>"))
        XCTAssertTrue(xml.hasSuffix("</w:p>"))
        XCTAssertTrue(xml.contains("Hello World"))
    }

    func testParagraphWithProperties() {
        var props = ParagraphProperties()
        props.alignment = .center
        let para = Paragraph(text: "Centered", properties: props)
        let xml = para.toXML()
        XCTAssertTrue(xml.contains("<w:pPr>"))
        XCTAssertTrue(xml.contains("<w:jc w:val=\"center\"/>"))
        XCTAssertTrue(xml.contains("Centered"))
    }

    func testParagraphWithPageBreak() {
        var para = Paragraph(text: "After break")
        para.hasPageBreak = true
        let xml = para.toXML()
        XCTAssertTrue(xml.contains("<w:br w:type=\"page\"/>"))
    }

    func testParagraphWithCommentIds() {
        var para = Paragraph(text: "Commented")
        para.commentIds = [1, 2]
        let xml = para.toXML()
        XCTAssertTrue(xml.contains("<w:commentRangeStart w:id=\"1\"/>"))
        XCTAssertTrue(xml.contains("<w:commentRangeStart w:id=\"2\"/>"))
        XCTAssertTrue(xml.contains("<w:commentRangeEnd w:id=\"1\"/>"))
        XCTAssertTrue(xml.contains("<w:commentRangeEnd w:id=\"2\"/>"))
        XCTAssertTrue(xml.contains("<w:commentReference w:id=\"1\"/>"))
    }

    func testParagraphWithFootnoteIds() {
        var para = Paragraph(text: "Text")
        para.footnoteIds = [1]
        let xml = para.toXML()
        XCTAssertTrue(xml.contains("<w:footnoteReference w:id=\"1\"/>"))
        XCTAssertTrue(xml.contains("FootnoteReference"))
    }

    func testParagraphWithEndnoteIds() {
        var para = Paragraph(text: "Text")
        para.endnoteIds = [1]
        let xml = para.toXML()
        XCTAssertTrue(xml.contains("<w:endnoteReference w:id=\"1\"/>"))
        XCTAssertTrue(xml.contains("EndnoteReference"))
    }

    func testParagraphWithSectionBreak() {
        var props = ParagraphProperties()
        props.sectionBreak = .nextPage
        let para = Paragraph(text: "", properties: props)
        let xml = para.toXML()
        XCTAssertTrue(xml.contains("<w:sectPr>"))
        XCTAssertTrue(xml.contains("<w:type w:val=\"nextPage\"/>"))
    }

    // MARK: - Convenience Initializers

    func testSpacingPoints() {
        let spacing = Spacing.points(before: 12, after: 6, lineSpacing: 1.5)
        XCTAssertEqual(spacing.before, 240)   // 12 * 20
        XCTAssertEqual(spacing.after, 120)     // 6 * 20
        XCTAssertEqual(spacing.line, 360)      // 1.5 * 240
        XCTAssertEqual(spacing.lineRule, .exact)
    }

    func testIndentationCharacters() {
        let indent = Indentation.characters(left: 2, firstLine: 1)
        XCTAssertEqual(indent.left, 480)       // 2 * 240
        XCTAssertEqual(indent.firstLine, 240)  // 1 * 240
    }

    func testParagraphGetText() {
        let para = Paragraph(runs: [
            Run(text: "Hello "),
            Run(text: "World")
        ])
        XCTAssertEqual(para.getText(), "Hello World")
    }

    func testParagraphMergeProperties() {
        var base = ParagraphProperties()
        base.alignment = .left

        var overlay = ParagraphProperties()
        overlay.alignment = .center
        overlay.keepNext = true

        base.merge(with: overlay)
        XCTAssertEqual(base.alignment, .center)
        XCTAssertTrue(base.keepNext)
    }
}
