import XCTest
@testable import OOXMLSwift

final class RunTests: XCTestCase {

    // MARK: - RunProperties.toXML()

    func testEmptyProperties() {
        let props = RunProperties()
        XCTAssertEqual(props.toXML(), "")
    }

    func testBold() {
        var props = RunProperties()
        props.bold = true
        XCTAssertEqual(props.toXML(), "<w:b/>")
    }

    func testItalic() {
        var props = RunProperties()
        props.italic = true
        XCTAssertEqual(props.toXML(), "<w:i/>")
    }

    func testBoldItalic() {
        var props = RunProperties()
        props.bold = true
        props.italic = true
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:b/>"))
        XCTAssertTrue(xml.contains("<w:i/>"))
    }

    func testUnderlineSingle() {
        var props = RunProperties()
        props.underline = .single
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"single\"/>"))
    }

    func testUnderlineDouble() {
        var props = RunProperties()
        props.underline = .double
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"double\"/>"))
    }

    func testUnderlineDotted() {
        var props = RunProperties()
        props.underline = .dotted
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"dotted\"/>"))
    }

    func testUnderlineWave() {
        var props = RunProperties()
        props.underline = .wave
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"wave\"/>"))
    }

    func testStrikethrough() {
        var props = RunProperties()
        props.strikethrough = true
        XCTAssertEqual(props.toXML(), "<w:strike/>")
    }

    func testFontSize() {
        var props = RunProperties()
        props.fontSize = 24  // 12pt
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:sz w:val=\"24\"/>"))
        XCTAssertTrue(xml.contains("<w:szCs w:val=\"24\"/>"))
    }

    func testFontName() {
        var props = RunProperties()
        props.fontName = "Arial"
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:ascii=\"Arial\""))
        XCTAssertTrue(xml.contains("w:hAnsi=\"Arial\""))
        XCTAssertTrue(xml.contains("w:eastAsia=\"Arial\""))
        XCTAssertTrue(xml.contains("w:cs=\"Arial\""))
    }

    func testColor() {
        var props = RunProperties()
        props.color = "FF0000"
        XCTAssertTrue(props.toXML().contains("<w:color w:val=\"FF0000\"/>"))
    }

    func testHighlightYellow() {
        var props = RunProperties()
        props.highlight = .yellow
        XCTAssertTrue(props.toXML().contains("<w:highlight w:val=\"yellow\"/>"))
    }

    func testHighlightCyan() {
        var props = RunProperties()
        props.highlight = .cyan
        XCTAssertTrue(props.toXML().contains("<w:highlight w:val=\"cyan\"/>"))
    }

    func testVerticalAlignSuperscript() {
        var props = RunProperties()
        props.verticalAlign = .superscript
        XCTAssertTrue(props.toXML().contains("<w:vertAlign w:val=\"superscript\"/>"))
    }

    func testVerticalAlignSubscript() {
        var props = RunProperties()
        props.verticalAlign = .subscript
        XCTAssertTrue(props.toXML().contains("<w:vertAlign w:val=\"subscript\"/>"))
    }

    func testSmallCaps() {
        var props = RunProperties()
        props.smallCaps = true
        XCTAssertTrue(props.toXML().contains("<w:smallCaps/>"))
    }

    func testAllCaps() {
        var props = RunProperties()
        props.allCaps = true
        XCTAssertTrue(props.toXML().contains("<w:caps/>"))
    }

    func testCharacterSpacing() {
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(spacing: 20)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:spacing w:val=\"20\"/>"))
    }

    func testCharacterSpacingWithPosition() {
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(spacing: 10, position: 5)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:spacing w:val=\"10\"/>"))
        XCTAssertTrue(xml.contains("<w:position w:val=\"5\"/>"))
    }

    func testCharacterSpacingWithKern() {
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(kern: 16)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:kern w:val=\"16\"/>"))
    }

    // MARK: - Run.toXML()

    func testSimpleRun() {
        let run = Run(text: "Hello")
        let xml = run.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:r>"))
        XCTAssertTrue(xml.hasSuffix("</w:r>"))
        XCTAssertTrue(xml.contains("<w:t xml:space=\"preserve\">Hello</w:t>"))
    }

    func testRunWithProperties() {
        let run = Run(text: "Bold", properties: RunProperties(bold: true))
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("<w:rPr><w:b/></w:rPr>"))
        XCTAssertTrue(xml.contains("Bold"))
    }

    func testRunXMLEscaping() {
        let run = Run(text: "A < B & C > D")
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("A &lt; B &amp; C &gt; D"))
    }

    func testRunWithRawXML() {
        var run = Run(text: "ignored")
        run.rawXML = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        let xml = run.toXML()
        XCTAssertEqual(xml, "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
    }

    func testRunPropertiesRawXML() {
        var run = Run(text: "ignored")
        run.properties.rawXML = "<w:sdt><w:sdtContent/></w:sdt>"
        let xml = run.toXML()
        XCTAssertEqual(xml, "<w:sdt><w:sdtContent/></w:sdt>")
    }

    // MARK: - RunProperties Merge

    func testMergeProperties() {
        var base = RunProperties()
        base.bold = true

        var overlay = RunProperties()
        overlay.italic = true
        overlay.fontSize = 24

        base.merge(with: overlay)
        XCTAssertTrue(base.bold)
        XCTAssertTrue(base.italic)
        XCTAssertEqual(base.fontSize, 24)
    }

    func testMergePropertiesIncludesCaps() {
        var base = RunProperties()

        var overlay = RunProperties()
        overlay.smallCaps = true
        overlay.allCaps = true

        base.merge(with: overlay)
        XCTAssertTrue(base.smallCaps)
        XCTAssertTrue(base.allCaps)
    }

    // MARK: - Combined Properties

    func testFullyFormattedRun() {
        var props = RunProperties()
        props.bold = true
        props.italic = true
        props.underline = .single
        props.fontSize = 28
        props.fontName = "Times New Roman"
        props.color = "0000FF"

        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:b/>"))
        XCTAssertTrue(xml.contains("<w:i/>"))
        XCTAssertTrue(xml.contains("<w:u w:val=\"single\"/>"))
        XCTAssertTrue(xml.contains("<w:sz w:val=\"28\"/>"))
        XCTAssertTrue(xml.contains("w:ascii=\"Times New Roman\""))
        XCTAssertTrue(xml.contains("<w:color w:val=\"0000FF\"/>"))
    }
}
