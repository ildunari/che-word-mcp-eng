import XCTest
@testable import OOXMLSwift

final class NumberingTests: XCTestCase {

    // MARK: - Level.toXML()

    func testBulletLevel() {
        let level = Level(ilvl: 0, start: 1, numFmt: .bullet, lvlText: "\u{F0B7}", indent: 720, fontName: "Symbol")
        let xml = level.toXML()
        XCTAssertTrue(xml.contains("<w:lvl w:ilvl=\"0\">"))
        XCTAssertTrue(xml.contains("<w:numFmt w:val=\"bullet\"/>"))
        XCTAssertTrue(xml.contains("<w:start w:val=\"1\"/>"))
        XCTAssertTrue(xml.contains("<w:ind w:left=\"720\" w:hanging=\"360\"/>"))
        XCTAssertTrue(xml.contains("w:ascii=\"Symbol\""))
    }

    func testDecimalLevel() {
        let level = Level(ilvl: 0, start: 1, numFmt: .decimal, lvlText: "%1.", indent: 720)
        let xml = level.toXML()
        XCTAssertTrue(xml.contains("<w:numFmt w:val=\"decimal\"/>"))
        XCTAssertTrue(xml.contains("<w:lvlText w:val=\"%1.\"/>"))
        // Decimal levels should NOT have rPr with font
        XCTAssertFalse(xml.contains("<w:rPr>"))
    }

    func testLowerLetterLevel() {
        let level = Level(ilvl: 1, start: 1, numFmt: .lowerLetter, lvlText: "%2.", indent: 1440)
        let xml = level.toXML()
        XCTAssertTrue(xml.contains("<w:numFmt w:val=\"lowerLetter\"/>"))
        XCTAssertTrue(xml.contains("<w:lvlText w:val=\"%2.\"/>"))
        XCTAssertTrue(xml.contains("<w:ind w:left=\"1440\""))
    }

    func testLowerRomanLevel() {
        let level = Level(ilvl: 2, start: 1, numFmt: .lowerRoman, lvlText: "%3.", indent: 2160)
        let xml = level.toXML()
        XCTAssertTrue(xml.contains("<w:numFmt w:val=\"lowerRoman\"/>"))
    }

    func testUpperLetterLevel() {
        let level = Level(ilvl: 0, start: 1, numFmt: .upperLetter, lvlText: "%1)", indent: 720)
        let xml = level.toXML()
        XCTAssertTrue(xml.contains("<w:numFmt w:val=\"upperLetter\"/>"))
    }

    func testUpperRomanLevel() {
        let level = Level(ilvl: 0, start: 1, numFmt: .upperRoman, lvlText: "%1.", indent: 720)
        let xml = level.toXML()
        XCTAssertTrue(xml.contains("<w:numFmt w:val=\"upperRoman\"/>"))
    }

    // MARK: - AbstractNum.toXML()

    func testAbstractNum() {
        let abstractNum = AbstractNum(
            abstractNumId: 0,
            levels: [
                Level(ilvl: 0, start: 1, numFmt: .bullet, lvlText: "\u{F0B7}", indent: 720, fontName: "Symbol")
            ]
        )
        let xml = abstractNum.toXML()
        XCTAssertTrue(xml.contains("<w:abstractNum w:abstractNumId=\"0\">"))
        XCTAssertTrue(xml.contains("</w:abstractNum>"))
        XCTAssertTrue(xml.contains("<w:lvl w:ilvl=\"0\">"))
    }

    // MARK: - Num.toXML()

    func testNum() {
        let num = Num(numId: 1, abstractNumId: 0)
        let xml = num.toXML()
        XCTAssertEqual(xml, "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>")
    }

    // MARK: - Numbering.toXML()

    func testNumberingXML() {
        var numbering = Numbering()
        numbering.abstractNums = [
            AbstractNum(
                abstractNumId: 0,
                levels: [Level(ilvl: 0, start: 1, numFmt: .bullet, lvlText: "\u{F0B7}", indent: 720, fontName: "Symbol")]
            )
        ]
        numbering.nums = [Num(numId: 1, abstractNumId: 0)]

        let xml = numbering.toXML()
        XCTAssertTrue(xml.contains("<?xml"))
        XCTAssertTrue(xml.contains("<w:numbering"))
        XCTAssertTrue(xml.contains("<w:abstractNum"))
        XCTAssertTrue(xml.contains("<w:num w:numId=\"1\">"))
        XCTAssertTrue(xml.contains("</w:numbering>"))
    }

    // MARK: - Numbering convenience methods

    func testCreateBulletList() {
        var numbering = Numbering()
        let numId = numbering.createBulletList()
        XCTAssertEqual(numId, 1)
        XCTAssertEqual(numbering.abstractNums.count, 1)
        XCTAssertEqual(numbering.nums.count, 1)
        XCTAssertEqual(numbering.abstractNums[0].levels.count, 9)  // 9 levels
        XCTAssertEqual(numbering.abstractNums[0].levels[0].numFmt, .bullet)
    }

    func testCreateNumberedList() {
        var numbering = Numbering()
        let numId = numbering.createNumberedList()
        XCTAssertEqual(numId, 1)
        XCTAssertEqual(numbering.abstractNums.count, 1)
        XCTAssertEqual(numbering.abstractNums[0].levels[0].numFmt, .decimal)
        XCTAssertEqual(numbering.abstractNums[0].levels[1].numFmt, .lowerLetter)
        XCTAssertEqual(numbering.abstractNums[0].levels[2].numFmt, .lowerRoman)
    }

    func testMultipleLists() {
        var numbering = Numbering()
        let bulletId = numbering.createBulletList()
        let numberedId = numbering.createNumberedList()
        XCTAssertEqual(bulletId, 1)
        XCTAssertEqual(numberedId, 2)
        XCTAssertEqual(numbering.abstractNums.count, 2)
        XCTAssertEqual(numbering.nums.count, 2)
    }

    func testNextAbstractNumId() {
        var numbering = Numbering()
        XCTAssertEqual(numbering.nextAbstractNumId, 0)
        _ = numbering.createBulletList()
        XCTAssertEqual(numbering.nextAbstractNumId, 1)
    }

    func testNextNumId() {
        var numbering = Numbering()
        XCTAssertEqual(numbering.nextNumId, 1)
        _ = numbering.createBulletList()
        XCTAssertEqual(numbering.nextNumId, 2)
    }
}
