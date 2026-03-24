import XCTest
@testable import OOXMLSwift

final class TableTests: XCTestCase {

    // MARK: - TableProperties.toXML()

    func testEmptyTableProperties() {
        let props = TableProperties()
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:tblPr>"))
        XCTAssertTrue(xml.contains("</w:tblPr>"))
    }

    func testTableWidth() {
        var props = TableProperties()
        props.width = 9000
        props.widthType = .dxa
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:tblW w:w=\"9000\" w:type=\"dxa\"/>"))
    }

    func testTableAlignment() {
        var props = TableProperties()
        props.alignment = .center
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:jc w:val=\"center\"/>"))
    }

    func testTableLayout() {
        var props = TableProperties()
        props.layout = .fixed
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:tblLayout w:type=\"fixed\"/>"))
    }

    func testTableBorders() {
        var props = TableProperties()
        props.borders = TableBorders.all(Border(style: .single, size: 4, color: "000000"))
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:tblBorders>"))
        XCTAssertTrue(xml.contains("<w:top w:val=\"single\" w:sz=\"4\" w:color=\"000000\"/>"))
        XCTAssertTrue(xml.contains("<w:insideH"))
        XCTAssertTrue(xml.contains("<w:insideV"))
    }

    func testTableCellMargins() {
        var props = TableProperties()
        props.cellMargins = TableCellMargins.all(108)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:tblCellMar>"))
        XCTAssertTrue(xml.contains("<w:top w:w=\"108\" w:type=\"dxa\"/>"))
        XCTAssertTrue(xml.contains("<w:bottom w:w=\"108\" w:type=\"dxa\"/>"))
    }

    // MARK: - TableRow.toXML()

    func testSimpleRow() {
        let row = TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")])
        let xml = row.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:tr>"))
        XCTAssertTrue(xml.hasSuffix("</w:tr>"))
        XCTAssertTrue(xml.contains("<w:tc>"))
    }

    func testHeaderRow() {
        var rowProps = TableRowProperties()
        rowProps.isHeader = true
        let row = TableRow(cells: [TableCell(text: "Header")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:tblHeader/>"))
    }

    func testRowHeight() {
        var rowProps = TableRowProperties()
        rowProps.height = 400
        rowProps.heightRule = .exact
        let row = TableRow(cells: [TableCell(text: "Fixed")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:trHeight w:val=\"400\" w:hRule=\"exact\"/>"))
    }

    func testRowCantSplit() {
        var rowProps = TableRowProperties()
        rowProps.cantSplit = true
        let row = TableRow(cells: [TableCell(text: "NoSplit")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:cantSplit/>"))
    }

    // MARK: - TableCell.toXML()

    func testSimpleCell() {
        let cell = TableCell(text: "Content")
        let xml = cell.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:tc>"))
        XCTAssertTrue(xml.hasSuffix("</w:tc>"))
        XCTAssertTrue(xml.contains("Content"))
    }

    func testCellWidth() {
        var cellProps = TableCellProperties()
        cellProps.width = 3000
        cellProps.widthType = .dxa
        let cell = TableCell(paragraphs: [Paragraph(text: "W")], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:tcW w:w=\"3000\" w:type=\"dxa\"/>"))
    }

    func testCellGridSpan() {
        var cellProps = TableCellProperties()
        cellProps.gridSpan = 3
        let cell = TableCell(paragraphs: [Paragraph(text: "Merged")], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:gridSpan w:val=\"3\"/>"))
    }

    func testCellVerticalMerge() {
        var cellProps = TableCellProperties()
        cellProps.verticalMerge = .restart
        let cell = TableCell(paragraphs: [Paragraph(text: "Start")], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:vMerge w:val=\"restart\"/>"))
    }

    func testCellVerticalAlignment() {
        var cellProps = TableCellProperties()
        cellProps.verticalAlignment = .center
        let cell = TableCell(paragraphs: [Paragraph(text: "Centered")], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:vAlign w:val=\"center\"/>"))
    }

    func testCellBorders() {
        var cellProps = TableCellProperties()
        cellProps.borders = CellBorders.all(Border(style: .double, size: 6, color: "FF0000"))
        let cell = TableCell(paragraphs: [Paragraph(text: "Bordered")], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:tcBorders>"))
        XCTAssertTrue(xml.contains("w:val=\"double\""))
        XCTAssertTrue(xml.contains("w:color=\"FF0000\""))
    }

    func testCellShading() {
        var cellProps = TableCellProperties()
        cellProps.shading = CellShading.solid("CCCCCC")
        let cell = TableCell(paragraphs: [Paragraph(text: "Shaded")], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("w:fill=\"CCCCCC\""))
    }

    func testMultiParagraphCell() {
        let cell = TableCell(paragraphs: [
            Paragraph(text: "Line 1"),
            Paragraph(text: "Line 2")
        ])
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("Line 1"))
        XCTAssertTrue(xml.contains("Line 2"))
        // Should have two <w:p> elements
        let pCount = xml.components(separatedBy: "<w:p>").count - 1
        XCTAssertEqual(pCount, 2)
    }

    func testEmptyCellHasDefaultParagraph() {
        let cell = TableCell(paragraphs: [])
        let xml = cell.toXML()
        // Even empty cells must have at least one <w:p>
        XCTAssertTrue(xml.contains("<w:p>"))
    }

    // MARK: - Table.toXML()

    func testSimpleTable() {
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")]),
            TableRow(cells: [TableCell(text: "C"), TableCell(text: "D")])
        ])
        let xml = table.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:tbl>"))
        XCTAssertTrue(xml.hasSuffix("</w:tbl>"))
        XCTAssertTrue(xml.contains("<w:tblGrid>"))
        XCTAssertTrue(xml.contains("<w:gridCol"))
        XCTAssertTrue(xml.contains("A"))
        XCTAssertTrue(xml.contains("D"))
    }

    func testTableGrid() {
        var cellProps = TableCellProperties()
        cellProps.width = 4500
        let table = Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [Paragraph(text: "Col1")], properties: cellProps),
                TableCell(paragraphs: [Paragraph(text: "Col2")], properties: cellProps)
            ])
        ])
        let xml = table.toXML()
        XCTAssertTrue(xml.contains("<w:gridCol w:w=\"4500\"/>"))
    }

    func testTableGetText() {
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")]),
            TableRow(cells: [TableCell(text: "C"), TableCell(text: "D")])
        ])
        let text = table.getText()
        XCTAssertTrue(text.contains("A"))
        XCTAssertTrue(text.contains("B"))
        XCTAssertTrue(text.contains("C"))
        XCTAssertTrue(text.contains("D"))
    }

    func testConvenienceInitRowColumn() {
        let table = Table(rowCount: 3, columnCount: 2)
        XCTAssertEqual(table.rows.count, 3)
        XCTAssertEqual(table.rows[0].cells.count, 2)
    }
}
