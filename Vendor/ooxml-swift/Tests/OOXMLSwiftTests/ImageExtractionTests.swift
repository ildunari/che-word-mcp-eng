import XCTest
@testable import OOXMLSwift

final class ImageExtractionTests: XCTestCase {

    // MARK: - Relationship Parsing Tests

    func testRelationshipTypeImage() {
        let type = RelationshipType(rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        XCTAssertEqual(type, .image)
    }

    func testRelationshipTypeHyperlink() {
        let type = RelationshipType(rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
        XCTAssertEqual(type, .hyperlink)
    }

    func testRelationshipTypeUnknown() {
        let type = RelationshipType(rawValue: "http://unknown.type")
        XCTAssertEqual(type, .unknown)
    }

    // MARK: - Relationships Collection Tests

    func testRelationshipsCollectionGetById() {
        var collection = RelationshipsCollection()
        let rel1 = Relationship(id: "rId1", type: .image, target: "media/image1.png")
        let rel2 = Relationship(id: "rId2", type: .hyperlink, target: "http://example.com")
        collection.relationships = [rel1, rel2]

        XCTAssertNotNil(collection.get(by: "rId1"))
        XCTAssertEqual(collection.get(by: "rId1")?.target, "media/image1.png")
        XCTAssertNil(collection.get(by: "rId3"))
    }

    func testRelationshipsCollectionImageFiltering() {
        var collection = RelationshipsCollection()
        let rel1 = Relationship(id: "rId1", type: .image, target: "media/image1.png")
        let rel2 = Relationship(id: "rId2", type: .hyperlink, target: "http://example.com")
        let rel3 = Relationship(id: "rId3", type: .image, target: "media/image2.jpg")
        collection.relationships = [rel1, rel2, rel3]

        XCTAssertEqual(collection.imageRelationships.count, 2)
        XCTAssertEqual(collection.hyperlinkRelationships.count, 1)
    }

    // MARK: - Image Reference Tests

    func testImageReferenceMIMEType() {
        let pngRef = ImageReference(id: "rId1", fileName: "image1.png", contentType: "image/png", data: Data())
        XCTAssertEqual(pngRef.contentType, "image/png")

        let jpegRef = ImageReference(id: "rId2", fileName: "image2.jpg", contentType: "image/jpeg", data: Data())
        XCTAssertEqual(jpegRef.contentType, "image/jpeg")
    }

    // MARK: - Drawing Tests

    func testDrawingFromPixels() {
        let drawing = Drawing.from(widthPx: 100, heightPx: 200, imageId: "rId1", name: "TestImage")

        // 1 pixel = 9525 EMU @ 96 DPI
        XCTAssertEqual(drawing.width, 100 * 9525)
        XCTAssertEqual(drawing.height, 200 * 9525)
        XCTAssertEqual(drawing.imageId, "rId1")
        XCTAssertEqual(drawing.name, "TestImage")
        XCTAssertEqual(drawing.type, .inline)
    }

    func testDrawingWidthInPixels() {
        let drawing = Drawing(type: .inline, width: 952500, height: 1905000, imageId: "rId1")
        XCTAssertEqual(drawing.widthInPixels, 100)
        XCTAssertEqual(drawing.heightInPixels, 200)
    }

    func testAnchorDrawing() {
        let position = AnchorPosition.centeredOnPage()
        let drawing = Drawing.anchor(width: 1000000, height: 500000, imageId: "rId1", position: position)

        XCTAssertEqual(drawing.type, .anchor)
        XCTAssertEqual(drawing.anchorPosition.horizontalAlignment, .center)
        XCTAssertEqual(drawing.anchorPosition.verticalAlignment, .center)
    }
}
