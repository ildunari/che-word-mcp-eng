import XCTest
import OOXMLSwift
@testable import CheWordMCP

final class InlineEditingTests: XCTestCase {
    func testReplaceTextRangePreservesUnaffectedFormattingAcrossRuns() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello "),
                TestFixtures.makeRun("world", bold: true),
                TestFixtures.makeRun("!")
            ])
        ])

        try document.replaceTextRange(at: 0, start: 6, end: 11, replacement: "team", replacementProperties: nil)

        let runs = document.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "team", "!"])
        XCTAssertFalse(runs[0].properties.bold)
        XCTAssertTrue(runs[1].properties.bold)
        XCTAssertFalse(runs[2].properties.bold)
    }

    func testReplaceTextRangeSplitsSingleRunAtInsertionPoint() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello world", italic: true)
            ])
        ])

        try document.replaceTextRange(at: 0, start: 5, end: 5, replacement: ", brave", replacementProperties: nil)

        let runs = document.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello", ", brave", " world"])
        XCTAssertTrue(runs.allSatisfy(\.properties.italic))
    }

    func testFormatTextRangeOnlyTouchesTargetSpan() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello world")
            ])
        ])
        var format = RunProperties()
        format.bold = true
        format.color = "FF0000"

        try document.formatTextRange(at: 0, start: 6, end: 11, format: format)

        let runs = document.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello ", "world"])
        XCTAssertFalse(runs[0].properties.bold)
        XCTAssertTrue(runs[1].properties.bold)
        XCTAssertEqual(runs[1].properties.color, "FF0000")
    }

    func testFormatTextRangeCanClearExistingHighlight() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello "),
                TestFixtures.makeRun("world", highlight: .yellow)
            ])
        ])
        var format = RunProperties()
        format.clearHighlight = true

        try document.formatTextRange(at: 0, start: 6, end: 11, format: format)

        let runs = document.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["Hello world"])
        XCTAssertNil(runs[0].properties.highlight)
    }

    func testReplaceTextRangeSupportsUnicodeCharacterOffsets() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Cafe\u{301} noir")
            ])
        ])

        try document.replaceTextRange(at: 0, start: 0, end: 4, replacement: "Bistro", replacementProperties: nil)

        XCTAssertEqual(document.getParagraphs()[0].getText(), "Bistro noir")
    }

    func testReplaceTextRangeRejectsParagraphsContainingHyperlinks() throws {
        let hyperlink = Hyperlink.external(
            id: "link-1",
            text: "docs",
            url: "https://example.com",
            relationshipId: "rId1"
        )
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(
                runs: [TestFixtures.makeRun("See ")],
                hyperlinks: [hyperlink]
            )
        ])

        XCTAssertThrowsError(
            try document.replaceTextRange(at: 0, start: 0, end: 0, replacement: "Updated ", replacementProperties: nil)
        ) { error in
            XCTAssertEqual(error as? InlineEditingError, .unsupportedParagraphContent("paragraph contains hyperlinks"))
        }
    }

    func testReplaceTextRangeRejectsUnsupportedRawXMLRunSplits() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("prefix "),
                TestFixtures.makeRun("field", rawXML: "<w:fldSimple/>"),
                TestFixtures.makeRun(" suffix")
            ])
        ])

        XCTAssertThrowsError(
            try document.replaceTextRange(at: 0, start: 8, end: 10, replacement: "xx", replacementProperties: nil)
        ) { error in
            XCTAssertEqual(error as? InlineEditingError, .unsupportedParagraphContent("edit would split a non-text run"))
        }
    }

    func testFormatTextRangeRejectsInvalidOffsets() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [TestFixtures.makeRun("Hello")])
        ])
        let format = RunProperties()

        XCTAssertThrowsError(
            try document.formatTextRange(at: 0, start: -1, end: 2, format: format)
        ) { error in
            XCTAssertEqual(error as? InlineEditingError, .invalidRange(start: -1, end: 2, length: 5))
        }
    }

    func testReplaceTextRangeMergesAdjacentRunsAfterDelete() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("ab", bold: true),
                TestFixtures.makeRun("cd", bold: true),
                TestFixtures.makeRun("ef")
            ])
        ])

        try document.replaceTextRange(at: 0, start: 2, end: 4, replacement: "", replacementProperties: nil)

        let runs = document.getParagraphs()[0].runs
        XCTAssertEqual(runs.map(\.text), ["ab", "ef"])
        XCTAssertTrue(runs[0].properties.bold)
    }

    func testTrackedReplaceTextRangeKeepsVisibleRunsInSyncForFollowUpEdits() throws {
        var document = TestFixtures.makeDocument(paragraphs: [
            TestFixtures.makeParagraph(runs: [
                TestFixtures.makeRun("Hello world")
            ])
        ])
        document.enableTrackChanges(author: "Test")

        try document.replaceTextRange(at: 0, start: 6, end: 11, replacement: "teammates", replacementProperties: nil)

        XCTAssertEqual(document.getParagraphs()[0].getText(), "Hello teammates")
        XCTAssertEqual(document.getParagraphs()[0].runs.map(\.text).joined(), "Hello teammates")

        try document.replaceTextRange(at: 0, start: 6, end: 15, replacement: "crew", replacementProperties: nil)

        XCTAssertEqual(document.getParagraphs()[0].getText(), "Hello crew")
        XCTAssertEqual(document.getParagraphs()[0].runs.map(\.text).joined(), "Hello crew")
    }
}
