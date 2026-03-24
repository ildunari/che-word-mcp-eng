import XCTest
import OOXMLSwift
import DocConverterSwift
@testable import WordToMDSwift

/// Bijection 驗證：不同的 Word 文件必須產生不同的 Tier 3 輸出。
///
/// 核心邏輯（docs/lossless-conversion.md §7.2）：
/// ```
/// ∀ w₁, w₂ ∈ W_test:
///     if w₁ ≠ w₂:
///         assert: convert₃(w₁) ≠ convert₃(w₂)
/// ```
///
/// 策略：建立成對的文件，僅差一個 Layer C 屬性。
/// Tier 1 (markdown) 應該相同，但 Tier 3 (markdown + metadata) 必須不同。
final class BijectionTests: XCTestCase {
    let converter = WordConverter()

    // MARK: - Helpers

    /// Tier 3 轉換，回傳 (markdown, metadataYAML)
    private func convertTier3(_ doc: WordDocument) throws -> (markdown: String, metadata: String) {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("bijection-test-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)

        defer { try? FileManager.default.removeItem(at: tempDir) }

        let metaURL = tempDir.appendingPathComponent("doc.meta.yaml")

        var options = ConversionOptions.default
        options.fidelity = .marker
        options.metadataOutput = metaURL

        let markdown = try converter.convertToString(document: doc, options: options)
        let metadata = try String(contentsOf: metaURL, encoding: .utf8)

        return (markdown, metadata)
    }

    // MARK: - Color Difference

    /// 紅色粗體 vs 藍色粗體：Markdown 相同（都是 **text**），Metadata 必須不同
    func testColorDifferenceCapturedInMetadata() throws {
        // Doc A: 紅色文字
        var docA = WordDocument()
        let runA = Run(text: "important", properties: RunProperties(bold: true, color: "FF0000"))
        docA.appendParagraph(Paragraph(runs: [runA]))

        // Doc B: 藍色文字
        var docB = WordDocument()
        let runB = Run(text: "important", properties: RunProperties(bold: true, color: "0000FF"))
        docB.appendParagraph(Paragraph(runs: [runB]))

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        // Tier 1 (markdown) 應該相同
        XCTAssertEqual(resultA.markdown, resultB.markdown,
                       "Same bold text should produce same Markdown regardless of color")

        // Tier 3 (metadata) 必須不同
        XCTAssertNotEqual(resultA.metadata, resultB.metadata,
                          "Different colors must produce different metadata")

        // 驗證 metadata 包含各自的顏色值
        XCTAssertTrue(resultA.metadata.contains("FF0000"), "Metadata A should contain red color")
        XCTAssertTrue(resultB.metadata.contains("0000FF"), "Metadata B should contain blue color")
    }

    // MARK: - Font Difference

    /// 不同字體的文字：Markdown 相同，Metadata 不同
    func testFontDifferenceCapturedInMetadata() throws {
        var docA = WordDocument()
        let runA = Run(text: "text", properties: RunProperties(fontName: "Times New Roman"))
        docA.appendParagraph(Paragraph(runs: [runA]))

        var docB = WordDocument()
        let runB = Run(text: "text", properties: RunProperties(fontName: "Arial"))
        docB.appendParagraph(Paragraph(runs: [runB]))

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        XCTAssertEqual(resultA.markdown, resultB.markdown)
        XCTAssertNotEqual(resultA.metadata, resultB.metadata)
        XCTAssertTrue(resultA.metadata.contains("Times New Roman"))
        XCTAssertTrue(resultB.metadata.contains("Arial"))
    }

    // MARK: - Font Size Difference

    func testFontSizeDifferenceCapturedInMetadata() throws {
        var docA = WordDocument()
        let runA = Run(text: "text", properties: RunProperties(fontSize: 24))  // 12pt
        docA.appendParagraph(Paragraph(runs: [runA]))

        var docB = WordDocument()
        let runB = Run(text: "text", properties: RunProperties(fontSize: 48))  // 24pt
        docB.appendParagraph(Paragraph(runs: [runB]))

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        XCTAssertEqual(resultA.markdown, resultB.markdown)
        XCTAssertNotEqual(resultA.metadata, resultB.metadata)
    }

    // MARK: - Alignment Difference

    /// 置中 vs 靠右：Markdown 相同（段落文字一樣），Metadata 不同
    func testAlignmentDifferenceCapturedInMetadata() throws {
        var docA = WordDocument()
        var paraA = Paragraph(text: "aligned text")
        paraA.properties.alignment = .center
        docA.appendParagraph(paraA)

        var docB = WordDocument()
        var paraB = Paragraph(text: "aligned text")
        paraB.properties.alignment = .right
        docB.appendParagraph(paraB)

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        XCTAssertEqual(resultA.markdown, resultB.markdown)
        XCTAssertNotEqual(resultA.metadata, resultB.metadata)
        XCTAssertTrue(resultA.metadata.contains("center"))
        XCTAssertTrue(resultB.metadata.contains("right"))
    }

    // MARK: - Spacing Difference

    func testSpacingDifferenceCapturedInMetadata() throws {
        var docA = WordDocument()
        var paraA = Paragraph(text: "spaced")
        paraA.properties.spacing = Spacing(before: 240, after: 120)
        docA.appendParagraph(paraA)

        var docB = WordDocument()
        var paraB = Paragraph(text: "spaced")
        paraB.properties.spacing = Spacing(before: 480, after: 240)
        docB.appendParagraph(paraB)

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        XCTAssertEqual(resultA.markdown, resultB.markdown)
        XCTAssertNotEqual(resultA.metadata, resultB.metadata)
    }

    // MARK: - Highlight Color Difference (Layer B off)

    /// 黃色螢光筆 vs 綠色螢光筆：不啟用 HTML extensions 時 Markdown 相同
    func testHighlightColorDifferenceWithoutHTMLExtensions() throws {
        var docA = WordDocument()
        let runA = Run(text: "highlighted", properties: RunProperties(highlight: .yellow))
        docA.appendParagraph(Paragraph(runs: [runA]))

        var docB = WordDocument()
        let runB = Run(text: "highlighted", properties: RunProperties(highlight: .green))
        docB.appendParagraph(Paragraph(runs: [runB]))

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        // 不啟用 HTML extensions，Markdown 相同（都只是 "highlighted"）
        XCTAssertEqual(resultA.markdown, resultB.markdown)
        // 但 metadata 記錄了不同的螢光筆顏色
        XCTAssertNotEqual(resultA.metadata, resultB.metadata)
        XCTAssertTrue(resultA.metadata.contains("yellow"))
        XCTAssertTrue(resultB.metadata.contains("green"))
    }

    // MARK: - Underline Type Difference

    /// 單底線 vs 雙底線：即使啟用 HTML extensions，<u> 無法區分子類型
    func testUnderlineTypeDifferenceCapturedInMetadata() throws {
        var docA = WordDocument()
        let runA = Run(text: "underlined", properties: RunProperties(underline: .single))
        docA.appendParagraph(Paragraph(runs: [runA]))

        var docB = WordDocument()
        let runB = Run(text: "underlined", properties: RunProperties(underline: .double))
        docB.appendParagraph(Paragraph(runs: [runB]))

        let resultA = try convertTier3(docA)
        let resultB = try convertTier3(docB)

        // Markdown 可能相同（都是普通文字，或都是 <u>）
        // 但 metadata 必須區分 single vs double
        XCTAssertNotEqual(resultA.metadata, resultB.metadata)
        XCTAssertTrue(resultA.metadata.contains("single"))
        XCTAssertTrue(resultB.metadata.contains("double"))
    }
}
