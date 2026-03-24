import Foundation
import DocConverterSwift
import OOXMLSwift
import MarkdownSwift

/// Word 轉 Markdown 轉換器
///
/// 遵循資訊下沉原則（docs/lossless-conversion.md §3.4）：
/// 每個元素在能被表達的最低 Tier 中表達。
public struct WordConverter: DocumentConverter {
    public static let sourceFormat = "docx"

    public init() {}

    public func convert<W: DocConverterSwift.StreamingOutput>(
        input: URL,
        output: inout W,
        options: ConversionOptions
    ) throws {
        let document = try DocxReader.read(from: input)
        try convert(document: document, output: &output, options: options)
    }

    /// 直接從 WordDocument 轉換（供 MCP 等已載入文件的場景使用）
    public func convert<W: DocConverterSwift.StreamingOutput>(
        document: WordDocument,
        output: inout W,
        options: ConversionOptions = .default
    ) throws {
        // 建立轉換上下文（收集 footnote definitions 等跨段落資訊）
        var context = ConversionContext(document: document, options: options)

        // Tier 3: 收集文件級 metadata
        if options.fidelity == .marker {
            context.metadataCollector = MetadataCollector()
            context.metadataCollector?.collectDocument(document)
        }

        // Tier 2+: 初始化圖片提取器
        if options.fidelity >= .markdownWithFigures, let figDir = options.figuresDirectory {
            context.figureExtractor = FigureExtractor(directory: figDir)
            try context.figureExtractor?.createDirectory()
        }

        // Practical Mode: heading heuristic pre-scan
        if options.headingHeuristic {
            var heuristic = HeadingHeuristic()
            heuristic.analyze(children: document.body.children, styles: document.styles)
            context.headingHeuristic = heuristic
        }

        if options.includeFrontmatter {
            try writeFrontmatter(document: document, output: &output)
        }

        for (index, child) in document.body.children.enumerated() {
            switch child {
            case .paragraph(let paragraph):
                try processParagraph(
                    paragraph,
                    context: &context,
                    output: &output
                )
            case .table(let table):
                try processTable(table, context: &context, output: &output)
            }

            // Tier 3: 收集元素級 metadata
            if options.fidelity == .marker {
                context.metadataCollector?.collectElement(child, index: index)
            }
        }

        // 輸出 footnote definitions（資訊下沉：footnotes 屬於 Tier 1）
        try emitFootnoteDefinitions(context: context, output: &output)

        // Tier 3: 寫出 metadata sidecar
        if options.fidelity == .marker, let metaURL = options.metadataOutput {
            try context.metadataCollector?.writeYAML(to: metaURL)
        }
    }

    /// 從 WordDocument 直接轉為字串
    public func convertToString(
        document: WordDocument,
        options: ConversionOptions = .default
    ) throws -> String {
        var writer = DocConverterSwift.StringOutput()
        try convert(document: document, output: &writer, options: options)
        return writer.content
    }

    // MARK: - Frontmatter

    private func writeFrontmatter<W: DocConverterSwift.StreamingOutput>(
        document: WordDocument,
        output: inout W
    ) throws {
        try output.writeLine("---")

        let props = document.properties
        if let title = props.title, !title.isEmpty {
            try output.writeLine("title: \"\(escapeYAML(title))\"")
        }
        if let author = props.creator, !author.isEmpty {
            try output.writeLine("author: \"\(escapeYAML(author))\"")
        }
        if let subject = props.subject, !subject.isEmpty {
            try output.writeLine("subject: \"\(escapeYAML(subject))\"")
        }

        try output.writeLine("---")
        try output.writeBlankLine()
    }

    // MARK: - Paragraph Processing

    private func processParagraph<W: DocConverterSwift.StreamingOutput>(
        _ paragraph: Paragraph,
        context: inout ConversionContext,
        output: inout W
    ) throws {
        // Horizontal rule: page break → ---（tier_min = 1）
        if paragraph.hasPageBreak || paragraph.properties.pageBreakBefore {
            try output.writeLine("---")
            try output.writeBlankLine()
            // 如果段落只有 page break 沒有文字，到此結束
            if paragraph.runs.isEmpty && paragraph.hyperlinks.isEmpty {
                return
            }
        }

        // Code block 偵測（tier_min = 1）
        if let styleName = paragraph.properties.style,
           isCodeStyle(styleName, styles: context.styles) {
            let rawText = collectPlainText(paragraph)
            if !rawText.isEmpty {
                // 收集連續的 code style 段落，累積成一個 code block
                // 目前先簡單輸出單行 code block（用 fenced）
                try output.writeLine("```")
                try output.writeLine(rawText)
                try output.writeLine("```")
                try output.writeBlankLine()
            }
            return
        }

        // Blockquote 偵測（tier_min = 1）
        if let styleName = paragraph.properties.style,
           isBlockquoteStyle(styleName, styles: context.styles) {
            let text = formatParagraphContent(paragraph, context: &context)
            if !text.trimmingCharacters(in: .whitespaces).isEmpty {
                try output.writeLine("> \(text)")
                try output.writeBlankLine()
            }
            return
        }

        let text = formatParagraphContent(paragraph, context: &context)

        // 空段落 → 跳過
        if text.trimmingCharacters(in: .whitespaces).isEmpty {
            return
        }

        // Heading 偵測（style-based 優先）
        if let styleName = paragraph.properties.style,
           let headingLevel = detectHeadingLevel(styleName: styleName, styles: context.styles) {
            let prefix = String(repeating: "#", count: headingLevel)
            try output.writeLine("\(prefix) \(text)")
            try output.writeBlankLine()
            return
        }
        // Heading heuristic fallback（Practical Mode）
        else if let headingLevel = context.headingHeuristic?.inferLevel(for: paragraph) {
            let prefix = String(repeating: "#", count: headingLevel)
            try output.writeLine("\(prefix) \(text)")
            try output.writeBlankLine()
            return
        }

        // List 偵測
        if let numInfo = paragraph.properties.numbering {
            let isBullet = isListBullet(
                numId: numInfo.numId,
                level: numInfo.level,
                numbering: context.numbering
            )
            let prefix = isBullet ? "- " : "1. "
            let indent = String(repeating: "  ", count: numInfo.level)
            try output.writeLine("\(indent)\(prefix)\(text)")
            return
        }

        // 一般段落
        try output.writeLine(text)
        try output.writeBlankLine()
    }

    // MARK: - Paragraph Content Formatting

    /// 格式化段落完整內容：runs + hyperlinks + footnote refs + images
    ///
    /// 資訊下沉原則：hyperlinks、footnotes、inline images 都在 Tier 1 表達。
    private func formatParagraphContent(
        _ paragraph: Paragraph,
        context: inout ConversionContext
    ) -> String {
        var result = ""

        // 1. 格式化 runs（含 inline images、code spans、Layer B）
        for run in paragraph.runs {
            result += formatRun(run, context: &context)
        }

        // 2. 格式化 hyperlinks（tier_min = 1）
        for hyperlink in paragraph.hyperlinks {
            result += formatHyperlink(hyperlink, context: context)
        }

        // 3. 插入 footnote references（tier_min = 1）
        for footnoteId in paragraph.footnoteIds {
            result += MarkdownInline.footnoteRef("\(footnoteId)")
            // 註冊到 context，稍後輸出 definition
            context.registerFootnote(id: footnoteId)
        }

        // 4. 插入 endnote references（合併為 footnote 語法，tier_min = 1）
        for endnoteId in paragraph.endnoteIds {
            let mappedId = "en\(endnoteId)"
            result += MarkdownInline.footnoteRef(mappedId)
            context.registerEndnote(id: endnoteId, mappedId: mappedId)
        }

        return result
    }

    // MARK: - Run Formatting

    /// 格式化單一 Run
    private func formatRun(
        _ run: Run,
        context: inout ConversionContext
    ) -> String {
        // Image（tier_min for reference = 1, tier_min for file = 2）
        if let drawing = run.drawing {
            return formatDrawing(drawing, context: &context)
        }

        var text = run.text

        // 跳過空文字
        if text.isEmpty { return "" }

        let props = run.properties
        let options = context.options

        // Inline code 偵測（tier_min = 1）：semantic type 為 codeBlock
        if let semantic = run.semantic, semantic.type == .codeBlock {
            return MarkdownInline.code(text)
        }

        // Layer A: bold / italic / strikethrough（tier_min = 1）
        if props.bold && props.italic {
            text = MarkdownInline.boldItalic(text)
        } else if props.bold {
            text = MarkdownInline.bold(text)
        } else if props.italic {
            text = MarkdownInline.italic(text)
        }

        if props.strikethrough {
            text = MarkdownInline.strikethrough(text)
        }

        // Layer B: HTML extensions（tier_min = 1 when enabled, otherwise 3）
        if options.useHTMLExtensions {
            if props.underline != nil {
                text = MarkdownInline.rawHTML("<u>\(text)</u>")
            }
            if props.verticalAlign == .superscript {
                text = MarkdownInline.rawHTML("<sup>\(text)</sup>")
            }
            if props.verticalAlign == .subscript {
                text = MarkdownInline.rawHTML("<sub>\(text)</sub>")
            }
            if props.highlight != nil {
                text = MarkdownInline.rawHTML("<mark>\(text)</mark>")
            }
        }

        return text
    }

    // MARK: - Hyperlink Formatting

    /// 格式化超連結（tier_min = 1）
    private func formatHyperlink(
        _ hyperlink: Hyperlink,
        context: ConversionContext
    ) -> String {
        let text = hyperlink.text

        switch hyperlink.type {
        case .external:
            // 外部連結：先查 hyperlinkReferences 取得 URL
            if let url = hyperlink.url, !url.isEmpty {
                return MarkdownInline.link(text, url: url)
            }
            // 透過 relationshipId 查詢
            if let rId = hyperlink.relationshipId,
               let ref = context.document.hyperlinkReferences.first(where: { $0.relationshipId == rId }) {
                return MarkdownInline.link(text, url: ref.url)
            }
            return text

        case .internal:
            // 內部連結（書籤）
            if let anchor = hyperlink.anchor {
                return MarkdownInline.link(text, url: "#\(anchor)")
            }
            return text
        }
    }

    // MARK: - Image / Drawing Formatting

    /// 格式化圖片（reference tier_min = 1, file tier_min = 2）
    private func formatDrawing(
        _ drawing: Drawing,
        context: inout ConversionContext
    ) -> String {
        let alt = drawing.description.isEmpty ? drawing.name : drawing.description

        // 查找對應的 ImageReference
        let imageRef = context.document.images.first { $0.id == drawing.imageId }

        // Tier 2+: 提取圖片檔案
        let preserveOriginal = context.options.preserveOriginalFormat
        if context.options.fidelity >= .markdownWithFigures,
           let imageRef = imageRef,
           context.figureExtractor != nil {
            if let relativePath = try? context.figureExtractor?.extract(
                imageRef,
                preserveOriginalFormat: preserveOriginal
            ) {
                // Tier 3: 收集 figure metadata
                if context.options.fidelity == .marker {
                    context.metadataCollector?.collectFigure(drawing, imageRef: imageRef, path: relativePath)
                }
                return MarkdownInline.image(alt, url: relativePath)
            }
        }

        // Tier 1: 僅放 reference（使用 imageId 作為 placeholder）
        let fileName = imageRef?.fileName ?? drawing.imageId
        return MarkdownInline.image(alt, url: fileName)
    }

    // MARK: - Footnote Definitions

    /// 在文件末尾輸出所有 footnote / endnote definitions
    private func emitFootnoteDefinitions<W: DocConverterSwift.StreamingOutput>(
        context: ConversionContext,
        output: inout W
    ) throws {
        let hasFootnotes = !context.referencedFootnoteIds.isEmpty
        let hasEndnotes = !context.referencedEndnoteIds.isEmpty

        guard hasFootnotes || hasEndnotes else { return }

        try output.writeBlankLine()

        // Footnotes
        for id in context.referencedFootnoteIds.sorted() {
            if let footnote = context.document.footnotes.footnotes.first(where: { $0.id == id }) {
                try output.writeLine("[^\(id)]: \(footnote.text)")
            }
        }

        // Endnotes（使用 mapped ID）
        for (id, mappedId) in context.endnoteIdMapping.sorted(by: { $0.key < $1.key }) {
            if let endnote = context.document.endnotes.endnotes.first(where: { $0.id == id }) {
                try output.writeLine("[^\(mappedId)]: \(endnote.text)")
            }
        }
    }

    // MARK: - Style Detection

    /// 偵測 code style（用於 inline code 和 code block）
    private func isCodeStyle(_ styleName: String, styles: [Style]) -> Bool {
        let lower = styleName.lowercased()
        let codePatterns = ["code", "source", "listing", "verbatim", "preformatted"]
        for pattern in codePatterns {
            if lower.contains(pattern) { return true }
        }
        // 檢查繼承鏈
        if let style = styles.first(where: { $0.id.lowercased() == lower }),
           let basedOn = style.basedOn {
            return isCodeStyle(basedOn, styles: styles)
        }
        return false
    }

    /// 偵測 blockquote style
    private func isBlockquoteStyle(_ styleName: String, styles: [Style]) -> Bool {
        let lower = styleName.lowercased()
        let quotePatterns = ["quote", "block text"]
        for pattern in quotePatterns {
            if lower.contains(pattern) { return true }
        }
        if let style = styles.first(where: { $0.id.lowercased() == lower }),
           let basedOn = style.basedOn {
            return isBlockquoteStyle(basedOn, styles: styles)
        }
        return false
    }

    // MARK: - List Detection

    private func isListBullet(numId: Int, level: Int, numbering: Numbering) -> Bool {
        guard let num = numbering.nums.first(where: { $0.numId == numId }) else {
            return true
        }
        guard let abstractNum = numbering.abstractNums.first(where: { $0.abstractNumId == num.abstractNumId }) else {
            return true
        }
        guard let levelDef = abstractNum.levels.first(where: { $0.ilvl == level }) else {
            return true
        }
        return levelDef.numFmt == .bullet
    }

    // MARK: - Heading Detection

    private func detectHeadingLevel(styleName: String, styles: [Style]) -> Int? {
        let lowerName = styleName.lowercased()

        let headingPatterns: [(String, Int)] = [
            ("heading1", 1), ("heading 1", 1), ("標題 1", 1), ("標題1", 1),
            ("heading2", 2), ("heading 2", 2), ("標題 2", 2), ("標題2", 2),
            ("heading3", 3), ("heading 3", 3), ("標題 3", 3), ("標題3", 3),
            ("heading4", 4), ("heading 4", 4), ("標題 4", 4), ("標題4", 4),
            ("heading5", 5), ("heading 5", 5), ("標題 5", 5), ("標題5", 5),
            ("heading6", 6), ("heading 6", 6), ("標題 6", 6), ("標題6", 6),
            ("title", 1), ("subtitle", 2),
        ]

        for (pattern, level) in headingPatterns {
            if lowerName == pattern {
                return level
            }
        }

        if let style = styles.first(where: { $0.id.lowercased() == lowerName }),
           let basedOn = style.basedOn {
            return detectHeadingLevel(styleName: basedOn, styles: styles)
        }

        return nil
    }

    // MARK: - Table Processing

    private func processTable<W: DocConverterSwift.StreamingOutput>(
        _ table: Table,
        context: inout ConversionContext,
        output: inout W
    ) throws {
        guard !table.rows.isEmpty else { return }

        let columnCount = table.rows.map { $0.cells.count }.max() ?? 0
        guard columnCount > 0 else { return }

        let normalizedRows = table.rows.map { row -> [String] in
            var cells = row.cells.map { cell -> String in
                let content = cell.paragraphs.map { para in
                    formatParagraphContent(para, context: &context)
                }.joined(separator: " ")
                return MarkdownEscaping.escape(content, context: .tableCell)
            }
            while cells.count < columnCount {
                cells.append("")
            }
            return cells
        }

        let headerRow = normalizedRows[0]
        try output.writeLine("| " + headerRow.joined(separator: " | ") + " |")

        let separator = Array(repeating: "---", count: columnCount)
        try output.writeLine("|" + separator.joined(separator: "|") + "|")

        for row in normalizedRows.dropFirst() {
            try output.writeLine("| " + row.joined(separator: " | ") + " |")
        }

        try output.writeBlankLine()
    }

    // MARK: - Helpers

    /// 收集段落純文字（不含格式，用於 code block）
    private func collectPlainText(_ paragraph: Paragraph) -> String {
        var text = paragraph.runs.map { $0.text }.joined()
        for hyperlink in paragraph.hyperlinks {
            text += hyperlink.text
        }
        return text
    }

    private func escapeYAML(_ text: String) -> String {
        text
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "\"", with: "\\\"")
    }
}

// MARK: - Conversion Context

/// 轉換上下文：在整個文件轉換過程中共享的狀態
struct ConversionContext {
    let document: WordDocument
    let options: ConversionOptions
    var styles: [Style] { document.styles }
    var numbering: Numbering { document.numbering }

    // Footnote tracking
    var referencedFootnoteIds: Set<Int> = []
    var referencedEndnoteIds: Set<Int> = []
    var endnoteIdMapping: [Int: String] = [:]  // endnoteId → mapped footnote id

    // Tier 2+
    var figureExtractor: FigureExtractor?

    // Practical Mode
    var headingHeuristic: HeadingHeuristic?

    // Tier 3
    var metadataCollector: MetadataCollector?

    init(document: WordDocument, options: ConversionOptions) {
        self.document = document
        self.options = options
    }

    mutating func registerFootnote(id: Int) {
        referencedFootnoteIds.insert(id)
    }

    mutating func registerEndnote(id: Int, mappedId: String) {
        referencedEndnoteIds.insert(id)
        endnoteIdMapping[id] = mappedId
    }
}
