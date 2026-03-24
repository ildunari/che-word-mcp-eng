import Foundation
import OOXMLSwift

/// Metadata 收集器（Tier 3 / Marker）
///
/// 在 streaming 過程中收集所有 Layer C 資訊——
/// 這些是 Markdown 無法表達的屬性（Sparse Metadata 原則）。
struct MetadataCollector {
    var version = "1.0"
    var sourceFormat = "docx"
    var sourceFile: String?

    // Document-level metadata
    var documentProperties: DocumentProperties?
    var sectionProperties: SectionProperties?
    var styles: [StyleMeta] = []
    var comments: [CommentMeta] = []
    var numberingDefs: [NumberingDefMeta] = []

    // Per-paragraph metadata (sparse: 只記錄有非預設屬性的段落)
    var paragraphMeta: [ParagraphMeta] = []

    // Per-table metadata (sparse)
    var tableMeta: [TableMeta] = []

    // Per-figure metadata
    var figureMeta: [FigureMeta] = []

    // MARK: - Collection

    /// 收集文件級 metadata
    mutating func collectDocument(_ doc: WordDocument) {
        self.documentProperties = doc.properties
        self.sectionProperties = doc.sectionProperties
        self.sourceFile = doc.properties.title

        // 收集 style 定義
        for style in doc.styles {
            styles.append(StyleMeta(
                id: style.id,
                name: style.name,
                basedOn: style.basedOn
            ))
        }

        // 收集 comment 內容
        for comment in doc.comments.comments {
            comments.append(CommentMeta(
                id: comment.id,
                author: comment.author,
                text: comment.text,
                paragraphIndex: comment.paragraphIndex,
                parentId: comment.parentId,
                done: comment.done
            ))
        }

        // 收集 numbering definitions
        for abstractNum in doc.numbering.abstractNums {
            var levels: [NumberingLevelMeta] = []
            for level in abstractNum.levels {
                levels.append(NumberingLevelMeta(
                    ilvl: level.ilvl,
                    numFmt: level.numFmt.rawValue,
                    lvlText: level.lvlText,
                    start: level.start,
                    indent: level.indent,
                    fontName: level.fontName
                ))
            }
            numberingDefs.append(NumberingDefMeta(
                abstractNumId: abstractNum.abstractNumId,
                levels: levels
            ))
        }
    }

    /// 收集元素級 metadata
    mutating func collectElement(_ child: BodyChild, index: Int) {
        switch child {
        case .paragraph(let para):
            collectParagraph(para, index: index)
        case .table(let table):
            collectTable(table, index: index)
        }
    }

    /// 收集段落級 metadata（只記錄 Markdown 無法表達的屬性）
    private mutating func collectParagraph(_ para: Paragraph, index: Int) {
        let props = para.properties

        // 檢查是否有需要記錄的 Layer C 屬性
        let hasAlignment = props.alignment != nil
        let hasSpacing = props.spacing != nil
        let hasIndentation = props.indentation != nil
        let hasComments = !para.commentIds.isEmpty
        let hasBookmarks = !para.bookmarks.isEmpty
        let hasKeepNext = props.keepNext
        let hasKeepLines = props.keepLines
        let hasPageBreakBefore = props.pageBreakBefore
        let hasBorder = props.border != nil
        let hasShading = props.shading != nil
        let hasRunMeta = para.runs.contains { hasLayerCProperties($0.properties) }

        guard hasAlignment || hasSpacing || hasIndentation || hasComments || hasBookmarks ||
              hasKeepNext || hasKeepLines || hasPageBreakBefore || hasBorder || hasShading ||
              hasRunMeta else {
            return
        }

        var meta = ParagraphMeta(index: index)
        meta.alignment = props.alignment?.rawValue
        meta.spacing = props.spacing.map { SpacingMeta(before: $0.before, after: $0.after, line: $0.line) }
        meta.indentation = props.indentation.map { IndentationMeta(left: $0.left, right: $0.right, firstLine: $0.firstLine, hanging: $0.hanging) }
        meta.commentIds = para.commentIds.isEmpty ? nil : para.commentIds
        meta.bookmarkNames = para.bookmarks.isEmpty ? nil : para.bookmarks.map { $0.name }
        meta.keepNext = props.keepNext ? true : nil
        meta.keepLines = props.keepLines ? true : nil
        meta.pageBreakBefore = props.pageBreakBefore ? true : nil

        if let border = props.border {
            meta.border = ParagraphBorderMeta(
                top: border.top.map { BorderStyleMeta(type: $0.type.rawValue, color: $0.color, size: $0.size) },
                bottom: border.bottom.map { BorderStyleMeta(type: $0.type.rawValue, color: $0.color, size: $0.size) },
                left: border.left.map { BorderStyleMeta(type: $0.type.rawValue, color: $0.color, size: $0.size) },
                right: border.right.map { BorderStyleMeta(type: $0.type.rawValue, color: $0.color, size: $0.size) }
            )
        }

        if let shading = props.shading {
            meta.shading = ShadingMeta(fill: shading.fill, pattern: shading.pattern?.rawValue)
        }

        // Run-level Layer C 屬性
        var offset = 0
        for run in para.runs {
            let length = run.text.count
            if hasLayerCProperties(run.properties) {
                var runMeta = RunMeta(range: [offset, offset + length])
                runMeta.fontName = run.properties.fontName
                runMeta.fontSize = run.properties.fontSize
                runMeta.color = run.properties.color
                runMeta.highlightColor = run.properties.highlight?.rawValue
                runMeta.underlineType = run.properties.underline?.rawValue
                if let cs = run.properties.characterSpacing {
                    runMeta.characterSpacing = CharacterSpacingMeta(
                        spacing: cs.spacing,
                        position: cs.position,
                        kern: cs.kern
                    )
                }
                meta.runs.append(runMeta)
            }
            offset += length
        }

        paragraphMeta.append(meta)
    }

    /// 收集 table metadata
    private mutating func collectTable(_ table: Table, index: Int) {
        let props = table.properties
        let hasWidth = props.width != nil
        let hasAlignment = props.alignment != nil
        let hasBorders = props.borders != nil
        let hasLayout = props.layout != nil

        guard hasWidth || hasAlignment || hasBorders || hasLayout else { return }

        var meta = TableMeta(index: index)
        meta.width = props.width
        meta.widthType = props.widthType?.rawValue
        meta.alignment = props.alignment?.rawValue
        meta.layout = props.layout?.rawValue

        // Collect row/cell info
        for (rowIdx, row) in table.rows.enumerated() {
            if row.properties.isHeader || row.properties.height != nil {
                var rowMeta = TableRowMeta(rowIndex: rowIdx)
                rowMeta.isHeader = row.properties.isHeader ? true : nil
                rowMeta.height = row.properties.height
                meta.rows.append(rowMeta)
            }
        }

        tableMeta.append(meta)
    }

    /// 收集 figure metadata
    mutating func collectFigure(_ drawing: Drawing, imageRef: ImageReference, path: String) {
        figureMeta.append(FigureMeta(
            id: imageRef.id,
            file: path,
            contentType: imageRef.contentType,
            placement: drawing.type == .inline ? "inline" : "anchor",
            width: drawing.width,
            height: drawing.height,
            altText: drawing.description.isEmpty ? nil : drawing.description,
            name: drawing.name
        ))
    }

    /// 檢查 run 是否有 Layer C 屬性（Markdown 無法表達的）
    private func hasLayerCProperties(_ props: RunProperties) -> Bool {
        props.fontName != nil ||
        props.fontSize != nil ||
        props.color != nil ||
        props.highlight != nil ||
        props.underline != nil ||
        props.characterSpacing != nil
    }

    // MARK: - YAML Output

    /// 手動生成 YAML（不引入外部 library）
    func writeYAML(to url: URL) throws {
        var lines: [String] = []

        lines.append("version: \"\(version)\"")
        lines.append("source:")
        lines.append("  format: \"\(sourceFormat)\"")
        if let file = sourceFile {
            lines.append("  file: \"\(escapeYAML(file))\"")
        }

        // Document properties
        if let props = documentProperties {
            lines.append("")
            lines.append("document:")
            lines.append("  properties:")
            if let title = props.title { lines.append("    title: \"\(escapeYAML(title))\"") }
            if let creator = props.creator { lines.append("    creator: \"\(escapeYAML(creator))\"") }
            if let subject = props.subject { lines.append("    subject: \"\(escapeYAML(subject))\"") }
            if let description = props.description { lines.append("    description: \"\(escapeYAML(description))\"") }
            if let keywords = props.keywords { lines.append("    keywords: \"\(escapeYAML(keywords))\"") }
            if let created = props.created {
                lines.append("    created: \"\(ISO8601DateFormatter().string(from: created))\"")
            }
            if let modified = props.modified {
                lines.append("    modified: \"\(ISO8601DateFormatter().string(from: modified))\"")
            }
        }

        // Styles
        if !styles.isEmpty {
            lines.append("")
            lines.append("  styles:")
            for style in styles {
                lines.append("    - id: \"\(escapeYAML(style.id))\"")
                lines.append("      name: \"\(escapeYAML(style.name))\"")
                if let basedOn = style.basedOn {
                    lines.append("      basedOn: \"\(escapeYAML(basedOn))\"")
                }
            }
        }

        // Section properties
        if let section = sectionProperties {
            lines.append("")
            lines.append("  sections:")
            lines.append("    - pageSize:")
            lines.append("        width: \(section.pageSize.width)")
            lines.append("        height: \(section.pageSize.height)")
            lines.append("      orientation: \(section.orientation == .landscape ? "landscape" : "portrait")")
            lines.append("      margins:")
            lines.append("        top: \(section.pageMargins.top)")
            lines.append("        bottom: \(section.pageMargins.bottom)")
            lines.append("        left: \(section.pageMargins.left)")
            lines.append("        right: \(section.pageMargins.right)")
        }

        // Comments
        if !comments.isEmpty {
            lines.append("")
            lines.append("  comments:")
            for comment in comments {
                lines.append("    - id: \(comment.id)")
                lines.append("      author: \"\(escapeYAML(comment.author))\"")
                lines.append("      text: \"\(escapeYAML(comment.text))\"")
                lines.append("      paragraphIndex: \(comment.paragraphIndex)")
                if let parentId = comment.parentId {
                    lines.append("      parentId: \(parentId)")
                }
                if comment.done {
                    lines.append("      done: true")
                }
            }
        }

        // Numbering definitions
        if !numberingDefs.isEmpty {
            lines.append("")
            lines.append("  numbering:")
            for def in numberingDefs {
                lines.append("    - abstractNumId: \(def.abstractNumId)")
                lines.append("      levels:")
                for level in def.levels {
                    lines.append("        - ilvl: \(level.ilvl)")
                    lines.append("          numFmt: \(level.numFmt)")
                    lines.append("          lvlText: \"\(escapeYAML(level.lvlText))\"")
                    lines.append("          start: \(level.start)")
                    lines.append("          indent: \(level.indent)")
                    if let fontName = level.fontName {
                        lines.append("          fontName: \"\(escapeYAML(fontName))\"")
                    }
                }
            }
        }

        // Paragraphs (sparse)
        if !paragraphMeta.isEmpty {
            lines.append("")
            lines.append("paragraphs:")
            for para in paragraphMeta {
                lines.append("  - index: \(para.index)")
                if let alignment = para.alignment {
                    lines.append("    alignment: \(alignment)")
                }
                if let spacing = para.spacing {
                    var spacingParts: [String] = []
                    if let before = spacing.before { spacingParts.append("before: \(before)") }
                    if let after = spacing.after { spacingParts.append("after: \(after)") }
                    if let line = spacing.line { spacingParts.append("line: \(line)") }
                    lines.append("    spacing: { \(spacingParts.joined(separator: ", ")) }")
                }
                if let indent = para.indentation {
                    var parts: [String] = []
                    if let left = indent.left { parts.append("left: \(left)") }
                    if let right = indent.right { parts.append("right: \(right)") }
                    if let firstLine = indent.firstLine { parts.append("firstLine: \(firstLine)") }
                    if let hanging = indent.hanging { parts.append("hanging: \(hanging)") }
                    lines.append("    indentation: { \(parts.joined(separator: ", ")) }")
                }
                if let commentIds = para.commentIds {
                    lines.append("    commentIds: [\(commentIds.map { String($0) }.joined(separator: ", "))]")
                }
                if para.keepNext == true { lines.append("    keepNext: true") }
                if para.keepLines == true { lines.append("    keepLines: true") }
                if para.pageBreakBefore == true { lines.append("    pageBreakBefore: true") }
                if let border = para.border {
                    lines.append("    border:")
                    if let top = border.top { lines.append("      top: { type: \(top.type), color: \"\(top.color)\", size: \(top.size) }") }
                    if let bottom = border.bottom { lines.append("      bottom: { type: \(bottom.type), color: \"\(bottom.color)\", size: \(bottom.size) }") }
                    if let left = border.left { lines.append("      left: { type: \(left.type), color: \"\(left.color)\", size: \(left.size) }") }
                    if let right = border.right { lines.append("      right: { type: \(right.type), color: \"\(right.color)\", size: \(right.size) }") }
                }
                if let shading = para.shading {
                    var shadingParts = ["fill: \"\(shading.fill)\""]
                    if let pattern = shading.pattern { shadingParts.append("pattern: \(pattern)") }
                    lines.append("    shading: { \(shadingParts.joined(separator: ", ")) }")
                }
                if !para.runs.isEmpty {
                    lines.append("    runs:")
                    for run in para.runs {
                        lines.append("      - range: [\(run.range[0]), \(run.range[1])]")
                        if let font = run.fontName { lines.append("        font: \"\(escapeYAML(font))\"") }
                        if let size = run.fontSize { lines.append("        fontSize: \(size)") }
                        if let color = run.color { lines.append("        color: \"#\(color)\"") }
                        if let highlight = run.highlightColor { lines.append("        highlight: \(highlight)") }
                        if let underline = run.underlineType { lines.append("        underline: \(underline)") }
                        if let cs = run.characterSpacing {
                            var csParts: [String] = []
                            if let s = cs.spacing { csParts.append("spacing: \(s)") }
                            if let p = cs.position { csParts.append("position: \(p)") }
                            if let k = cs.kern { csParts.append("kern: \(k)") }
                            lines.append("        characterSpacing: { \(csParts.joined(separator: ", ")) }")
                        }
                    }
                }
            }
        }

        // Tables (sparse)
        if !tableMeta.isEmpty {
            lines.append("")
            lines.append("tables:")
            for table in tableMeta {
                lines.append("  - index: \(table.index)")
                if let width = table.width { lines.append("    width: \(width)") }
                if let widthType = table.widthType { lines.append("    widthType: \(widthType)") }
                if let alignment = table.alignment { lines.append("    alignment: \(alignment)") }
                if let layout = table.layout { lines.append("    layout: \(layout)") }
                if !table.rows.isEmpty {
                    lines.append("    rows:")
                    for row in table.rows {
                        lines.append("      - rowIndex: \(row.rowIndex)")
                        if row.isHeader == true { lines.append("        isHeader: true") }
                        if let height = row.height { lines.append("        height: \(height)") }
                    }
                }
            }
        }

        // Figures
        if !figureMeta.isEmpty {
            lines.append("")
            lines.append("figures:")
            for fig in figureMeta {
                lines.append("  - id: \"\(fig.id)\"")
                lines.append("    file: \"\(fig.file)\"")
                lines.append("    contentType: \"\(fig.contentType)\"")
                lines.append("    placement: \(fig.placement)")
                lines.append("    width: \(fig.width)")
                lines.append("    height: \(fig.height)")
                if let alt = fig.altText { lines.append("    altText: \"\(escapeYAML(alt))\"") }
            }
        }

        let yaml = lines.joined(separator: "\n") + "\n"
        try yaml.write(to: url, atomically: true, encoding: .utf8)
    }

    private func escapeYAML(_ text: String) -> String {
        text
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "\"", with: "\\\"")
    }
}

// MARK: - Metadata Models

struct StyleMeta {
    let id: String
    let name: String
    let basedOn: String?
}

struct CommentMeta {
    let id: Int
    let author: String
    let text: String
    let paragraphIndex: Int
    let parentId: Int?
    let done: Bool
}

struct NumberingDefMeta {
    let abstractNumId: Int
    let levels: [NumberingLevelMeta]
}

struct NumberingLevelMeta {
    let ilvl: Int
    let numFmt: String
    let lvlText: String
    let start: Int
    let indent: Int
    let fontName: String?
}

struct ParagraphMeta {
    let index: Int
    var alignment: String?
    var spacing: SpacingMeta?
    var indentation: IndentationMeta?
    var commentIds: [Int]?
    var bookmarkNames: [String]?
    var keepNext: Bool?
    var keepLines: Bool?
    var pageBreakBefore: Bool?
    var border: ParagraphBorderMeta?
    var shading: ShadingMeta?
    var runs: [RunMeta] = []
}

struct ParagraphBorderMeta {
    let top: BorderStyleMeta?
    let bottom: BorderStyleMeta?
    let left: BorderStyleMeta?
    let right: BorderStyleMeta?
}

struct BorderStyleMeta {
    let type: String
    let color: String
    let size: Int
}

struct ShadingMeta {
    let fill: String
    let pattern: String?
}

struct SpacingMeta {
    let before: Int?
    let after: Int?
    let line: Int?
}

struct IndentationMeta {
    let left: Int?
    let right: Int?
    let firstLine: Int?
    let hanging: Int?
}

struct RunMeta {
    let range: [Int]  // [start, end]
    var fontName: String?
    var fontSize: Int?
    var color: String?
    var highlightColor: String?
    var underlineType: String?
    var characterSpacing: CharacterSpacingMeta?
}

struct CharacterSpacingMeta {
    let spacing: Int?
    let position: Int?
    let kern: Int?
}

struct TableMeta {
    let index: Int
    var width: Int?
    var widthType: String?
    var alignment: String?
    var layout: String?
    var rows: [TableRowMeta] = []
}

struct TableRowMeta {
    let rowIndex: Int
    var isHeader: Bool?
    var height: Int?
}

struct FigureMeta {
    let id: String
    let file: String
    let contentType: String
    let placement: String
    let width: Int
    let height: Int
    let altText: String?
    let name: String
}
