import Foundation

// MARK: - Helper Functions

/// XML 特殊字元跳脫
private func escapeXML(_ string: String) -> String {
    return string
        .replacingOccurrences(of: "&", with: "&amp;")
        .replacingOccurrences(of: "<", with: "&lt;")
        .replacingOccurrences(of: ">", with: "&gt;")
        .replacingOccurrences(of: "\"", with: "&quot;")
        .replacingOccurrences(of: "'", with: "&apos;")
}

// MARK: - Revision Types

/// 修訂類型
public enum RevisionType: String, Codable {
    case insertion = "ins"              // 插入
    case deletion = "del"               // 刪除
    case moveFrom = "moveFrom"          // 移動來源
    case moveTo = "moveTo"              // 移動目標
    case formatChange = "rPrChange"     // Run 格式變更
    case formatting = "rPrChange2"      // 格式變更（舊版相容）
    case paragraphChange = "pPrChange"  // 段落屬性變更
}

/// 修訂記錄
public struct Revision: Equatable {
    public var id: Int                     // 修訂 ID
    public var type: RevisionType          // 修訂類型
    public var author: String              // 作者
    public var date: Date                  // 日期
    public var content: String?            // 修訂內容（用於插入/刪除）
    public var previousFormat: RunPropertiesSnapshot?  // 變更前格式（用於格式變更）

    // 舊版相容屬性
    public var paragraphIndex: Int = 0     // 段落索引
    public var originalText: String?       // 原始文字（刪除時）
    public var newText: String?            // 新文字（插入時）

    public init(id: Int, type: RevisionType, author: String, date: Date = Date(), content: String? = nil, previousFormat: RunPropertiesSnapshot? = nil) {
        self.id = id
        self.type = type
        self.author = author
        self.date = date
        self.content = content
        self.previousFormat = previousFormat
    }

    /// 舊版相容初始化器
    public init(id: Int, type: RevisionType, author: String, paragraphIndex: Int,
         originalText: String? = nil, newText: String? = nil, date: Date = Date()) {
        self.id = id
        self.type = type
        self.author = author
        self.paragraphIndex = paragraphIndex
        self.originalText = originalText
        self.newText = newText
        self.date = date
        self.content = newText ?? originalText
    }
}

// MARK: - Tracked Content

/// 追蹤修訂的 Run（可包含插入或刪除標記）
public struct TrackedRun: Equatable {
    public var run: Run                    // 原始 Run
    public var revision: Revision?         // 關聯的修訂（如果有）
    public var isDeleted: Bool             // 是否為刪除內容

    public init(run: Run, revision: Revision? = nil, isDeleted: Bool = false) {
        self.run = run
        self.revision = revision
        self.isDeleted = isDeleted
    }
}

/// 追蹤修訂的段落
public struct TrackedParagraph: Equatable {
    public var paragraph: Paragraph        // 原始段落
    public var trackedRuns: [TrackedRun]   // 帶修訂標記的 Runs
    public var paragraphRevision: Revision? // 段落級別的修訂（如整段插入/刪除）

    public init(paragraph: Paragraph, trackedRuns: [TrackedRun]? = nil, paragraphRevision: Revision? = nil) {
        self.paragraph = paragraph
        self.trackedRuns = trackedRuns ?? paragraph.runs.map { TrackedRun(run: $0) }
        self.paragraphRevision = paragraphRevision
    }
}

// MARK: - Revision Manager

/// 修訂管理器
public class RevisionManager {
    public var isTrackingEnabled: Bool = false
    public var currentAuthor: String = "Unknown"
    public var revisions: [Revision] = []
    private var nextRevisionId: Int = 0

    /// 建立新修訂
    func createRevision(type: RevisionType, content: String? = nil, previousFormat: RunPropertiesSnapshot? = nil) -> Revision {
        let revision = Revision(
            id: nextRevisionId,
            type: type,
            author: currentAuthor,
            date: Date(),
            content: content,
            previousFormat: previousFormat
        )
        nextRevisionId += 1
        revisions.append(revision)
        return revision
    }

    /// 接受修訂
    func acceptRevision(id: Int) -> Bool {
        guard let index = revisions.firstIndex(where: { $0.id == id }) else {
            return false
        }
        revisions.remove(at: index)
        return true
    }

    /// 拒絕修訂
    func rejectRevision(id: Int) -> Bool {
        guard let index = revisions.firstIndex(where: { $0.id == id }) else {
            return false
        }
        revisions.remove(at: index)
        return true
    }

    /// 取得所有修訂
    func getAllRevisions() -> [Revision] {
        return revisions
    }

    /// 依作者篩選修訂
    func getRevisions(byAuthor author: String) -> [Revision] {
        return revisions.filter { $0.author == author }
    }

    /// 依類型篩選修訂
    func getRevisions(byType type: RevisionType) -> [Revision] {
        return revisions.filter { $0.type == type }
    }
}

// MARK: - XML Generation

extension Revision {
    /// 產生修訂開始標籤
    func toOpeningXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        switch type {
        case .insertion:
            return "<w:ins w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\">"
        case .deletion:
            return "<w:del w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\">"
        case .moveFrom:
            return "<w:moveFrom w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\">"
        case .moveTo:
            return "<w:moveTo w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\">"
        case .formatChange, .formatting:
            return "<w:rPrChange w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\">"
        case .paragraphChange:
            return "<w:pPrChange w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\">"
        }
    }

    /// 產生修訂結束標籤
    func toClosingXML() -> String {
        switch type {
        case .insertion:
            return "</w:ins>"
        case .deletion:
            return "</w:del>"
        case .moveFrom:
            return "</w:moveFrom>"
        case .moveTo:
            return "</w:moveTo>"
        case .formatChange, .formatting:
            return "</w:rPrChange>"
        case .paragraphChange:
            return "</w:pPrChange>"
        }
    }
}

extension TrackedRun {
    /// 產生帶修訂標記的 Run XML
    func toXML() -> String {
        var xml = ""

        if let revision = revision {
            xml += revision.toOpeningXML()
        }

        if isDeleted {
            // 刪除的文字使用 w:delText
            let deletedRun = run
            xml += "<w:r>"
            let propsXML = deletedRun.properties.toXML()
            if !propsXML.isEmpty {
                xml += "<w:rPr>\(propsXML)</w:rPr>"
            }
            xml += "<w:delText xml:space=\"preserve\">\(escapeXML(deletedRun.text))</w:delText>"
            xml += "</w:r>"
        } else {
            xml += run.toXML()
        }

        if let revision = revision {
            xml += revision.toClosingXML()
        }

        return xml
    }
}

extension TrackedParagraph {
    /// 產生帶修訂標記的段落 XML
    func toXML() -> String {
        var xml = ""

        // 如果整個段落是新插入的
        if let revision = paragraphRevision, revision.type == .insertion {
            xml += revision.toOpeningXML()
        }

        xml += "<w:p>"

        // 段落屬性
        let propsXML = paragraph.properties.toXML()
        if !propsXML.isEmpty {
            xml += "<w:pPr>\(propsXML)</w:pPr>"
        }

        // 輸出所有帶修訂的 Runs
        for trackedRun in trackedRuns {
            xml += trackedRun.toXML()
        }

        xml += "</w:p>"

        if let revision = paragraphRevision, revision.type == .insertion {
            xml += revision.toClosingXML()
        }

        return xml
    }
}

// MARK: - Format Change Support

extension RunProperties {
    /// 產生 rPrChange 內的舊格式 XML
    func toChangeXML() -> String {
        var xml = "<w:rPr>"

        if bold == true { xml += "<w:b/>" }
        if italic == true { xml += "<w:i/>" }
        if let underline = underline { xml += "<w:u w:val=\"\(underline.rawValue)\"/>" }
        if let color = color { xml += "<w:color w:val=\"\(color)\"/>" }
        if let fontSize = fontSize { xml += "<w:sz w:val=\"\(fontSize * 2)\"/>" }
        if let fontName = fontName {
            xml += "<w:rFonts w:ascii=\"\(escapeXML(fontName))\" w:hAnsi=\"\(escapeXML(fontName))\"/>"
        }

        xml += "</w:rPr>"
        return xml
    }
}

// MARK: - Paragraph Format Change

/// 段落格式變更記錄
public struct ParagraphFormatChange: Equatable {
    public var id: Int
    public var author: String
    public var date: Date
    public var previousProperties: ParagraphPropertiesSnapshot

    func toXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        return """
        <w:pPrChange w:id="\(id)" w:author="\(escapeXML(author))" w:date="\(dateString)">
        \(previousProperties.toXML())
        </w:pPrChange>
        """
    }
}

public struct ParagraphPropertiesSnapshot: Equatable {
    public var alignment: Alignment?
    public var spacing: Spacing?
    public var indentation: Indentation?
    public var style: String?
    public var numbering: NumberingInfo?
    public var keepNext: Bool
    public var keepLines: Bool
    public var pageBreakBefore: Bool
    public var sectionBreak: SectionBreakType?
    public var border: ParagraphBorder?
    public var shading: ParagraphShading?

    public init(from properties: ParagraphProperties) {
        self.alignment = properties.alignment
        self.spacing = properties.spacing
        self.indentation = properties.indentation
        self.style = properties.style
        self.numbering = properties.numbering
        self.keepNext = properties.keepNext
        self.keepLines = properties.keepLines
        self.pageBreakBefore = properties.pageBreakBefore
        self.sectionBreak = properties.sectionBreak
        self.border = properties.border
        self.shading = properties.shading
    }

    public func apply(to properties: inout ParagraphProperties) {
        properties.alignment = alignment
        properties.spacing = spacing
        properties.indentation = indentation
        properties.style = style
        properties.numbering = numbering
        properties.keepNext = keepNext
        properties.keepLines = keepLines
        properties.pageBreakBefore = pageBreakBefore
        properties.sectionBreak = sectionBreak
        properties.border = border
        properties.shading = shading
        properties.formatChange = nil
    }

    public func toXML() -> String {
        var properties = ParagraphProperties()
        properties.alignment = alignment
        properties.spacing = spacing
        properties.indentation = indentation
        properties.style = style
        properties.numbering = numbering
        properties.keepNext = keepNext
        properties.keepLines = keepLines
        properties.pageBreakBefore = pageBreakBefore
        properties.sectionBreak = sectionBreak
        properties.border = border
        properties.shading = shading
        return "<w:pPr>\(properties.toXML())</w:pPr>"
    }
}

public struct RunPropertiesSnapshot: Equatable {
    public var bold: Bool = false
    public var italic: Bool = false
    public var underline: UnderlineType?
    public var strikethrough: Bool = false
    public var fontSize: Int?
    public var fontName: String?
    public var color: String?
    public var highlight: HighlightColor?
    public var verticalAlign: VerticalAlign?
    public var characterSpacing: CharacterSpacing?
    public var textEffect: TextEffect?

    public init(from properties: RunProperties) {
        self.bold = properties.bold
        self.italic = properties.italic
        self.underline = properties.underline
        self.strikethrough = properties.strikethrough
        self.fontSize = properties.fontSize
        self.fontName = properties.fontName
        self.color = properties.color
        self.highlight = properties.highlight
        self.verticalAlign = properties.verticalAlign
        self.characterSpacing = properties.characterSpacing
        self.textEffect = properties.textEffect
    }

    public func apply(to properties: inout RunProperties) {
        properties.bold = bold
        properties.italic = italic
        properties.underline = underline
        properties.strikethrough = strikethrough
        properties.fontSize = fontSize
        properties.fontName = fontName
        properties.color = color
        properties.highlight = highlight
        properties.verticalAlign = verticalAlign
        properties.characterSpacing = characterSpacing
        properties.textEffect = textEffect
    }

    public func toXML() -> String {
        var xml = "<w:rPr>"
        if bold { xml += "<w:b/>" }
        if italic { xml += "<w:i/>" }
        if let underline { xml += "<w:u w:val=\"\(underline.rawValue)\"/>" }
        if strikethrough { xml += "<w:strike/>" }
        if let fontSize {
            xml += "<w:sz w:val=\"\(fontSize)\"/>"
            xml += "<w:szCs w:val=\"\(fontSize)\"/>"
        }
        if let fontName {
            xml += "<w:rFonts w:ascii=\"\(escapeXML(fontName))\" w:hAnsi=\"\(escapeXML(fontName))\" w:eastAsia=\"\(escapeXML(fontName))\" w:cs=\"\(escapeXML(fontName))\"/>"
        }
        if let color { xml += "<w:color w:val=\"\(color)\"/>" }
        if let highlight { xml += "<w:highlight w:val=\"\(highlight.rawValue)\"/>" }
        if let verticalAlign { xml += "<w:vertAlign w:val=\"\(verticalAlign.rawValue)\"/>" }
        if let characterSpacing { xml += characterSpacing.toXML() }
        if let textEffect { xml += textEffect.toXML() }
        xml += "</w:rPr>"
        return xml
    }
}

public struct RunFormatChange: Equatable {
    public var revision: Revision
    public var previousProperties: RunPropertiesSnapshot

    public init(revision: Revision, previousProperties: RunPropertiesSnapshot) {
        self.revision = revision
        self.previousProperties = previousProperties
    }

    public func toXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: revision.date)
        return """
        <w:rPrChange w:id="\(revision.id)" w:author="\(escapeXML(revision.author))" w:date="\(dateString)">
        \(previousProperties.toXML())
        </w:rPrChange>
        """
    }
}

// MARK: - Move Tracking

/// 移動追蹤（用於追蹤剪下貼上）
public struct MoveTracking: Equatable {
    public var moveId: String              // 移動配對 ID
    public var fromRevision: Revision      // 來源修訂
    public var toRevision: Revision        // 目標修訂

    /// 產生移動來源範圍 XML
    func toMoveFromRangeXML() -> String {
        return """
        <w:moveFromRangeStart w:id="\(fromRevision.id)" w:name="\(moveId)"/>
        """
    }

    func toMoveFromRangeEndXML() -> String {
        return """
        <w:moveFromRangeEnd w:id="\(fromRevision.id)"/>
        """
    }

    /// 產生移動目標範圍 XML
    func toMoveToRangeXML() -> String {
        return """
        <w:moveToRangeStart w:id="\(toRevision.id)" w:name="\(moveId)"/>
        """
    }

    func toMoveToRangeEndXML() -> String {
        return """
        <w:moveToRangeEnd w:id="\(toRevision.id)"/>
        """
    }
}

// MARK: - Table Revision

/// 表格修訂（插入/刪除列）
public struct TableRowRevision: Equatable {
    public var id: Int
    public var type: RevisionType
    public var author: String
    public var date: Date

    func toTrPrXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        switch type {
        case .insertion:
            return "<w:ins w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\"/>"
        case .deletion:
            return "<w:del w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\"/>"
        default:
            return ""
        }
    }
}

/// 儲存格修訂
public struct TableCellRevision: Equatable {
    public var id: Int
    public var author: String
    public var date: Date
    public var isMerged: Bool              // 是否為合併變更
    public var isDeleted: Bool             // 是否為刪除

    func toTcPrXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        if isDeleted {
            return "<w:cellDel w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\"/>"
        } else if isMerged {
            return "<w:cellMerge w:id=\"\(id)\" w:author=\"\(escapeXML(author))\" w:date=\"\(dateString)\"/>"
        }
        return ""
    }
}

// MARK: - Revision Error

public enum RevisionError: Error, LocalizedError {
    case notFound(Int)
    case trackChangesDisabled
    case cannotAccept(String)
    case cannotReject(String)

    public var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Revision with id \(id) not found"
        case .trackChangesDisabled:
            return "Track changes is not enabled"
        case .cannotAccept(let reason):
            return "Cannot accept revision: \(reason)"
        case .cannotReject(let reason):
            return "Cannot reject revision: \(reason)"
        }
    }
}

// MARK: - Track Changes Settings

/// 修訂追蹤設定
public struct TrackChangesSettings: Equatable {
    public var enabled: Bool = false           // 是否啟用修訂追蹤
    public var author: String = "Unknown"      // 修訂作者
    public var dateTime: Date?                 // 修訂時間（啟用追蹤時才設定）

    public init(enabled: Bool = false, author: String = "Unknown") {
        self.enabled = enabled
        self.author = author
    }
}

// MARK: - Revisions Collection

/// 修訂集合
public struct RevisionsCollection: Equatable {
    public var revisions: [Revision] = []
    public var settings: TrackChangesSettings = TrackChangesSettings()

    /// 取得下一個修訂 ID
    mutating func nextRevisionId() -> Int {
        let maxId = revisions.map { $0.id }.max() ?? 0
        return maxId + 1
    }
}
