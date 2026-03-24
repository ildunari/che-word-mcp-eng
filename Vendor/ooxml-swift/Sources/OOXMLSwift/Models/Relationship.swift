import Foundation

// MARK: - Relationship

/// OOXML 關係定義（用於解析 .rels 檔案）
public struct Relationship {
    public let id: String           // 關係 ID (rId1, rId2, ...)
    public let type: RelationshipType
    public let target: String       // 目標路徑 (media/image1.png)

    public init(id: String, type: RelationshipType, target: String) {
        self.id = id
        self.type = type
        self.target = target
    }
}

// MARK: - Relationship Type

/// 關係類型
public enum RelationshipType: String {
    case image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    case hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    case styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    case numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    case settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    case header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    case footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    case footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    case endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    case comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    case theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    case fontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    case webSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
    case unknown = ""

    public init(rawValue: String) {
        switch rawValue {
        case RelationshipType.image.rawValue:
            self = .image
        case RelationshipType.hyperlink.rawValue:
            self = .hyperlink
        case RelationshipType.styles.rawValue:
            self = .styles
        case RelationshipType.numbering.rawValue:
            self = .numbering
        case RelationshipType.settings.rawValue:
            self = .settings
        case RelationshipType.header.rawValue:
            self = .header
        case RelationshipType.footer.rawValue:
            self = .footer
        case RelationshipType.footnotes.rawValue:
            self = .footnotes
        case RelationshipType.endnotes.rawValue:
            self = .endnotes
        case RelationshipType.comments.rawValue:
            self = .comments
        case RelationshipType.theme.rawValue:
            self = .theme
        case RelationshipType.fontTable.rawValue:
            self = .fontTable
        case RelationshipType.webSettings.rawValue:
            self = .webSettings
        default:
            self = .unknown
        }
    }
}

// MARK: - Relationships Collection

/// 關係集合（來自 .rels 檔案）
public struct RelationshipsCollection {
    public var relationships: [Relationship] = []

    public init() {}

    /// 根據 ID 取得關係
    public func get(by id: String) -> Relationship? {
        return relationships.first { $0.id == id }
    }

    /// 取得所有圖片關係
    public var imageRelationships: [Relationship] {
        return relationships.filter { $0.type == .image }
    }

    /// 取得所有超連結關係
    public var hyperlinkRelationships: [Relationship] {
        return relationships.filter { $0.type == .hyperlink }
    }
}
