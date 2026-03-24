import Foundation

// MARK: - Section Properties

/// 分節屬性（頁面設定）
public struct SectionProperties: Equatable {
    public var pageSize: PageSize
    public var pageMargins: PageMargins
    public var orientation: PageOrientation
    public var headerReference: String?     // rId for header
    public var footerReference: String?     // rId for footer
    public var columns: Int                 // 欄數
    public var docGrid: DocumentGrid?       // 文件格線

    public init(pageSize: PageSize = .letter,
         pageMargins: PageMargins = .normal,
         orientation: PageOrientation = .portrait,
         headerReference: String? = nil,
         footerReference: String? = nil,
         columns: Int = 1,
         docGrid: DocumentGrid? = nil) {
        self.pageSize = pageSize
        self.pageMargins = pageMargins
        self.orientation = orientation
        self.headerReference = headerReference
        self.footerReference = footerReference
        self.columns = columns
        self.docGrid = docGrid
    }
}

// MARK: - Page Size

/// 頁面大小
public struct PageSize: Equatable {
    public var width: Int   // twips (1/20 point)
    public var height: Int  // twips

    public init(width: Int, height: Int) {
        self.width = width
        self.height = height
    }

    // 常用紙張大小
    public static let letter = PageSize(width: 12240, height: 15840)     // 8.5 x 11 inches
    public static let a4 = PageSize(width: 11906, height: 16838)         // 210 x 297 mm
    public static let legal = PageSize(width: 12240, height: 20160)      // 8.5 x 14 inches
    public static let a3 = PageSize(width: 16838, height: 23811)         // 297 x 420 mm
    public static let a5 = PageSize(width: 8391, height: 11906)          // 148 x 210 mm
    public static let b5 = PageSize(width: 10319, height: 14570)         // 182 x 257 mm
    public static let executive = PageSize(width: 10800, height: 14400)  // 7.5 x 10 inches

    /// 從名稱建立頁面大小
    public static func from(name: String) -> PageSize? {
        switch name.lowercased() {
        case "letter": return .letter
        case "a4": return .a4
        case "legal": return .legal
        case "a3": return .a3
        case "a5": return .a5
        case "b5": return .b5
        case "executive": return .executive
        default: return nil
        }
    }

    /// 取得頁面名稱
    public var name: String {
        switch self {
        case .letter: return "Letter"
        case .a4: return "A4"
        case .legal: return "Legal"
        case .a3: return "A3"
        case .a5: return "A5"
        case .b5: return "B5"
        case .executive: return "Executive"
        default: return "Custom (\(widthInInches)\" x \(heightInInches)\")"
        }
    }

    /// 寬度（英寸）
    public var widthInInches: Double {
        Double(width) / 1440.0
    }

    /// 高度（英寸）
    public var heightInInches: Double {
        Double(height) / 1440.0
    }

    /// 橫向版本
    public var landscape: PageSize {
        PageSize(width: height, height: width)
    }
}

// MARK: - Page Margins

/// 頁邊距
public struct PageMargins: Equatable {
    public var top: Int      // twips
    public var right: Int    // twips
    public var bottom: Int   // twips
    public var left: Int     // twips
    public var header: Int   // twips (頁首距離)
    public var footer: Int   // twips (頁尾距離)
    public var gutter: Int   // twips (裝訂邊)

    public init(top: Int = 1440,
         right: Int = 1440,
         bottom: Int = 1440,
         left: Int = 1440,
         header: Int = 720,
         footer: Int = 720,
         gutter: Int = 0) {
        self.top = top
        self.right = right
        self.bottom = bottom
        self.left = left
        self.header = header
        self.footer = footer
        self.gutter = gutter
    }

    // 預設邊距
    public static let normal = PageMargins(top: 1440, right: 1440, bottom: 1440, left: 1440)
    public static let narrow = PageMargins(top: 720, right: 720, bottom: 720, left: 720)
    public static let moderate = PageMargins(top: 1440, right: 1080, bottom: 1440, left: 1080)
    public static let wide = PageMargins(top: 1440, right: 2880, bottom: 1440, left: 2880)

    /// 從名稱建立頁邊距
    public static func from(name: String) -> PageMargins? {
        switch name.lowercased() {
        case "normal": return .normal
        case "narrow": return .narrow
        case "moderate": return .moderate
        case "wide": return .wide
        default: return nil
        }
    }

    /// 取得邊距名稱
    public var name: String {
        switch self {
        case .normal: return "Normal"
        case .narrow: return "Narrow"
        case .moderate: return "Moderate"
        case .wide: return "Wide"
        default: return "Custom"
        }
    }
}

// MARK: - Page Orientation

/// 頁面方向
public enum PageOrientation: String {
    case portrait = "portrait"    // 直向
    case landscape = "landscape"  // 橫向
}

// MARK: - Document Grid

/// 文件格線（用於 CJK 文件）
public struct DocumentGrid: Equatable {
    public var linePitch: Int  // 行距（twips）
    public var charSpace: Int? // 字元間距

    public init(linePitch: Int = 360, charSpace: Int? = nil) {
        self.linePitch = linePitch
        self.charSpace = charSpace
    }
}

// MARK: - Section Break

/// 分節符類型
public enum SectionBreakType: String {
    case nextPage = "nextPage"       // 下一頁
    case continuous = "continuous"   // 連續
    case evenPage = "evenPage"       // 偶數頁
    case oddPage = "oddPage"         // 奇數頁
}

// MARK: - Page Break

/// 分頁符
public struct PageBreak: Equatable {
    // 分頁符是簡單的 <w:br w:type="page"/> 元素
    // 放在 Run 內
}

// MARK: - XML Generation

extension SectionProperties {
    /// 轉換為 XML（完整的 sectPr）
    func toXML() -> String {
        var xml = "<w:sectPr>"

        // 頁首參照
        if let headerRef = headerReference {
            xml += "<w:headerReference w:type=\"default\" r:id=\"\(headerRef)\"/>"
        }

        // 頁尾參照
        if let footerRef = footerReference {
            xml += "<w:footerReference w:type=\"default\" r:id=\"\(footerRef)\"/>"
        }

        // 頁面大小
        var pgSzAttrs = "w:w=\"\(pageSize.width)\" w:h=\"\(pageSize.height)\""
        if orientation == .landscape {
            pgSzAttrs += " w:orient=\"landscape\""
        }
        xml += "<w:pgSz \(pgSzAttrs)/>"

        // 頁邊距
        xml += "<w:pgMar w:top=\"\(pageMargins.top)\" w:right=\"\(pageMargins.right)\" w:bottom=\"\(pageMargins.bottom)\" w:left=\"\(pageMargins.left)\" w:header=\"\(pageMargins.header)\" w:footer=\"\(pageMargins.footer)\" w:gutter=\"\(pageMargins.gutter)\"/>"

        // 欄設定
        xml += "<w:cols w:space=\"720\" w:num=\"\(columns)\"/>"

        // 文件格線
        if let grid = docGrid {
            var gridAttrs = "w:linePitch=\"\(grid.linePitch)\""
            if let charSpace = grid.charSpace {
                gridAttrs += " w:charSpace=\"\(charSpace)\""
            }
            xml += "<w:docGrid \(gridAttrs)/>"
        } else {
            xml += "<w:docGrid w:linePitch=\"360\"/>"
        }

        xml += "</w:sectPr>"
        return xml
    }

    /// 轉換為分節符 XML（放在段落內）
    func toSectionBreakXML(type: SectionBreakType = .nextPage) -> String {
        var xml = "<w:sectPr>"

        // 分節類型
        xml += "<w:type w:val=\"\(type.rawValue)\"/>"

        // 頁面大小
        var pgSzAttrs = "w:w=\"\(pageSize.width)\" w:h=\"\(pageSize.height)\""
        if orientation == .landscape {
            pgSzAttrs += " w:orient=\"landscape\""
        }
        xml += "<w:pgSz \(pgSzAttrs)/>"

        // 頁邊距
        xml += "<w:pgMar w:top=\"\(pageMargins.top)\" w:right=\"\(pageMargins.right)\" w:bottom=\"\(pageMargins.bottom)\" w:left=\"\(pageMargins.left)\" w:header=\"\(pageMargins.header)\" w:footer=\"\(pageMargins.footer)\" w:gutter=\"\(pageMargins.gutter)\"/>"

        xml += "</w:sectPr>"
        return xml
    }
}

extension PageBreak {
    /// 轉換為 XML
    func toXML() -> String {
        return "<w:r><w:br w:type=\"page\"/></w:r>"
    }
}
