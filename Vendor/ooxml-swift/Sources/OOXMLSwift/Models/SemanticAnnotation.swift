import Foundation

// MARK: - Semantic Element Types

/// 語義元素類型
public enum SemanticElementType: Equatable, Hashable {
    // 結構元素
    case heading(level: Int)           // 1-6
    case paragraph
    case title
    case subtitle

    // 清單元素
    case bulletListItem(level: Int)
    case numberedListItem(level: Int)

    // 表格元素
    case table
    case tableHeaderRow
    case tableDataRow
    case tableCell

    // 豐富內容
    case formula(FormulaType)          // 數學公式
    case image(ImageClassification)    // 圖片（含分類）
    case codeBlock
    case blockquote

    // 特殊元素
    case pageBreak
    case sectionBreak
    case footnote
    case hyperlink

    // 未知
    case unknown(hint: String?)
}

// MARK: - Formula Type

/// 公式類型
public enum FormulaType: Equatable, Hashable {
    case omml                          // Word 原生 OMML (<m:oMath>)
    case mathType                      // MathType OLE 物件
    case latex                         // 嵌入式 LaTeX
    case imageFormula                  // 公式圖片（已確認）
    case unknown                       // 需後續分類
}

// MARK: - Image Classification

/// 圖片分類
public enum ImageClassification: Equatable, Hashable {
    case regular                       // 一般圖片（預設）
    case photo                         // 照片
    case diagram                       // 流程圖/架構圖
    case screenshot                    // 螢幕截圖
    case formulaImage                  // 公式截圖
    case icon                          // 小圖示
    case decoration                    // 裝飾性圖片
    case unknown                       // 需 ImageClassifier 分類
}

// MARK: - Annotation Source

/// 標註來源
public enum AnnotationSource: Equatable, Hashable {
    case parsed        // 解析時確定（如 style="Heading1"）
    case inferred      // 推斷得出（如偵測 <m:oMath>）
    case classified    // 外部分類器結果
    case pending       // 等待分類
}

// MARK: - Semantic Annotation

/// 語義標註
public struct SemanticAnnotation: Equatable, Hashable {
    public var type: SemanticElementType
    public var confidence: Float       // 0.0 - 1.0, 1.0 = 確定
    public var source: AnnotationSource

    public init(
        type: SemanticElementType,
        confidence: Float = 1.0,
        source: AnnotationSource = .parsed
    ) {
        self.type = type
        self.confidence = confidence
        self.source = source
    }

    // MARK: - Convenience Factories

    /// 標題
    public static func heading(_ level: Int) -> SemanticAnnotation {
        SemanticAnnotation(type: .heading(level: level))
    }

    /// 項目清單項目
    public static func bulletItem(level: Int = 0) -> SemanticAnnotation {
        SemanticAnnotation(type: .bulletListItem(level: level))
    }

    /// 編號清單項目
    public static func numberedItem(level: Int = 0) -> SemanticAnnotation {
        SemanticAnnotation(type: .numberedListItem(level: level))
    }

    /// 一般段落
    public static let paragraph = SemanticAnnotation(type: .paragraph)

    /// 分頁符
    public static let pageBreak = SemanticAnnotation(type: .pageBreak)

    /// OMML 公式
    public static let ommlFormula = SemanticAnnotation(
        type: .formula(.omml),
        source: .inferred
    )

    /// 未知圖片（等待分類）
    public static let unknownImage = SemanticAnnotation(
        type: .image(.unknown),
        source: .pending
    )

    /// 一般圖片
    public static let regularImage = SemanticAnnotation(
        type: .image(.regular),
        source: .parsed
    )
}

// MARK: - Helpers

extension SemanticAnnotation {
    /// 是否為標題
    public var isHeading: Bool {
        if case .heading = type { return true }
        return false
    }

    /// 取得標題層級（如果是標題）
    public var headingLevel: Int? {
        if case .heading(let level) = type { return level }
        return nil
    }

    /// 是否為清單項目
    public var isListItem: Bool {
        switch type {
        case .bulletListItem, .numberedListItem:
            return true
        default:
            return false
        }
    }

    /// 是否為公式
    public var isFormula: Bool {
        if case .formula = type { return true }
        return false
    }

    /// 是否為圖片
    public var isImage: Bool {
        if case .image = type { return true }
        return false
    }

    /// 是否需要後續分類
    public var needsClassification: Bool {
        source == .pending
    }
}
