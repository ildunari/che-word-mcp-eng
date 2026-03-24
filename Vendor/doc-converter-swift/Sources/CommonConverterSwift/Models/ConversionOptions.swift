import Foundation

/// 保真度層級（參見 docs/lossless-conversion.md §3）
///
/// 遵循資訊下沉原則（§3.4）：每個元素在能被表達的最低 Tier 表達。
/// - Tier 1 (markdown): 純 Markdown，有損但可讀
/// - Tier 2 (markdownWithFigures): + 圖片提取
/// - Tier 3 (marker): MD + Figures + Metadata，bijective / lossless
public enum FidelityTier: Int, Sendable, Comparable {
    case markdown = 1
    case markdownWithFigures = 2
    case marker = 3

    public static func < (lhs: FidelityTier, rhs: FidelityTier) -> Bool {
        lhs.rawValue < rhs.rawValue
    }
}

/// 轉換選項
public struct ConversionOptions: Sendable {
    /// 是否包含 YAML frontmatter
    public var includeFrontmatter: Bool

    /// 是否將軟換行轉為硬換行
    public var hardLineBreaks: Bool

    /// 表格樣式
    public var tableStyle: TableStyle

    /// 標題樣式
    public var headingStyle: HeadingStyle

    /// 保真度層級（預設 .markdown）
    public var fidelity: FidelityTier

    /// 是否啟用 HTML 擴展（Layer B: <u>, <sup>, <sub>, <mark>）
    public var useHTMLExtensions: Bool

    /// Tier 2+: 圖片輸出目錄
    public var figuresDirectory: URL?

    /// Tier 3: metadata sidecar 輸出路徑（.meta.yaml）
    public var metadataOutput: URL?

    /// Practical Mode: 是否保留原始圖片格式（預設 false → 非 web-friendly 格式自動轉 PNG）
    public var preserveOriginalFormat: Bool

    /// Practical Mode: 是否啟用 heading heuristic 統計推斷（預設 true）
    public var headingHeuristic: Bool

    public static let `default` = ConversionOptions(
        includeFrontmatter: false,
        hardLineBreaks: false,
        tableStyle: .pipe,
        headingStyle: .atx
    )

    public init(
        includeFrontmatter: Bool = false,
        hardLineBreaks: Bool = false,
        tableStyle: TableStyle = .pipe,
        headingStyle: HeadingStyle = .atx,
        fidelity: FidelityTier = .markdown,
        useHTMLExtensions: Bool = false,
        figuresDirectory: URL? = nil,
        metadataOutput: URL? = nil,
        preserveOriginalFormat: Bool = false,
        headingHeuristic: Bool = true
    ) {
        self.includeFrontmatter = includeFrontmatter
        self.hardLineBreaks = hardLineBreaks
        self.tableStyle = tableStyle
        self.headingStyle = headingStyle
        self.fidelity = fidelity
        self.useHTMLExtensions = useHTMLExtensions
        self.figuresDirectory = figuresDirectory
        self.metadataOutput = metadataOutput
        self.preserveOriginalFormat = preserveOriginalFormat
        self.headingHeuristic = headingHeuristic
    }

    /// 表格樣式
    public enum TableStyle: Sendable {
        case pipe    // | col1 | col2 |
        case simple  // col1    col2
    }

    /// 標題樣式
    public enum HeadingStyle: Sendable {
        case atx     // # Heading
        case setext  // Heading\n=======
    }
}
