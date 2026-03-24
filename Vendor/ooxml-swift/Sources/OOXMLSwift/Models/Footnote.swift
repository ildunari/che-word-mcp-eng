import Foundation

// MARK: - Footnote

/// 腳註（出現在頁面底部）
public struct Footnote: Equatable {
    public var id: Int                 // 腳註唯一 ID
    public var text: String            // 腳註文字
    public var paragraphIndex: Int     // 腳註附加的段落索引

    public init(id: Int, text: String, paragraphIndex: Int) {
        self.id = id
        self.text = text
        self.paragraphIndex = paragraphIndex
    }
}

// MARK: - Footnote XML Generation

extension Footnote {
    /// 產生 footnotes.xml 中的單一腳註 XML
    func toXML() -> String {
        return """
        <w:footnote w:id="\(id)">
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="FootnoteText"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="FootnoteReference"/>
                    </w:rPr>
                    <w:footnoteRef/>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve"> \(escapeXML(text))</w:t>
                </w:r>
            </w:p>
        </w:footnote>
        """
    }

    /// 產生文件中的腳註參照標記
    func toReferenceXML() -> String {
        return """
        <w:r>
            <w:rPr>
                <w:rStyle w:val="FootnoteReference"/>
            </w:rPr>
            <w:footnoteReference w:id="\(id)"/>
        </w:r>
        """
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Footnotes Collection

/// 腳註集合
public struct FootnotesCollection: Equatable {
    public var footnotes: [Footnote] = []

    /// 取得下一個可用的腳註 ID
    mutating func nextFootnoteId() -> Int {
        // 腳註 ID 從 1 開始（0 和 -1 是保留的分隔符）
        let maxId = footnotes.map { $0.id }.max() ?? 0
        return max(1, maxId + 1)
    }

    /// 產生完整的 footnotes.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:footnote w:type="separator" w:id="-1">
                <w:p>
                    <w:r>
                        <w:separator/>
                    </w:r>
                </w:p>
            </w:footnote>
            <w:footnote w:type="continuationSeparator" w:id="0">
                <w:p>
                    <w:r>
                        <w:continuationSeparator/>
                    </w:r>
                </w:p>
            </w:footnote>
        """

        for footnote in footnotes {
            xml += footnote.toXML()
        }

        xml += "</w:footnotes>"
        return xml
    }

    /// Content Type for footnotes.xml
    public static let contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

    /// Relationship type for footnotes
    public static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
}

// MARK: - Endnote

/// 尾註（出現在文件結尾）
public struct Endnote: Equatable {
    public var id: Int                 // 尾註唯一 ID
    public var text: String            // 尾註文字
    public var paragraphIndex: Int     // 尾註附加的段落索引

    public init(id: Int, text: String, paragraphIndex: Int) {
        self.id = id
        self.text = text
        self.paragraphIndex = paragraphIndex
    }
}

// MARK: - Endnote XML Generation

extension Endnote {
    /// 產生 endnotes.xml 中的單一尾註 XML
    func toXML() -> String {
        return """
        <w:endnote w:id="\(id)">
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="EndnoteText"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="EndnoteReference"/>
                    </w:rPr>
                    <w:endnoteRef/>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve"> \(escapeXML(text))</w:t>
                </w:r>
            </w:p>
        </w:endnote>
        """
    }

    /// 產生文件中的尾註參照標記
    func toReferenceXML() -> String {
        return """
        <w:r>
            <w:rPr>
                <w:rStyle w:val="EndnoteReference"/>
            </w:rPr>
            <w:endnoteReference w:id="\(id)"/>
        </w:r>
        """
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Endnotes Collection

/// 尾註集合
public struct EndnotesCollection: Equatable {
    public var endnotes: [Endnote] = []

    /// 取得下一個可用的尾註 ID
    mutating func nextEndnoteId() -> Int {
        // 尾註 ID 從 1 開始（0 和 -1 是保留的分隔符）
        let maxId = endnotes.map { $0.id }.max() ?? 0
        return max(1, maxId + 1)
    }

    /// 產生完整的 endnotes.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:endnote w:type="separator" w:id="-1">
                <w:p>
                    <w:r>
                        <w:separator/>
                    </w:r>
                </w:p>
            </w:endnote>
            <w:endnote w:type="continuationSeparator" w:id="0">
                <w:p>
                    <w:r>
                        <w:continuationSeparator/>
                    </w:r>
                </w:p>
            </w:endnote>
        """

        for endnote in endnotes {
            xml += endnote.toXML()
        }

        xml += "</w:endnotes>"
        return xml
    }

    /// Content Type for endnotes.xml
    public static let contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"

    /// Relationship type for endnotes
    public static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
}

// MARK: - Errors

public enum FootnoteError: Error, LocalizedError {
    case notFound(Int)
    case invalidParagraphIndex(Int)

    public var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Footnote with id \(id) not found"
        case .invalidParagraphIndex(let index):
            return "Invalid paragraph index: \(index)"
        }
    }
}

public enum EndnoteError: Error, LocalizedError {
    case notFound(Int)
    case invalidParagraphIndex(Int)

    public var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Endnote with id \(id) not found"
        case .invalidParagraphIndex(let index):
            return "Invalid paragraph index: \(index)"
        }
    }
}
