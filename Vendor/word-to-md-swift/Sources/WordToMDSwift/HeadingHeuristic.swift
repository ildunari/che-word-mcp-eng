import OOXMLSwift

/// 統計推斷 heading level（Practical Mode）
///
/// 演算法：
/// 1. 掃描全文段落的 font size（跳過已有 heading style 的）
/// 2. 出現最多的 font size = body text
/// 3. 比 body 大且 bold + 短段落 → heading 候選
/// 4. 候選按大小排序 → H1, H2, H3...
struct HeadingHeuristic {
    private var sizeToLevel: [Int: Int] = [:]  // fontSize (half-points) → heading level
    private var styleMap: [String: Style] = [:] // styleId → Style（用於 resolve 繼承）

    /// 分析文件中的段落，建立 fontSize → heading level 對照表
    mutating func analyze(children: [BodyChild], styles: [Style]) {
        // 建立 style lookup table
        styleMap = Dictionary(uniqueKeysWithValues: styles.map { ($0.id, $0) })

        // 收集 font size 分佈
        var sizeCounts: [Int: Int] = [:]  // fontSize → 段落數
        var headingSizes: Set<Int> = []   // 候選 heading 的 size

        for child in children {
            guard case .paragraph(let para) = child else { continue }
            // 跳過已有 heading style 的段落
            if let style = para.properties.style,
               isHeadingStyle(style, styles: styles) { continue }

            guard let fontSize = effectiveFontSize(para) else { continue }
            sizeCounts[fontSize, default: 0] += 1
        }

        guard !sizeCounts.isEmpty else { return }

        // body size = 出現最多的
        let bodySize = sizeCounts.max(by: { $0.value < $1.value })!.key

        // 比 body 大的 → 候選（還要檢查 bold + 段落短）
        for child in children {
            guard case .paragraph(let para) = child else { continue }
            if let style = para.properties.style,
               isHeadingStyle(style, styles: styles) { continue }

            guard let fontSize = effectiveFontSize(para),
                  fontSize > bodySize else { continue }

            let isShort = para.getText().count < 200

            if effectiveBold(para) && isShort {
                headingSizes.insert(fontSize)
            }
        }

        // 按大小排序 → 對應 H1~H6
        let sorted = headingSizes.sorted(by: >)
        for (index, size) in sorted.prefix(6).enumerated() {
            sizeToLevel[size] = index + 1
        }
    }

    /// 推斷段落的 heading level
    /// - Returns: 1~6（H1~H6），或 nil 表示不是 heading
    func inferLevel(for paragraph: Paragraph) -> Int? {
        guard let fontSize = effectiveFontSize(paragraph) else { return nil }
        guard let level = sizeToLevel[fontSize] else { return nil }

        // 額外驗證：bold + 段落短
        let isShort = paragraph.getText().count < 200
        guard effectiveBold(paragraph) && isShort else { return nil }

        return level
    }

    // MARK: - Private

    /// 取得段落的有效 font size（先看 run-level，再 fallback 到 style 繼承鏈）
    private func effectiveFontSize(_ paragraph: Paragraph) -> Int? {
        // 1. 先嘗試 run-level font size
        let sizes = paragraph.runs.compactMap { $0.properties.fontSize }
        if !sizes.isEmpty {
            return sizes.max()
        }
        // 2. Fallback: 從段落的 style 繼承鏈取 fontSize
        if let styleId = paragraph.properties.style {
            return resolvedFontSize(styleId: styleId)
        }
        return nil
    }

    /// 判斷段落是否為粗體（先看 run-level，再 fallback 到 style 繼承鏈）
    private func effectiveBold(_ paragraph: Paragraph) -> Bool {
        let nonEmptyRuns = paragraph.runs.filter {
            !$0.text.trimmingCharacters(in: .whitespaces).isEmpty
        }
        guard !nonEmptyRuns.isEmpty else { return false }

        // 如果有任何 run 明確設定 bold，以 run-level 為準
        let hasExplicitBold = nonEmptyRuns.contains { $0.properties.bold }
        if hasExplicitBold {
            return nonEmptyRuns.allSatisfy { $0.properties.bold }
        }

        // Fallback: 檢查 style 繼承鏈的 bold
        if let styleId = paragraph.properties.style {
            return resolvedBold(styleId: styleId)
        }
        return false
    }

    /// 從 style 繼承鏈解析 fontSize（遞迴查 basedOn）
    private func resolvedFontSize(styleId: String) -> Int? {
        var current: String? = styleId
        var visited: Set<String> = []
        while let id = current, !visited.contains(id) {
            visited.insert(id)
            if let style = styleMap[id] {
                if let fontSize = style.runProperties?.fontSize {
                    return fontSize
                }
                current = style.basedOn
            } else {
                break
            }
        }
        return nil
    }

    /// 從 style 繼承鏈解析 bold
    private func resolvedBold(styleId: String) -> Bool {
        var current: String? = styleId
        var visited: Set<String> = []
        while let id = current, !visited.contains(id) {
            visited.insert(id)
            if let style = styleMap[id] {
                if let rp = style.runProperties {
                    return rp.bold
                }
                current = style.basedOn
            } else {
                break
            }
        }
        return false
    }

    /// 檢查 style 是否為 heading（與 WordConverter.detectHeadingLevel 相同的模式）
    private func isHeadingStyle(_ styleName: String, styles: [Style]) -> Bool {
        let lower = styleName.lowercased()
        let patterns = ["heading", "標題", "title", "subtitle"]
        return patterns.contains(where: { lower.contains($0) })
    }
}
