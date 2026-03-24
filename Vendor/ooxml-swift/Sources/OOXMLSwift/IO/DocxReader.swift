import Foundation

/// DOCX 檔案讀取器
public struct DocxReader {

    /// 讀取 .docx 檔案並解析為 WordDocument
    public static func read(from url: URL) throws -> WordDocument {
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw WordError.fileNotFound(url.path)
        }

        // 1. 解壓縮 ZIP
        let tempDir = try ZipHelper.unzip(url)

        defer {
            ZipHelper.cleanup(tempDir)
        }

        // 2. 讀取關係檔案 word/_rels/document.xml.rels
        let relationships = try parseRelationships(from: tempDir)

        // 3. 提取圖片資源
        let images = try extractImages(from: tempDir, relationships: relationships)

        // 4. 讀取 document.xml
        let documentURL = tempDir.appendingPathComponent("word/document.xml")
        guard FileManager.default.fileExists(atPath: documentURL.path) else {
            throw WordError.parseError("找不到 word/document.xml")
        }

        let documentData = try Data(contentsOf: documentURL)
        let documentXML = try XMLDocument(data: documentData)

        // 5. 讀取 styles.xml（先解析，用於語義標註）
        var document = WordDocument()
        let stylesURL = tempDir.appendingPathComponent("word/styles.xml")
        if FileManager.default.fileExists(atPath: stylesURL.path) {
            let stylesData = try Data(contentsOf: stylesURL)
            let stylesXML = try XMLDocument(data: stylesData)
            document.styles = try parseStyles(from: stylesXML)
        }

        // 6. 讀取 numbering.xml（可選，用於清單語義標註）
        let numberingURL = tempDir.appendingPathComponent("word/numbering.xml")
        if FileManager.default.fileExists(atPath: numberingURL.path) {
            let numberingData = try Data(contentsOf: numberingURL)
            let numberingXML = try XMLDocument(data: numberingData)
            document.numbering = try parseNumbering(from: numberingXML)
        }

        let settingsURL = tempDir.appendingPathComponent("word/settings.xml")
        if FileManager.default.fileExists(atPath: settingsURL.path) {
            let settingsData = try Data(contentsOf: settingsURL)
            let settingsXML = try XMLDocument(data: settingsData)
            parseSettings(from: settingsXML, into: &document)
        }

        // 7. 解析文件內容（傳入 styles 和 numbering 用於語義標註）
        document.body = try parseBody(
            from: documentXML,
            relationships: relationships,
            styles: document.styles,
            numbering: document.numbering
        )
        document.revisions.revisions = collectParsedRevisions(from: document.body)
        document.images = images

        // 8. 讀取 core.xml（可選）
        let coreURL = tempDir.appendingPathComponent("docProps/core.xml")
        if FileManager.default.fileExists(atPath: coreURL.path) {
            let coreData = try Data(contentsOf: coreURL)
            let coreXML = try XMLDocument(data: coreData)
            document.properties = try parseCoreProperties(from: coreXML)
        }

        // 8. 讀取 comments.xml（可選）
        let commentsURL = tempDir.appendingPathComponent("word/comments.xml")
        if FileManager.default.fileExists(atPath: commentsURL.path) {
            let commentsData = try Data(contentsOf: commentsURL)
            let commentsXML = try XMLDocument(data: commentsData)
            document.comments = try parseComments(from: commentsXML)
        }

        // 9. Link comment paragraphIndex from paragraph commentIds
        for (index, child) in document.body.children.enumerated() {
            if case .paragraph(let para) = child {
                for commentId in para.commentIds {
                    if let idx = document.comments.comments.firstIndex(where: { $0.id == commentId }) {
                        document.comments.comments[idx].paragraphIndex = index
                    }
                }
            }
        }

        // 10. 讀取 commentsExtended.xml（可選，Word 2012+ 回覆與已解決狀態）
        let commentsExtURL = tempDir.appendingPathComponent("word/commentsExtended.xml")
        if FileManager.default.fileExists(atPath: commentsExtURL.path) {
            let extData = try Data(contentsOf: commentsExtURL)
            let extXML = try XMLDocument(data: extData)
            try parseCommentsExtended(from: extXML, into: &document.comments)
        }

        return document
    }

    // MARK: - Relationships Parsing

    /// 解析關係檔案
    private static func parseRelationships(from tempDir: URL) throws -> RelationshipsCollection {
        var collection = RelationshipsCollection()

        let relsURL = tempDir.appendingPathComponent("word/_rels/document.xml.rels")
        guard FileManager.default.fileExists(atPath: relsURL.path) else {
            // 沒有關係檔案也是合法的
            return collection
        }

        let relsData = try Data(contentsOf: relsURL)
        let relsXML = try XMLDocument(data: relsData)

        // 取得所有 Relationship 節點
        let relNodes = try relsXML.nodes(forXPath: "//*[local-name()='Relationship']")

        for node in relNodes {
            guard let element = node as? XMLElement else { continue }

            guard let id = element.attribute(forName: "Id")?.stringValue,
                  let typeStr = element.attribute(forName: "Type")?.stringValue,
                  let target = element.attribute(forName: "Target")?.stringValue else {
                continue
            }

            let relationship = Relationship(
                id: id,
                type: RelationshipType(rawValue: typeStr),
                target: target
            )
            collection.relationships.append(relationship)
        }

        return collection
    }

    private static func parseSettings(from xml: XMLDocument, into document: inout WordDocument) {
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='trackRevisions']"), !nodes.isEmpty {
            document.revisions.settings.enabled = true
        }
    }

    private static func collectParsedRevisions(from body: Body) -> [Revision] {
        var revisions: [Revision] = []
        for (paragraphIndex, paragraph) in body.children.compactMap({ child -> Paragraph? in
            if case .paragraph(let paragraph) = child { return paragraph }
            return nil
        }).enumerated() {
            if var revision = paragraph.paragraphRevision {
                revision.paragraphIndex = paragraphIndex
                revisions.append(revision)
            }
            if let formatChange = paragraph.properties.formatChange {
                revisions.append(
                    Revision(
                        id: formatChange.id,
                        type: .paragraphChange,
                        author: formatChange.author,
                        paragraphIndex: paragraphIndex,
                        date: formatChange.date
                    )
                )
            }
            for run in paragraph.runs {
                if let formatChange = run.properties.formatChange {
                    var revision = formatChange.revision
                    revision.paragraphIndex = paragraphIndex
                    revisions.append(revision)
                }
            }
            for trackedRun in paragraph.trackedRuns ?? [] {
                if var revision = trackedRun.revision {
                    revision.paragraphIndex = paragraphIndex
                    revision.originalText = trackedRun.isDeleted ? trackedRun.run.text : revision.originalText
                    revision.newText = trackedRun.isDeleted ? revision.newText : trackedRun.run.text
                    revisions.append(revision)
                }
                if let formatChange = trackedRun.run.properties.formatChange {
                    var revision = formatChange.revision
                    revision.paragraphIndex = paragraphIndex
                    revisions.append(revision)
                }
            }
        }
        return revisions.sorted { $0.id < $1.id }
    }

    // MARK: - Image Extraction

    /// 從 word/media/ 提取圖片
    private static func extractImages(from tempDir: URL, relationships: RelationshipsCollection) throws -> [ImageReference] {
        var images: [ImageReference] = []

        let mediaDir = tempDir.appendingPathComponent("word/media")
        guard FileManager.default.fileExists(atPath: mediaDir.path) else {
            // 沒有 media 目錄也是合法的
            return images
        }

        // 建立 target → rId 的映射
        var targetToId: [String: String] = [:]
        for rel in relationships.imageRelationships {
            // target 可能是 "media/image1.png" 或 "../media/image1.png"
            let normalizedTarget = rel.target.replacingOccurrences(of: "../", with: "")
            targetToId[normalizedTarget] = rel.id
        }

        // 讀取 media 目錄中的所有檔案
        let contents = try FileManager.default.contentsOfDirectory(atPath: mediaDir.path)

        for fileName in contents {
            let fileURL = mediaDir.appendingPathComponent(fileName)

            // 檢查是否為檔案（不是目錄）
            var isDirectory: ObjCBool = false
            guard FileManager.default.fileExists(atPath: fileURL.path, isDirectory: &isDirectory),
                  !isDirectory.boolValue else {
                continue
            }

            // 讀取檔案資料
            let data = try Data(contentsOf: fileURL)

            // 找對應的 relationship ID
            let targetPath = "media/\(fileName)"
            let relationshipId = targetToId[targetPath] ?? "rId_\(fileName)"

            // 取得 MIME 類型
            let ext = (fileName as NSString).pathExtension.lowercased()
            let contentType = mimeType(for: ext)

            let imageRef = ImageReference(
                id: relationshipId,
                fileName: fileName,
                contentType: contentType,
                data: data
            )
            images.append(imageRef)
        }

        return images
    }

    /// 取得副檔名對應的 MIME 類型
    private static func mimeType(for ext: String) -> String {
        switch ext {
        case "png": return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "gif": return "image/gif"
        case "bmp": return "image/bmp"
        case "tiff", "tif": return "image/tiff"
        case "webp": return "image/webp"
        case "emf": return "image/x-emf"
        case "wmf": return "image/x-wmf"
        default: return "application/octet-stream"
        }
    }

    // MARK: - Body Parsing

    private static func parseBody(
        from xml: XMLDocument,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> Body {
        var body = Body()

        // 取得所有段落和表格節點
        // XPath: //w:body/*
        let bodyNodes = try xml.nodes(forXPath: "//*[local-name()='body']/*")

        let count = bodyNodes.count
        guard count > 0 else { return body }

        // 小文件走 serial path（GCD dispatch overhead 不划算）
        let coreCount = ProcessInfo.processInfo.activeProcessorCount
        if count <= 200 || coreCount <= 1 {
            for node in bodyNodes {
                guard let element = node as? XMLElement else { continue }

                if element.localName == "p" {
                    let paragraph = try parseParagraph(
                        from: element,
                        relationships: relationships,
                        styles: styles,
                        numbering: numbering
                    )
                    body.children.append(.paragraph(paragraph))
                } else if element.localName == "ins",
                          let paragraphElement = element.children?.compactMap({ $0 as? XMLElement }).first(where: { $0.localName == "p" }) {
                    var paragraph = try parseParagraph(
                        from: paragraphElement,
                        relationships: relationships,
                        styles: styles,
                        numbering: numbering
                    )
                    paragraph.paragraphRevision = parseRevision(from: element, defaultType: .insertion)
                    body.children.append(.paragraph(paragraph))
                } else if element.localName == "tbl" {
                    let table = try parseTable(
                        from: element,
                        relationships: relationships,
                        styles: styles,
                        numbering: numbering
                    )
                    body.children.append(.table(table))
                    body.tables.append(table)
                }
            }
            return body
        }

        // 大文件走 parallel path
        // 每個 body child 的子樹互不相交，shared data (relationships/styles/numbering) 為唯讀
        // 使用 concurrentPerform 分 chunk 平行解析，再按順序合併
        let chunkSize = max(count / coreCount, 64)
        let chunkCount = (count + chunkSize - 1) / chunkSize

        // 每個 chunk 的結果陣列（各 chunk 獨立寫入，無 data race）
        let chunkResults = UnsafeMutablePointer<[BodyChild]>.allocate(capacity: chunkCount)
        chunkResults.initialize(repeating: [], count: chunkCount)
        defer {
            chunkResults.deinitialize(count: chunkCount)
            chunkResults.deallocate()
        }

        // Thread-safe error capture
        var firstError: Error?
        let errorLock = NSLock()

        DispatchQueue.concurrentPerform(iterations: chunkCount) { ci in
            // Early exit if another chunk already failed
            errorLock.lock()
            let hasError = firstError != nil
            errorLock.unlock()
            if hasError { return }

            let start = ci * chunkSize
            let end = min(start + chunkSize, count)
            var children: [BodyChild] = []
            children.reserveCapacity(end - start)

            for i in start..<end {
                guard let element = bodyNodes[i] as? XMLElement else { continue }

                do {
                    if element.localName == "p" {
                        let paragraph = try parseParagraph(
                            from: element,
                            relationships: relationships,
                            styles: styles,
                            numbering: numbering
                        )
                        children.append(.paragraph(paragraph))
                    } else if element.localName == "ins",
                              let paragraphElement = element.children?.compactMap({ $0 as? XMLElement }).first(where: { $0.localName == "p" }) {
                        var paragraph = try parseParagraph(
                            from: paragraphElement,
                            relationships: relationships,
                            styles: styles,
                            numbering: numbering
                        )
                        paragraph.paragraphRevision = parseRevision(from: element, defaultType: .insertion)
                        children.append(.paragraph(paragraph))
                    } else if element.localName == "tbl" {
                        let table = try parseTable(
                            from: element,
                            relationships: relationships,
                            styles: styles,
                            numbering: numbering
                        )
                        children.append(.table(table))
                    }
                } catch {
                    errorLock.lock()
                    if firstError == nil { firstError = error }
                    errorLock.unlock()
                    return
                }
            }
            chunkResults[ci] = children
        }

        if let error = firstError { throw error }

        // 按 chunk 順序合併，保持文件元素的原始順序
        body.children.reserveCapacity(count)
        for ci in 0..<chunkCount {
            for child in chunkResults[ci] {
                body.children.append(child)
                if case .table(let table) = child {
                    body.tables.append(table)
                }
            }
        }

        return body
    }

    // MARK: - Paragraph Parsing

    private static func parseParagraph(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> Paragraph {
        var paragraph = Paragraph()

        // 解析段落屬性
        if let pPr = element.elements(forName: "w:pPr").first {
            paragraph.properties = parseParagraphProperties(from: pPr)
        }

        var trackedRuns: [TrackedRun] = []

        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            switch childElement.localName {
            case "r":
                let parsedRun = try parseRun(from: childElement, relationships: relationships)
                paragraph.runs.append(parsedRun)
                trackedRuns.append(TrackedRun(run: parsedRun))
            case "ins":
                let revision = parseRevision(from: childElement, defaultType: .insertion)
                for runElement in childElement.children?.compactMap({ $0 as? XMLElement }).filter({ $0.localName == "r" }) ?? [] {
                    let parsedRun = try parseRun(from: runElement, relationships: relationships)
                    paragraph.runs.append(parsedRun)
                    trackedRuns.append(TrackedRun(run: parsedRun, revision: revision))
                }
            case "del":
                let revision = parseRevision(from: childElement, defaultType: .deletion)
                for runElement in childElement.children?.compactMap({ $0 as? XMLElement }).filter({ $0.localName == "r" }) ?? [] {
                    let parsedRun = try parseRun(from: runElement, relationships: relationships)
                    trackedRuns.append(TrackedRun(run: parsedRun, revision: revision, isDeleted: true))
                }
            default:
                continue
            }
        }

        if trackedRuns.contains(where: { $0.revision != nil || $0.isDeleted }) {
            paragraph.trackedRuns = trackedRuns
        }

        // 解析 comment anchors（commentRangeStart）
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            if childElement.localName == "commentRangeStart",
               let idStr = childElement.attribute(forName: "w:id")?.stringValue,
               let id = Int(idStr) {
                paragraph.commentIds.append(id)
            }
        }

        // 🆕 語義標註
        paragraph.semantic = detectParagraphSemantic(
            properties: paragraph.properties,
            runs: paragraph.runs,
            styles: styles,
            numbering: numbering
        )

        return paragraph
    }

    private static func parseParagraphProperties(from element: XMLElement) -> ParagraphProperties {
        var props = ParagraphProperties()

        // 樣式
        if let pStyle = element.elements(forName: "w:pStyle").first,
           let val = pStyle.attribute(forName: "w:val")?.stringValue {
            props.style = val
        }

        // 對齊
        if let jc = element.elements(forName: "w:jc").first,
           let val = jc.attribute(forName: "w:val")?.stringValue {
            props.alignment = Alignment(rawValue: val)
        }

        // 間距
        if let spacing = element.elements(forName: "w:spacing").first {
            var spacingProps = Spacing()
            if let before = spacing.attribute(forName: "w:before")?.stringValue {
                spacingProps.before = Int(before)
            }
            if let after = spacing.attribute(forName: "w:after")?.stringValue {
                spacingProps.after = Int(after)
            }
            if let line = spacing.attribute(forName: "w:line")?.stringValue {
                spacingProps.line = Int(line)
            }
            if let lineRule = spacing.attribute(forName: "w:lineRule")?.stringValue {
                spacingProps.lineRule = LineRule(rawValue: lineRule)
            }
            props.spacing = spacingProps
        }

        // 縮排
        if let ind = element.elements(forName: "w:ind").first {
            var indentation = Indentation()
            if let left = ind.attribute(forName: "w:left")?.stringValue {
                indentation.left = Int(left)
            }
            if let right = ind.attribute(forName: "w:right")?.stringValue {
                indentation.right = Int(right)
            }
            if let firstLine = ind.attribute(forName: "w:firstLine")?.stringValue {
                indentation.firstLine = Int(firstLine)
            }
            if let hanging = ind.attribute(forName: "w:hanging")?.stringValue {
                indentation.hanging = Int(hanging)
            }
            props.indentation = indentation
        }

        // 編號/項目符號 (w:numPr)
        if let numPr = element.elements(forName: "w:numPr").first {
            var numInfo: NumberingInfo?
            var numId: Int?
            var level: Int = 0

            if let ilvl = numPr.elements(forName: "w:ilvl").first,
               let val = ilvl.attribute(forName: "w:val")?.stringValue {
                level = Int(val) ?? 0
            }
            if let numIdEl = numPr.elements(forName: "w:numId").first,
               let val = numIdEl.attribute(forName: "w:val")?.stringValue {
                numId = Int(val)
            }

            if let id = numId {
                numInfo = NumberingInfo(numId: id, level: level)
            }
            props.numbering = numInfo
        }

        // 分頁控制
        if element.elements(forName: "w:keepNext").first != nil {
            props.keepNext = true
        }
        if element.elements(forName: "w:keepLines").first != nil {
            props.keepLines = true
        }
        if element.elements(forName: "w:pageBreakBefore").first != nil {
            props.pageBreakBefore = true
        }

        if let change = element.elements(forName: "w:pPrChange").first
            ?? element.children?.compactMap({ $0 as? XMLElement }).first(where: { $0.localName == "pPrChange" }) {
            let revision = parseRevision(from: change, defaultType: .paragraphChange)
            let previousNode = change.elements(forName: "w:pPr").first
                ?? change.children?.compactMap({ $0 as? XMLElement }).first(where: { $0.localName == "pPr" })
            let previousProperties = previousNode.map(parseParagraphProperties(from:)) ?? ParagraphProperties()
            props.formatChange = ParagraphFormatChange(
                id: revision.id,
                author: revision.author,
                date: revision.date,
                previousProperties: ParagraphPropertiesSnapshot(from: previousProperties)
            )
        }

        return props
    }

    // MARK: - Run Parsing

    private static func parseRun(from element: XMLElement, relationships: RelationshipsCollection) throws -> Run {
        var run = Run(text: "")

        // 解析 Run 屬性
        if let rPr = element.elements(forName: "w:rPr").first {
            run.properties = parseRunProperties(from: rPr)
        }

        // 解析文字
        for t in element.elements(forName: "w:t") {
            run.text += t.stringValue ?? ""
        }
        for delText in element.elements(forName: "w:delText") {
            run.text += delText.stringValue ?? ""
        }

        // 解析圖片 (w:drawing)
        if let drawingElement = element.elements(forName: "w:drawing").first {
            run.drawing = try parseDrawing(from: drawingElement, relationships: relationships)
            // 🆕 圖片語義標註（標為 unknown，等後續分類）
            run.semantic = SemanticAnnotation.unknownImage
        }

        // 🆕 檢查是否為 OMML 公式 (m:oMath 或 m:oMathPara)
        // 使用 children 遍歷取代 XPath，避免 O(n²) 效能問題
        if let oMathElement = findFirstDescendant(of: element, localNames: ["oMath", "oMathPara"]) {
            run.rawXML = oMathElement.xmlString
            run.semantic = SemanticAnnotation.ommlFormula
        }

        return run
    }

    // MARK: - Drawing Parsing

    /// 解析 <w:drawing> 元素
    private static func parseDrawing(from element: XMLElement, relationships: RelationshipsCollection) throws -> Drawing? {
        // 尋找 inline 或 anchor 元素（使用 children 遍歷取代 XPath）
        if let inlineElement = findFirstDescendant(of: element, localNames: ["inline"]) {
            return try parseInlineDrawing(from: inlineElement, relationships: relationships)
        } else if let anchorElement = findFirstDescendant(of: element, localNames: ["anchor"]) {
            return try parseAnchorDrawing(from: anchorElement, relationships: relationships)
        }

        return nil
    }

    /// 解析 inline drawing
    private static func parseInlineDrawing(from element: XMLElement, relationships: RelationshipsCollection) throws -> Drawing? {
        // 取得尺寸 (wp:extent)
        guard let extentElement = findFirstDescendant(of: element, localNames: ["extent"]),
              let cxStr = extentElement.attribute(forName: "cx")?.stringValue,
              let cyStr = extentElement.attribute(forName: "cy")?.stringValue,
              let cx = Int(cxStr),
              let cy = Int(cyStr) else {
            return nil
        }

        // 取得圖片參照 (a:blip r:embed)
        guard let blipElement = findFirstDescendant(of: element, localNames: ["blip"]) else {
            return nil
        }

        // r:embed 屬性包含 relationship ID
        let embedId = blipElement.attribute(forName: "r:embed")?.stringValue
            ?? blipElement.attribute(forName: "embed")?.stringValue

        guard let imageId = embedId else {
            return nil
        }

        // 取得圖片名稱和描述 (wp:docPr)
        var name = "Picture"
        var description = ""

        if let docPrElement = findFirstDescendant(of: element, localNames: ["docPr"]) {
            if let nameAttr = docPrElement.attribute(forName: "name")?.stringValue {
                name = nameAttr
            }
            if let descrAttr = docPrElement.attribute(forName: "descr")?.stringValue {
                description = descrAttr
            }
        }

        let drawing = Drawing(
            type: .inline,
            width: cx,
            height: cy,
            imageId: imageId,
            name: name,
            description: description
        )

        return drawing
    }

    /// 解析 anchor drawing (浮動圖片)
    private static func parseAnchorDrawing(from element: XMLElement, relationships: RelationshipsCollection) throws -> Drawing? {
        // 取得尺寸
        guard let extentElement = findFirstDescendant(of: element, localNames: ["extent"]),
              let cxStr = extentElement.attribute(forName: "cx")?.stringValue,
              let cyStr = extentElement.attribute(forName: "cy")?.stringValue,
              let cx = Int(cxStr),
              let cy = Int(cyStr) else {
            return nil
        }

        // 取得圖片參照
        guard let blipElement = findFirstDescendant(of: element, localNames: ["blip"]) else {
            return nil
        }

        let embedId = blipElement.attribute(forName: "r:embed")?.stringValue
            ?? blipElement.attribute(forName: "embed")?.stringValue

        guard let imageId = embedId else {
            return nil
        }

        // 取得名稱和描述
        var name = "Picture"
        var description = ""

        if let docPrElement = findFirstDescendant(of: element, localNames: ["docPr"]) {
            if let nameAttr = docPrElement.attribute(forName: "name")?.stringValue {
                name = nameAttr
            }
            if let descrAttr = docPrElement.attribute(forName: "descr")?.stringValue {
                description = descrAttr
            }
        }

        var drawing = Drawing(
            type: .anchor,
            width: cx,
            height: cy,
            imageId: imageId,
            name: name,
            description: description
        )

        // 解析定位屬性
        var anchorPos = AnchorPosition()

        // 水平定位
        if let posHElement = findFirstDescendant(of: element, localNames: ["positionH"]) {
            if let relativeFrom = posHElement.attribute(forName: "relativeFrom")?.stringValue {
                anchorPos.horizontalRelativeFrom = HorizontalRelativeFrom(rawValue: relativeFrom) ?? .column
            }

            // posOffset 或 align
            let hOffsetElement = findFirstDescendant(of: posHElement, localNames: ["posOffset"])
            let hAlignElement = findFirstDescendant(of: posHElement, localNames: ["align"])

            if let offsetEl = hOffsetElement, let offsetStr = offsetEl.stringValue, let offset = Int(offsetStr) {
                anchorPos.horizontalOffset = offset
            } else if let alignEl = hAlignElement, let alignStr = alignEl.stringValue {
                anchorPos.horizontalAlignment = HorizontalAlignment(rawValue: alignStr)
            }
        }

        // 垂直定位
        if let posVElement = findFirstDescendant(of: element, localNames: ["positionV"]) {
            if let relativeFrom = posVElement.attribute(forName: "relativeFrom")?.stringValue {
                anchorPos.verticalRelativeFrom = VerticalRelativeFrom(rawValue: relativeFrom) ?? .paragraph
            }

            let offsetElement = findFirstDescendant(of: posVElement, localNames: ["posOffset"])
            let alignElement = findFirstDescendant(of: posVElement, localNames: ["align"])

            if let offsetEl = offsetElement, let offsetStr = offsetEl.stringValue, let offset = Int(offsetStr) {
                anchorPos.verticalOffset = offset
            } else if let alignEl = alignElement, let alignStr = alignEl.stringValue {
                anchorPos.verticalAlignment = VerticalAlignment(rawValue: alignStr)
            }
        }

        drawing.anchorPosition = anchorPos

        return drawing
    }

    private static func parseRunProperties(from element: XMLElement) -> RunProperties {
        var props = RunProperties()

        // 粗體
        if element.elements(forName: "w:b").first != nil {
            props.bold = true
        }

        // 斜體
        if element.elements(forName: "w:i").first != nil {
            props.italic = true
        }

        // 底線
        if let u = element.elements(forName: "w:u").first,
           let val = u.attribute(forName: "w:val")?.stringValue {
            props.underline = UnderlineType(rawValue: val)
        }

        // 刪除線
        if element.elements(forName: "w:strike").first != nil {
            props.strikethrough = true
        }

        // 字型大小
        if let sz = element.elements(forName: "w:sz").first,
           let val = sz.attribute(forName: "w:val")?.stringValue {
            props.fontSize = Int(val)
        }

        // 字型
        if let rFonts = element.elements(forName: "w:rFonts").first,
           let ascii = rFonts.attribute(forName: "w:ascii")?.stringValue {
            props.fontName = ascii
        }

        // 顏色
        if let color = element.elements(forName: "w:color").first,
           let val = color.attribute(forName: "w:val")?.stringValue {
            props.color = val
        }

        // 螢光標記
        if let highlight = element.elements(forName: "w:highlight").first,
           let val = highlight.attribute(forName: "w:val")?.stringValue {
            props.highlight = HighlightColor(rawValue: val)
        }

        // 垂直對齊
        if let vertAlign = element.elements(forName: "w:vertAlign").first,
           let val = vertAlign.attribute(forName: "w:val")?.stringValue {
            props.verticalAlign = VerticalAlign(rawValue: val)
        }

        if let change = element.elements(forName: "w:rPrChange").first
            ?? element.children?.compactMap({ $0 as? XMLElement }).first(where: { $0.localName == "rPrChange" }) {
            let revision = parseRevision(from: change, defaultType: .formatChange)
            let previousNode = change.elements(forName: "w:rPr").first
                ?? change.children?.compactMap({ $0 as? XMLElement }).first(where: { $0.localName == "rPr" })
            let previousProperties = previousNode.map(parseRunProperties(from:)) ?? RunProperties()
            props.formatChange = RunFormatChange(
                revision: revision,
                previousProperties: RunPropertiesSnapshot(from: previousProperties)
            )
        }

        return props
    }

    private static func parseRevision(from element: XMLElement, defaultType: RevisionType) -> Revision {
        let idString = element.attribute(forName: "w:id")?.stringValue
            ?? element.attribute(forName: "id")?.stringValue
            ?? "0"
        let author = element.attribute(forName: "w:author")?.stringValue
            ?? element.attribute(forName: "author")?.stringValue
            ?? "Unknown"
        let dateString = element.attribute(forName: "w:date")?.stringValue
            ?? element.attribute(forName: "date")?.stringValue
        let date = dateString.flatMap { ISO8601DateFormatter().date(from: $0) } ?? Date()

        return Revision(
            id: Int(idString) ?? 0,
            type: RevisionType(rawValue: element.localName ?? defaultType.rawValue) ?? defaultType,
            author: author,
            date: date
        )
    }

    // MARK: - Table Parsing

    private static func parseTable(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> Table {
        var table = Table()

        // 解析表格屬性
        if let tblPr = element.elements(forName: "w:tblPr").first {
            table.properties = parseTableProperties(from: tblPr)
        }

        // 解析表格行
        for tr in element.elements(forName: "w:tr") {
            let row = try parseTableRow(
                from: tr,
                relationships: relationships,
                styles: styles,
                numbering: numbering
            )
            table.rows.append(row)
        }

        return table
    }

    private static func parseTableProperties(from element: XMLElement) -> TableProperties {
        var props = TableProperties()

        // 寬度
        if let tblW = element.elements(forName: "w:tblW").first {
            if let w = tblW.attribute(forName: "w:w")?.stringValue {
                props.width = Int(w)
            }
            if let type = tblW.attribute(forName: "w:type")?.stringValue {
                props.widthType = WidthType(rawValue: type)
            }
        }

        // 對齊
        if let jc = element.elements(forName: "w:jc").first,
           let val = jc.attribute(forName: "w:val")?.stringValue {
            props.alignment = Alignment(rawValue: val)
        }

        // 版面配置
        if let layout = element.elements(forName: "w:tblLayout").first,
           let val = layout.attribute(forName: "w:type")?.stringValue {
            props.layout = TableLayout(rawValue: val)
        }

        return props
    }

    private static func parseTableRow(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> TableRow {
        var row = TableRow()

        // 解析行屬性
        if let trPr = element.elements(forName: "w:trPr").first {
            row.properties = parseTableRowProperties(from: trPr)
        }

        // 解析儲存格
        for tc in element.elements(forName: "w:tc") {
            let cell = try parseTableCell(
                from: tc,
                relationships: relationships,
                styles: styles,
                numbering: numbering
            )
            row.cells.append(cell)
        }

        return row
    }

    private static func parseTableRowProperties(from element: XMLElement) -> TableRowProperties {
        var props = TableRowProperties()

        // 行高
        if let trHeight = element.elements(forName: "w:trHeight").first {
            if let val = trHeight.attribute(forName: "w:val")?.stringValue {
                props.height = Int(val)
            }
            if let hRule = trHeight.attribute(forName: "w:hRule")?.stringValue {
                props.heightRule = HeightRule(rawValue: hRule)
            }
        }

        // 表頭行
        if element.elements(forName: "w:tblHeader").first != nil {
            props.isHeader = true
        }

        // 禁止分割
        if element.elements(forName: "w:cantSplit").first != nil {
            props.cantSplit = true
        }

        return props
    }

    private static func parseTableCell(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> TableCell {
        var cell = TableCell()
        cell.paragraphs = []

        // 解析儲存格屬性
        if let tcPr = element.elements(forName: "w:tcPr").first {
            cell.properties = parseTableCellProperties(from: tcPr)
        }

        // 解析段落（傳入 styles 和 numbering 用於語義標註）
        for p in element.elements(forName: "w:p") {
            let para = try parseParagraph(
                from: p,
                relationships: relationships,
                styles: styles,
                numbering: numbering
            )
            cell.paragraphs.append(para)
        }

        // 確保至少有一個段落
        if cell.paragraphs.isEmpty {
            cell.paragraphs.append(Paragraph())
        }

        return cell
    }

    private static func parseTableCellProperties(from element: XMLElement) -> TableCellProperties {
        var props = TableCellProperties()

        // 寬度
        if let tcW = element.elements(forName: "w:tcW").first {
            if let w = tcW.attribute(forName: "w:w")?.stringValue {
                props.width = Int(w)
            }
            if let type = tcW.attribute(forName: "w:type")?.stringValue {
                props.widthType = WidthType(rawValue: type)
            }
        }

        // 水平合併
        if let gridSpan = element.elements(forName: "w:gridSpan").first,
           let val = gridSpan.attribute(forName: "w:val")?.stringValue {
            props.gridSpan = Int(val)
        }

        // 垂直合併
        if let vMerge = element.elements(forName: "w:vMerge").first,
           let val = vMerge.attribute(forName: "w:val")?.stringValue {
            props.verticalMerge = VerticalMerge(rawValue: val)
        }

        // 垂直對齊
        if let vAlign = element.elements(forName: "w:vAlign").first,
           let val = vAlign.attribute(forName: "w:val")?.stringValue {
            props.verticalAlignment = CellVerticalAlignment(rawValue: val)
        }

        // 底色
        if let shd = element.elements(forName: "w:shd").first,
           let fill = shd.attribute(forName: "w:fill")?.stringValue {
            var shading = CellShading(fill: fill)
            if let color = shd.attribute(forName: "w:color")?.stringValue {
                shading.color = color
            }
            if let val = shd.attribute(forName: "w:val")?.stringValue {
                shading.pattern = ShadingPattern(rawValue: val)
            }
            props.shading = shading
        }

        return props
    }

    // MARK: - Styles Parsing

    private static func parseStyles(from xml: XMLDocument) throws -> [Style] {
        var styles: [Style] = []

        let styleNodes = try xml.nodes(forXPath: "//*[local-name()='style']")

        for node in styleNodes {
            guard let element = node as? XMLElement else { continue }

            guard let styleId = element.attribute(forName: "w:styleId")?.stringValue else { continue }
            guard let typeStr = element.attribute(forName: "w:type")?.stringValue,
                  let type = StyleType(rawValue: typeStr) else { continue }

            var name = styleId
            if let nameElement = element.elements(forName: "w:name").first,
               let val = nameElement.attribute(forName: "w:val")?.stringValue {
                name = val
            }

            var style = Style(id: styleId, name: name, type: type)

            // 基於
            if let basedOn = element.elements(forName: "w:basedOn").first,
               let val = basedOn.attribute(forName: "w:val")?.stringValue {
                style.basedOn = val
            }

            // 下一樣式
            if let next = element.elements(forName: "w:next").first,
               let val = next.attribute(forName: "w:val")?.stringValue {
                style.nextStyle = val
            }

            // 預設
            if element.attribute(forName: "w:default")?.stringValue == "1" {
                style.isDefault = true
            }

            // 快速樣式
            style.isQuickStyle = element.elements(forName: "w:qFormat").first != nil

            // 段落屬性
            if let pPr = element.elements(forName: "w:pPr").first {
                style.paragraphProperties = parseParagraphProperties(from: pPr)
            }

            // Run 屬性
            if let rPr = element.elements(forName: "w:rPr").first {
                style.runProperties = parseRunProperties(from: rPr)
            }

            styles.append(style)
        }

        // 如果沒有讀到樣式，使用預設樣式
        if styles.isEmpty {
            styles = Style.defaultStyles
        }

        return styles
    }

    // MARK: - Core Properties Parsing

    private static func parseCoreProperties(from xml: XMLDocument) throws -> DocumentProperties {
        var props = DocumentProperties()

        // 標題
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='title']"),
           let node = nodes.first {
            props.title = node.stringValue
        }

        // 主題
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='subject']"),
           let node = nodes.first {
            props.subject = node.stringValue
        }

        // 作者
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='creator']"),
           let node = nodes.first {
            props.creator = node.stringValue
        }

        // 關鍵字
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='keywords']"),
           let node = nodes.first {
            props.keywords = node.stringValue
        }

        // 描述
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='description']"),
           let node = nodes.first {
            props.description = node.stringValue
        }

        // 最後修改者
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='lastModifiedBy']"),
           let node = nodes.first {
            props.lastModifiedBy = node.stringValue
        }

        // 版本
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='revision']"),
           let node = nodes.first,
           let value = node.stringValue {
            props.revision = Int(value)
        }

        // 建立日期
        let dateFormatter = ISO8601DateFormatter()
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='created']"),
           let node = nodes.first,
           let value = node.stringValue {
            props.created = dateFormatter.date(from: value)
        }

        // 修改日期
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='modified']"),
           let node = nodes.first,
           let value = node.stringValue {
            props.modified = dateFormatter.date(from: value)
        }

        return props
    }

    // MARK: - Numbering Parsing

    /// 解析 numbering.xml
    private static func parseNumbering(from xml: XMLDocument) throws -> Numbering {
        var numbering = Numbering()

        // 解析抽象編號定義 (w:abstractNum)
        let abstractNumNodes = try xml.nodes(forXPath: "//*[local-name()='abstractNum']")
        for node in abstractNumNodes {
            guard let element = node as? XMLElement,
                  let abstractNumIdStr = element.attribute(forName: "w:abstractNumId")?.stringValue,
                  let abstractNumId = Int(abstractNumIdStr) else { continue }

            var levels: [Level] = []

            // 解析層級 (w:lvl)
            for lvlElement in element.elements(forName: "w:lvl") {
                guard let ilvlStr = lvlElement.attribute(forName: "w:ilvl")?.stringValue,
                      let ilvl = Int(ilvlStr) else { continue }

                var numFmt: NumberFormat = .decimal
                var lvlText = ""
                var start = 1
                var indent = 720  // 預設縮排
                var fontName: String?

                // 編號格式 (w:numFmt)
                if let numFmtEl = lvlElement.elements(forName: "w:numFmt").first,
                   let val = numFmtEl.attribute(forName: "w:val")?.stringValue {
                    numFmt = NumberFormat(rawValue: val) ?? .decimal
                }

                // 文字格式 (w:lvlText)
                if let lvlTextEl = lvlElement.elements(forName: "w:lvlText").first,
                   let val = lvlTextEl.attribute(forName: "w:val")?.stringValue {
                    lvlText = val
                }

                // 起始值 (w:start)
                if let startEl = lvlElement.elements(forName: "w:start").first,
                   let val = startEl.attribute(forName: "w:val")?.stringValue {
                    start = Int(val) ?? 1
                }

                // 縮排 (w:pPr/w:ind)
                if let pPr = lvlElement.elements(forName: "w:pPr").first,
                   let ind = pPr.elements(forName: "w:ind").first,
                   let left = ind.attribute(forName: "w:left")?.stringValue {
                    indent = Int(left) ?? 720
                }

                // 字型 (w:rPr/w:rFonts)
                if let rPr = lvlElement.elements(forName: "w:rPr").first,
                   let rFonts = rPr.elements(forName: "w:rFonts").first,
                   let ascii = rFonts.attribute(forName: "w:ascii")?.stringValue {
                    fontName = ascii
                }

                let level = Level(
                    ilvl: ilvl,
                    start: start,
                    numFmt: numFmt,
                    lvlText: lvlText,
                    indent: indent,
                    fontName: fontName
                )
                levels.append(level)
            }

            let abstractNum = AbstractNum(abstractNumId: abstractNumId, levels: levels)
            numbering.abstractNums.append(abstractNum)
        }

        // 解析編號實例 (w:num)
        let numNodes = try xml.nodes(forXPath: "//*[local-name()='num']")
        for node in numNodes {
            guard let element = node as? XMLElement,
                  let numIdStr = element.attribute(forName: "w:numId")?.stringValue,
                  let numId = Int(numIdStr) else { continue }

            // 取得對應的 abstractNumId
            guard let abstractNumIdRef = element.elements(forName: "w:abstractNumId").first,
                  let abstractNumIdStr = abstractNumIdRef.attribute(forName: "w:val")?.stringValue,
                  let abstractNumId = Int(abstractNumIdStr) else { continue }

            let num = Num(numId: numId, abstractNumId: abstractNumId)
            numbering.nums.append(num)
        }

        return numbering
    }

    // MARK: - Semantic Detection

    /// 偵測段落的語義類型
    private static func detectParagraphSemantic(
        properties: ParagraphProperties,
        runs: [Run],
        styles: [Style],
        numbering: Numbering
    ) -> SemanticAnnotation? {
        // 1. 檢查標題樣式
        if let styleName = properties.style {
            if let headingLevel = detectHeadingLevel(styleName: styleName, styles: styles) {
                return SemanticAnnotation.heading(headingLevel)
            }

            // 檢查 Title/Subtitle
            let lowerStyle = styleName.lowercased()
            if lowerStyle == "title" || lowerStyle.contains("title") {
                return SemanticAnnotation(type: .title)
            }
            if lowerStyle == "subtitle" || lowerStyle.contains("subtitle") {
                return SemanticAnnotation(type: .subtitle)
            }
        }

        // 2. 檢查編號/項目符號
        if let numInfo = properties.numbering {
            let isBullet = isBulletList(numId: numInfo.numId, numbering: numbering)
            if isBullet {
                return SemanticAnnotation.bulletItem(level: numInfo.level)
            } else {
                return SemanticAnnotation.numberedItem(level: numInfo.level)
            }
        }

        // 3. 檢查分頁符
        if properties.pageBreakBefore {
            return SemanticAnnotation.pageBreak
        }

        // 4. 檢查 runs 中是否有公式或圖片（段落級別標註）
        for run in runs {
            // 有 OMML 公式
            if let rawXML = run.rawXML, rawXML.contains("oMath") {
                return SemanticAnnotation.ommlFormula
            }
            // 有圖片
            if run.drawing != nil {
                return SemanticAnnotation.unknownImage
            }
        }

        // 5. 預設為一般段落
        return SemanticAnnotation.paragraph
    }

    /// 從樣式名稱偵測標題層級
    private static func detectHeadingLevel(styleName: String, styles: [Style]) -> Int? {
        let lowerName = styleName.lowercased()

        // 直接比對常見標題樣式 ID
        // Word 預設: Heading1, Heading2, ... Heading9
        // 或中文: 標題1, 標題2, ...
        if lowerName.hasPrefix("heading") {
            let numPart = lowerName.dropFirst("heading".count)
            if let level = Int(numPart), level >= 1, level <= 9 {
                return level
            }
        }

        // 檢查樣式定義中的 name
        if let style = styles.first(where: { $0.id == styleName }) {
            let displayName = style.name.lowercased()
            if displayName.hasPrefix("heading") {
                let numPart = displayName.dropFirst("heading".count).trimmingCharacters(in: .whitespaces)
                if let level = Int(numPart), level >= 1, level <= 9 {
                    return level
                }
            }
            // 檢查 basedOn 是否為標題樣式
            if let basedOn = style.basedOn {
                return detectHeadingLevel(styleName: basedOn, styles: styles)
            }
        }

        return nil
    }

    /// 判斷是否為項目符號清單
    private static func isBulletList(numId: Int, numbering: Numbering) -> Bool {
        // 找到對應的 numbering instance (Num)
        guard let num = numbering.nums.first(where: { $0.numId == numId }) else {
            return false
        }

        // 找到對應的 abstract numbering (AbstractNum)
        guard let abstractNum = numbering.abstractNums.first(where: { $0.abstractNumId == num.abstractNumId }) else {
            return false
        }

        // 檢查第一層的格式
        if let firstLevel = abstractNum.levels.first {
            // bullet 格式通常是 .bullet 或文字是符號
            if firstLevel.numFmt == .bullet {
                return true
            }
            // 檢查文字是否為符號（如 •、○、■ 等）
            let text = firstLevel.lvlText
            let bulletSymbols = ["•", "○", "■", "□", "◆", "◇", "▪", "▫", "●", "○", "\u{F0B7}", "\u{F0A7}"]
            for symbol in bulletSymbols {
                if text.contains(symbol) {
                    return true
                }
            }
        }

        return false
    }

    // MARK: - Comments Extended Parsing

    /// 解析 commentsExtended.xml（Word 2012+ 回覆與已解決狀態）
    private static func parseCommentsExtended(from xml: XMLDocument, into comments: inout CommentsCollection) throws {
        // commentsExtended.xml 的結構：
        // <w15:commentsEx>
        //   <w15:commentEx w15:paraId="..." w15:paraIdParent="..." w15:done="1"/>
        // </w15:commentsEx>
        let extNodes = try xml.nodes(forXPath: "//*[local-name()='commentEx']")

        for node in extNodes {
            guard let element = node as? XMLElement else { continue }

            // 取得 paraId
            let paraId = element.attribute(forName: "w15:paraId")?.stringValue
                ?? element.attribute(forName: "paraId")?.stringValue
            guard let paraId = paraId else { continue }

            // 找到對應的 comment（透過 paraId）
            guard let idx = comments.comments.firstIndex(where: { $0.paraId == paraId }) else { continue }

            // 解析 parentId（回覆）
            let parentParaId = element.attribute(forName: "w15:paraIdParent")?.stringValue
                ?? element.attribute(forName: "paraIdParent")?.stringValue
            if let parentParaId = parentParaId,
               let parentComment = comments.comments.first(where: { $0.paraId == parentParaId }) {
                comments.comments[idx].parentId = parentComment.id
            }

            // 解析 done（已解決）
            let doneStr = element.attribute(forName: "w15:done")?.stringValue
                ?? element.attribute(forName: "done")?.stringValue
            if doneStr == "1" {
                comments.comments[idx].done = true
            }
        }
    }

    // MARK: - Comments Parsing

    private static func parseComments(from xml: XMLDocument) throws -> CommentsCollection {
        var collection = CommentsCollection()

        // 取得所有註解節點
        let commentNodes = try xml.nodes(forXPath: "//*[local-name()='comment']")

        for node in commentNodes {
            guard let element = node as? XMLElement else { continue }

            // 解析註解 ID
            guard let idStr = element.attribute(forName: "w:id")?.stringValue,
                  let id = Int(idStr) else { continue }

            // 解析作者
            let author = element.attribute(forName: "w:author")?.stringValue ?? "Unknown"

            // 解析縮寫
            let initials = element.attribute(forName: "w:initials")?.stringValue

            // 解析日期
            let dateFormatter = ISO8601DateFormatter()
            var date = Date()
            if let dateStr = element.attribute(forName: "w:date")?.stringValue {
                date = dateFormatter.date(from: dateStr) ?? Date()
            }

            // 解析註解文字（從 w:p/w:r/w:t 取得）
            var text = ""
            let textNodes = try element.nodes(forXPath: ".//*[local-name()='t']")
            for textNode in textNodes {
                text += textNode.stringValue ?? ""
            }

            // 建立 Comment 物件
            // 注意：從 comments.xml 讀取時，paragraphIndex 需要從文件中的 commentRangeStart 來確定
            // 這裡先設為 -1，表示需要從文件內容對應
            var comment = Comment(
                id: id,
                author: author,
                text: text.trimmingCharacters(in: .whitespacesAndNewlines),
                paragraphIndex: -1,
                date: date,
                initials: initials
            )

            // 嘗試解析 w14:paraId（用於回覆連結）
            // 從段落屬性中取得
            if let pElement = element.elements(forName: "w:p").first {
                // w14:paraId 可能在段落屬性中
                if let paraIdAttr = pElement.attribute(forName: "w14:paraId")?.stringValue {
                    comment.paraId = paraIdAttr
                }
            }

            collection.comments.append(comment)
        }

        return collection
    }

    // MARK: - XPath-free Helpers

    /// 在子孫中搜尋第一個符合 localName 的元素（取代 XPath，避免 O(n²)）
    private static func findFirstDescendant(of node: XMLNode, localNames: [String]) -> XMLElement? {
        guard let children = node.children else { return nil }
        for child in children {
            guard let el = child as? XMLElement else { continue }
            if let name = el.localName, localNames.contains(name) {
                return el
            }
            if let found = findFirstDescendant(of: el, localNames: localNames) {
                return found
            }
        }
        return nil
    }
}
