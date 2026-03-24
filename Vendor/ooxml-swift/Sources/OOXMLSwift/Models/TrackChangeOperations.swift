import Foundation

extension WordDocument {
    mutating func makeRevision(
        type: RevisionType,
        paragraphIndex: Int,
        originalText: String? = nil,
        newText: String? = nil,
        previousFormat: RunProperties? = nil
    ) -> Revision {
        let revision = Revision(
            id: revisions.nextRevisionId(),
            type: type,
            author: revisions.settings.author,
            paragraphIndex: paragraphIndex,
            originalText: originalText,
            newText: newText,
            date: Date()
        )
        revisions.revisions.append(revision)
        return revision
    }

    mutating func trackInsertedParagraph(_ paragraph: Paragraph, at index: Int) {
        let paragraphIndex = max(0, min(index, getParagraphs().count))
        var trackedParagraph = paragraph
        trackedParagraph.paragraphRevision = makeRevision(
            type: .insertion,
            paragraphIndex: paragraphIndex,
            newText: paragraph.getText()
        )
        let clampedIndex = min(max(0, index), body.children.count)
        body.children.insert(.paragraph(trackedParagraph), at: clampedIndex)
    }

    mutating func trackUpdatedParagraph(at index: Int, text: String) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            let deletedRevision = makeRevision(
                type: .deletion,
                paragraphIndex: index,
                originalText: para.getText()
            )
            let insertedRevision = makeRevision(
                type: .insertion,
                paragraphIndex: index,
                newText: text
            )
            let deletedRuns = (para.trackedRuns?.compactMap { $0.isDeleted ? $0 : TrackedRun(run: $0.run, revision: deletedRevision, isDeleted: true) }
                ?? para.runs.map { TrackedRun(run: $0, revision: deletedRevision, isDeleted: true) })
            let insertedRun = Run(text: text, properties: para.runs.last?.properties ?? RunProperties())
            para.trackedRuns = deletedRuns + [TrackedRun(run: insertedRun, revision: insertedRevision)]
            para.runs = [insertedRun]
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func trackDeletedParagraph(at index: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            let deletionRevision = makeRevision(
                type: .deletion,
                paragraphIndex: index,
                originalText: para.getText()
            )
            para.trackedRuns = (para.trackedRuns?.map { tracked in
                TrackedRun(run: tracked.run, revision: deletionRevision, isDeleted: true)
            } ?? para.runs.map { TrackedRun(run: $0, revision: deletionRevision, isDeleted: true) })
            para.runs = []
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func trackFormattedParagraph(at index: Int, with format: RunProperties) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            let revision = makeRevision(type: .formatChange, paragraphIndex: index)
            for runIndex in para.runs.indices {
                let previous = RunPropertiesSnapshot(from: para.runs[runIndex].properties)
                para.runs[runIndex].properties.merge(with: format)
                para.runs[runIndex].properties.formatChange = RunFormatChange(
                    revision: revision,
                    previousProperties: previous
                )
            }
            if let trackedRuns = para.trackedRuns {
                para.trackedRuns = trackedRuns.enumerated().map { offset, trackedRun in
                    guard offset < para.runs.count else { return trackedRun }
                    return TrackedRun(run: para.runs[offset], revision: trackedRun.revision, isDeleted: trackedRun.isDeleted)
                }
            }
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func trackParagraphProperties(at index: Int, updates: ParagraphProperties) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            let previous = para.properties
            let revision = makeRevision(type: .paragraphChange, paragraphIndex: index)
            para.properties.merge(with: updates)
            para.properties.formatChange = ParagraphFormatChange(
                id: revision.id,
                author: revision.author,
                date: revision.date,
                previousProperties: ParagraphPropertiesSnapshot(from: previous)
            )
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func normalizeTrackedParagraph(_ paragraph: Paragraph) -> Paragraph {
        var paragraph = paragraph
        if let trackedRuns = paragraph.trackedRuns, trackedRuns.allSatisfy({ !$0.isDeleted && $0.revision == nil }) {
            paragraph.runs = trackedRuns.map(\.run)
            paragraph.trackedRuns = nil
        }
        return paragraph
    }

    func pairedRevisionIDs(for revision: Revision) -> Set<Int> {
        guard revision.type == .insertion || revision.type == .deletion else {
            return [revision.id]
        }

        let partners = revisions.revisions.filter { candidate in
            candidate.id != revision.id
                && candidate.paragraphIndex == revision.paragraphIndex
                && abs(candidate.id - revision.id) == 1
                && Set([candidate.type, revision.type]) == Set([.insertion, .deletion])
        }

        return Set([revision.id] + partners.map(\.id))
    }

    public mutating func trackReplaceTextRange(
        at paragraphIndex: Int,
        start: Int,
        end: Int,
        replacement: String,
        replacementProperties: RunProperties?
    ) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }
        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        guard case .paragraph(var paragraph) = body.children[actualIndex] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let segments = splitVisibleSegments(for: paragraph, boundaries: [start, end])
        let paragraphLength = segments.last?.end ?? 0
        guard start >= 0, end >= start, end <= paragraphLength else {
            throw WordError.invalidParameter("range", "Invalid range \(start)..<\(end)")
        }

        let deletionRevision = start < end ? makeRevision(
            type: .deletion,
            paragraphIndex: paragraphIndex,
            originalText: visibleText(in: segments, from: start, to: end)
        ) : nil
        let insertionRevision = !replacement.isEmpty ? makeRevision(
            type: .insertion,
            paragraphIndex: paragraphIndex,
            newText: replacement
        ) : nil

        let inheritedProperties = replacementProperties ?? inheritedRunProperties(in: segments, start: start, end: end)
        var trackedRuns: [TrackedRun] = []
        var inserted = false

        for segment in segments {
            if segment.end <= start {
                trackedRuns.append(TrackedRun(run: segment.run))
                continue
            }

            if !inserted {
                if start < end {
                    let deletedSegments = segments.filter { $0.start < end && $0.end > start }
                    for deletedSegment in deletedSegments {
                        trackedRuns.append(TrackedRun(run: deletedSegment.run, revision: deletionRevision, isDeleted: true))
                    }
                }
                if let insertionRevision, !replacement.isEmpty {
                    trackedRuns.append(
                        TrackedRun(
                            run: Run(text: replacement, properties: inheritedProperties),
                            revision: insertionRevision
                        )
                    )
                }
                inserted = true
            }

            if segment.start >= end {
                trackedRuns.append(TrackedRun(run: segment.run))
            }
        }

        if !inserted, let insertionRevision, !replacement.isEmpty {
            trackedRuns.append(
                TrackedRun(
                    run: Run(text: replacement, properties: inheritedProperties),
                    revision: insertionRevision
                )
            )
        }

        paragraph.trackedRuns = trackedRuns
        paragraph = normalizeTrackedParagraph(paragraph)
        body.children[actualIndex] = .paragraph(paragraph)
    }

    public mutating func trackFormatTextRange(
        at paragraphIndex: Int,
        start: Int,
        end: Int,
        format: RunProperties
    ) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }
        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        guard case .paragraph(var paragraph) = body.children[actualIndex] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let segments = splitVisibleSegments(for: paragraph, boundaries: [start, end])
        let paragraphLength = segments.last?.end ?? 0
        guard start >= 0, end >= start, end <= paragraphLength else {
            throw WordError.invalidParameter("range", "Invalid range \(start)..<\(end)")
        }
        guard start < end else { return }

        let revision = makeRevision(type: .formatChange, paragraphIndex: paragraphIndex)
        var updatedRuns: [Run] = []

        for segment in segments {
            guard segment.start < end, segment.end > start else {
                updatedRuns.append(segment.run)
                continue
            }

            var updatedRun = segment.run
            let previous = RunPropertiesSnapshot(from: updatedRun.properties)
            updatedRun.properties.merge(with: format)
            updatedRun.properties.formatChange = RunFormatChange(revision: revision, previousProperties: previous)
            updatedRuns.append(updatedRun)
        }

        paragraph.runs = mergeRuns(updatedRuns)
        if let trackedRuns = paragraph.trackedRuns, trackedRuns.count == paragraph.runs.count {
            paragraph.trackedRuns = zip(trackedRuns, paragraph.runs).map { tracked, run in
                TrackedRun(run: run, revision: tracked.revision, isDeleted: tracked.isDeleted)
            }
        }
        body.children[actualIndex] = .paragraph(paragraph)
    }

    private func splitVisibleSegments(for paragraph: Paragraph, boundaries: [Int]) -> [(run: Run, start: Int, end: Int)] {
        var visibleRuns = paragraph.trackedRuns?.compactMap { $0.isDeleted ? nil : $0.run } ?? paragraph.runs
        visibleRuns = visibleRuns.flatMap { run -> [Run] in
            run.text.isEmpty ? [] : [run]
        }
        for boundary in Array(Set(boundaries)).sorted() {
            var cursor = 0
            for index in visibleRuns.indices {
                let run = visibleRuns[index]
                let start = cursor
                let end = cursor + run.text.count
                cursor = end

                guard boundary > start, boundary < end else { continue }

                let splitIndex = run.text.index(run.text.startIndex, offsetBy: boundary - start)
                let leftText = String(run.text[..<splitIndex])
                let rightText = String(run.text[splitIndex...])
                let leftRun = Run(text: leftText, properties: run.properties)
                let rightRun = Run(text: rightText, properties: run.properties)
                visibleRuns.remove(at: index)
                visibleRuns.insert(contentsOf: [leftRun, rightRun], at: index)
                break
            }
        }
        var cursor = 0
        return visibleRuns.map { run in
            let start = cursor
            cursor += run.text.count
            return (run, start, cursor)
        }
    }

    private func visibleText(in segments: [(run: Run, start: Int, end: Int)], from start: Int, to end: Int) -> String {
        segments.compactMap { segment in
            guard segment.start < end, segment.end > start else { return nil }
            return segment.run.text
        }.joined()
    }

    private func inheritedRunProperties(in segments: [(run: Run, start: Int, end: Int)], start: Int, end: Int) -> RunProperties {
        if start < end, let segment = segments.first(where: { start < $0.end && end > $0.start }) {
            return segment.run.properties
        }
        if let segment = segments.last(where: { $0.end == start }) {
            return segment.run.properties
        }
        if let segment = segments.first(where: { start >= $0.start && start < $0.end }) {
            return segment.run.properties
        }
        return RunProperties()
    }

    private func mergeRuns(_ runs: [Run]) -> [Run] {
        var merged: [Run] = []
        for run in runs where !run.text.isEmpty {
            if var previous = merged.last,
               previous.properties == run.properties,
               previous.rawXML == nil,
               run.rawXML == nil,
               previous.properties.rawXML == nil,
               run.properties.rawXML == nil,
               previous.drawing == nil,
               run.drawing == nil {
                previous.text += run.text
                merged[merged.count - 1] = previous
            } else {
                merged.append(run)
            }
        }
        return merged
    }

    mutating func acceptNativeRevision(_ revision: Revision) {
        for index in body.children.indices {
            guard case .paragraph(var paragraph) = body.children[index] else { continue }

            if paragraph.paragraphRevision?.id == revision.id {
                paragraph.paragraphRevision = nil
                body.children[index] = .paragraph(normalizeTrackedParagraph(paragraph))
                continue
            }

            if let trackedRuns = paragraph.trackedRuns {
                var updatedRuns: [TrackedRun] = []
                for trackedRun in trackedRuns {
                    guard trackedRun.revision?.id == revision.id else {
                        updatedRuns.append(trackedRun)
                        continue
                    }

                    if trackedRun.isDeleted {
                        continue
                    }

                    var acceptedRun = trackedRun
                    acceptedRun.revision = nil
                    updatedRuns.append(acceptedRun)
                }
                paragraph.trackedRuns = updatedRuns
            }

            if paragraph.properties.formatChange?.id == revision.id {
                paragraph.properties.formatChange = nil
            }

            for runIndex in paragraph.runs.indices where paragraph.runs[runIndex].properties.formatChange?.revision.id == revision.id {
                paragraph.runs[runIndex].properties.formatChange = nil
            }

            paragraph.runs = mergeRuns(paragraph.runs)
            let normalized = normalizeTrackedParagraph(paragraph)
            if normalized.trackedRuns?.isEmpty == true && normalized.getText().isEmpty {
                body.children.remove(at: index)
                return
            }
            body.children[index] = .paragraph(normalized)
        }
    }

    mutating func rejectNativeRevision(_ revision: Revision) {
        for index in body.children.indices {
            guard case .paragraph(var paragraph) = body.children[index] else { continue }

            if paragraph.paragraphRevision?.id == revision.id, revision.type == .insertion {
                body.children.remove(at: index)
                return
            }

            if let trackedRuns = paragraph.trackedRuns {
                var updatedRuns: [TrackedRun] = []
                for trackedRun in trackedRuns {
                    guard trackedRun.revision?.id == revision.id else {
                        updatedRuns.append(trackedRun)
                        continue
                    }

                    if trackedRun.isDeleted {
                        var restoredRun = trackedRun
                        restoredRun.isDeleted = false
                        restoredRun.revision = nil
                        updatedRuns.append(restoredRun)
                    }
                }
                paragraph.trackedRuns = updatedRuns
            }

            if let formatChange = paragraph.properties.formatChange, formatChange.id == revision.id {
                formatChange.previousProperties.apply(to: &paragraph.properties)
            }

            for runIndex in paragraph.runs.indices {
                if let formatChange = paragraph.runs[runIndex].properties.formatChange, formatChange.revision.id == revision.id {
                    var restored = paragraph.runs[runIndex].properties
                    formatChange.previousProperties.apply(to: &restored)
                    restored.formatChange = nil
                    paragraph.runs[runIndex].properties = restored
                }
            }

            paragraph.runs = mergeRuns(paragraph.runs)
            let normalized = normalizeTrackedParagraph(paragraph)
            if normalized.trackedRuns?.isEmpty == true && normalized.getText().isEmpty {
                body.children.remove(at: index)
                return
            }
            body.children[index] = .paragraph(normalized)
        }
    }
}
