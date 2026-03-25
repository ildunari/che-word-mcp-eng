import Foundation

private struct EditableTrackedSegment {
    let sourceTrackedIndex: Int
    let run: Run
    let revision: Revision?
    let start: Int
    let end: Int
}

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
        let clampedIndex = visibleParagraphInsertionStorageIndex(for: index)
        body.children.insert(.paragraph(trackedParagraph), at: clampedIndex)
    }

    mutating func trackUpdatedParagraph(at index: Int, text: String) throws {
        let actualIndex = try visibleParagraphStorageIndex(for: index)
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
        let actualIndex = try visibleParagraphStorageIndex(for: index)
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
        let actualIndex = try visibleParagraphStorageIndex(for: index)
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
        let actualIndex = try visibleParagraphStorageIndex(for: index)
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
        if paragraph.trackedRuns?.isEmpty == true {
            paragraph.trackedRuns = nil
        }
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

        let expectedPartnerId: Int
        let expectedPartnerType: RevisionType

        switch revision.type {
        case .deletion:
            expectedPartnerId = revision.id + 1
            expectedPartnerType = .insertion
        case .insertion:
            expectedPartnerId = revision.id - 1
            expectedPartnerType = .deletion
        default:
            return [revision.id]
        }

        guard let partner = revisions.revisions.first(where: { candidate in
            candidate.id == expectedPartnerId
                && candidate.paragraphIndex == revision.paragraphIndex
                && candidate.type == expectedPartnerType
        }) else {
            return [revision.id]
        }

        return Set([revision.id, partner.id])
    }

    public mutating func trackReplaceTextRange(
        at paragraphIndex: Int,
        start: Int,
        end: Int,
        replacement: String,
        replacementProperties: RunProperties?
    ) throws {
        let actualIndex = try visibleParagraphStorageIndex(for: paragraphIndex)
        guard case .paragraph(var paragraph) = body.children[actualIndex] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let sourceTrackedRuns = paragraph.trackedRuns ?? paragraph.runs.map { TrackedRun(run: $0) }
        let segments = splitEditableTrackedSegments(for: paragraph, boundaries: [start, end])
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
        let touchedSegments = segments.filter { $0.start < end && $0.end > start }
        let segmentsBySourceIndex = Dictionary(grouping: segments, by: \.sourceTrackedIndex)
        var trackedRuns: [TrackedRun] = []
        var inserted = false

        for (sourceIndex, trackedRun) in sourceTrackedRuns.enumerated() {
            if trackedRun.isDeleted {
                trackedRuns.append(trackedRun)
                continue
            }

            let sourceSegments = segmentsBySourceIndex[sourceIndex] ?? []
            for segment in sourceSegments {
                if segment.end <= start {
                    trackedRuns.append(TrackedRun(run: segment.run, revision: segment.revision))
                    continue
                }

                if segment.start >= end {
                    if !inserted {
                        if start < end {
                            for deletedSegment in touchedSegments {
                                trackedRuns.append(
                                    TrackedRun(
                                        run: deletedSegment.run,
                                        revision: deletionRevision,
                                        isDeleted: true
                                    )
                                )
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
                    trackedRuns.append(TrackedRun(run: segment.run, revision: segment.revision))
                }
            }
        }

        if !inserted, let insertionRevision, !replacement.isEmpty {
            if start < end {
                for deletedSegment in touchedSegments {
                    trackedRuns.append(
                        TrackedRun(
                            run: deletedSegment.run,
                            revision: deletionRevision,
                            isDeleted: true
                        )
                    )
                }
            }
            trackedRuns.append(
                TrackedRun(
                    run: Run(text: replacement, properties: inheritedProperties),
                    revision: insertionRevision
                )
            )
        } else if !inserted, start < end {
            for deletedSegment in touchedSegments {
                trackedRuns.append(
                    TrackedRun(
                        run: deletedSegment.run,
                        revision: deletionRevision,
                        isDeleted: true
                    )
                )
            }
        }

        let normalizedTrackedRuns = normalizedTrackedRunsForInlineEditing(trackedRuns)
        paragraph.trackedRuns = normalizedTrackedRuns
        paragraph.runs = mergeRuns(normalizedTrackedRuns.compactMap { trackedRun in
            trackedRun.isDeleted ? nil : trackedRun.run
        })
        paragraph = normalizeTrackedParagraph(paragraph)
        body.children[actualIndex] = .paragraph(paragraph)
    }

    public mutating func trackFormatTextRange(
        at paragraphIndex: Int,
        start: Int,
        end: Int,
        format: RunProperties
    ) throws {
        let actualIndex = try visibleParagraphStorageIndex(for: paragraphIndex)
        guard case .paragraph(var paragraph) = body.children[actualIndex] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let sourceTrackedRuns = paragraph.trackedRuns ?? paragraph.runs.map { TrackedRun(run: $0) }
        let segments = splitEditableTrackedSegments(for: paragraph, boundaries: [start, end])
        let paragraphLength = segments.last?.end ?? 0
        guard start >= 0, end >= start, end <= paragraphLength else {
            throw WordError.invalidParameter("range", "Invalid range \(start)..<\(end)")
        }
        guard start < end else { return }

        let revision = makeRevision(type: .formatChange, paragraphIndex: paragraphIndex)
        let segmentsBySourceIndex = Dictionary(grouping: segments, by: \.sourceTrackedIndex)
        var updatedTrackedRuns: [TrackedRun] = []

        for (sourceIndex, trackedRun) in sourceTrackedRuns.enumerated() {
            if trackedRun.isDeleted {
                updatedTrackedRuns.append(trackedRun)
                continue
            }

            let sourceSegments = segmentsBySourceIndex[sourceIndex] ?? []
            for segment in sourceSegments {
                var updatedRun = segment.run
                if segment.start < end, segment.end > start {
                    let previous = RunPropertiesSnapshot(from: updatedRun.properties)
                    updatedRun.properties.merge(with: format)
                    updatedRun.properties.formatChange = RunFormatChange(
                        revision: revision,
                        previousProperties: previous
                    )
                }
                updatedTrackedRuns.append(
                    TrackedRun(
                        run: updatedRun,
                        revision: segment.revision,
                        isDeleted: false
                    )
                )
            }
        }
        paragraph.trackedRuns = updatedTrackedRuns
        paragraph.runs = mergeRuns(updatedTrackedRuns.compactMap { trackedRun in
            trackedRun.isDeleted ? nil : trackedRun.run
        })
        body.children[actualIndex] = .paragraph(paragraph)
    }

    private func splitEditableTrackedSegments(for paragraph: Paragraph, boundaries: [Int]) -> [EditableTrackedSegment] {
        var visibleSegments: [(sourceTrackedIndex: Int, run: Run, revision: Revision?)] =
            (paragraph.trackedRuns ?? paragraph.runs.map { TrackedRun(run: $0) })
            .enumerated()
            .compactMap { index, trackedRun in
                guard !trackedRun.isDeleted, !trackedRun.run.text.isEmpty else { return nil }
                return (index, trackedRun.run, trackedRun.revision)
            }

        for boundary in Array(Set(boundaries)).sorted() {
            var cursor = 0
            for index in visibleSegments.indices {
                let segment = visibleSegments[index]
                let start = cursor
                let end = cursor + segment.run.text.count
                cursor = end

                guard boundary > start, boundary < end else { continue }

                let splitIndex = segment.run.text.index(segment.run.text.startIndex, offsetBy: boundary - start)
                let leftText = String(segment.run.text[..<splitIndex])
                let rightText = String(segment.run.text[splitIndex...])
                let leftRun = Run(text: leftText, properties: segment.run.properties)
                let rightRun = Run(text: rightText, properties: segment.run.properties)
                visibleSegments.remove(at: index)
                visibleSegments.insert(
                    contentsOf: [
                        (segment.sourceTrackedIndex, leftRun, segment.revision),
                        (segment.sourceTrackedIndex, rightRun, segment.revision),
                    ],
                    at: index
                )
                break
            }
        }

        var cursor = 0
        return visibleSegments.map { segment in
            let start = cursor
            cursor += segment.run.text.count
            return EditableTrackedSegment(
                sourceTrackedIndex: segment.sourceTrackedIndex,
                run: segment.run,
                revision: segment.revision,
                start: start,
                end: cursor
            )
        }
    }

    private func visibleText(in segments: [EditableTrackedSegment], from start: Int, to end: Int) -> String {
        segments.compactMap { segment in
            guard segment.start < end, segment.end > start else { return nil }
            return segment.run.text
        }.joined()
    }

    private func inheritedRunProperties(in segments: [EditableTrackedSegment], start: Int, end: Int) -> RunProperties {
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

    private func normalizedTrackedRunsForInlineEditing(_ trackedRuns: [TrackedRun]) -> [TrackedRun] {
        var normalized: [TrackedRun] = []
        var pendingWhitespace = ""
        var index = 0

        func isStandaloneWhitespace(_ trackedRun: TrackedRun) -> Bool {
            !trackedRun.isDeleted
                && trackedRun.revision == nil
                && !trackedRun.run.text.isEmpty
                && trackedRun.run.text.allSatisfy(\.isWhitespace)
                && trackedRun.run.rawXML == nil
                && trackedRun.run.properties.rawXML == nil
                && trackedRun.run.drawing == nil
        }

        while index < trackedRuns.count {
            let trackedRun = trackedRuns[index]
            if isStandaloneWhitespace(trackedRun) {
                pendingWhitespace += trackedRun.run.text
                index += 1
                continue
            }

            var adjusted = trackedRun
            if !pendingWhitespace.isEmpty {
                adjusted.run.text = pendingWhitespace + adjusted.run.text
                if adjusted.isDeleted {
                    normalized.append(adjusted)
                    index += 1
                    while index < trackedRuns.count {
                        var follower = trackedRuns[index]
                        follower.run.text = pendingWhitespace + follower.run.text
                        normalized.append(follower)
                        index += 1
                        if !follower.isDeleted {
                            break
                        }
                    }
                    pendingWhitespace = ""
                    continue
                }
                pendingWhitespace = ""
            }

            normalized.append(adjusted)
            index += 1
        }

        if !pendingWhitespace.isEmpty {
            if var last = normalized.last {
                last.run.text += pendingWhitespace
                normalized[normalized.count - 1] = last
            } else {
                normalized.append(TrackedRun(run: Run(text: pendingWhitespace)))
            }
        }

        return normalized
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
                if updatedRuns.isEmpty && paragraph.runs.isEmpty {
                    body.children.remove(at: index)
                    return
                }
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
                if updatedRuns.isEmpty && paragraph.runs.isEmpty {
                    body.children.remove(at: index)
                    return
                }
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
