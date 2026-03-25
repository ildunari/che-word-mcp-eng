import Foundation
import OOXMLSwift

enum InlineEditingError: Error, LocalizedError, Equatable {
    case invalidRange(start: Int, end: Int, length: Int)
    case unsupportedParagraphContent(String)

    var errorDescription: String? {
        switch self {
        case .invalidRange(let start, let end, let length):
            return "Invalid text range start=\(start) end=\(end) for visible paragraph length \(length)"
        case .unsupportedParagraphContent(let reason):
            return "Inline editing is not supported for this paragraph: \(reason)"
        }
    }
}

private struct InlineRunSegment {
    let runIndex: Int
    let start: Int
    let end: Int
    let run: Run

    var length: Int { end - start }
    var isTextOnly: Bool {
        run.rawXML == nil && run.properties.rawXML == nil && run.drawing == nil
    }
}

extension WordDocument {
    mutating func replaceTextRange(
        at paragraphIndex: Int,
        start: Int,
        end: Int,
        replacement: String,
        replacementProperties: RunProperties?
    ) throws {
        let actualIndex = try paragraphStorageIndex(for: paragraphIndex)
        guard case .paragraph(var paragraph) = body.children[actualIndex] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        try validateEditableParagraph(paragraph, start: start, end: end)
        if isTrackChangesEnabled() {
            try trackReplaceTextRange(
                at: paragraphIndex,
                start: start,
                end: end,
                replacement: replacement,
                replacementProperties: replacementProperties
            )
            return
        }
        paragraph = try paragraphBySplittingRuns(paragraph, at: [start, end])

        let alignedSegments = segments(for: paragraph)
        let inheritedProperties = replacementProperties ?? inheritedRunProperties(
            forRangeStart: start,
            end: end,
            in: alignedSegments
        )

        var newRuns: [Run] = []
        var insertedReplacement = false

        for segment in alignedSegments {
            if segment.end <= start {
                newRuns.append(segment.run)
                continue
            }

            if !insertedReplacement {
                if !replacement.isEmpty {
                    newRuns.append(Run(text: replacement, properties: inheritedProperties))
                }
                insertedReplacement = true
            }

            if segment.start >= end {
                newRuns.append(segment.run)
            }
        }

        if !insertedReplacement, !replacement.isEmpty {
            newRuns.append(Run(text: replacement, properties: inheritedProperties))
        }

        if replacement.isEmpty {
            paragraph.runs = mergedRuns(newRuns)
        } else {
            paragraph.runs = newRuns.filter { !$0.text.isEmpty }
        }
        body.children[actualIndex] = .paragraph(paragraph)
    }

    mutating func formatTextRange(
        at paragraphIndex: Int,
        start: Int,
        end: Int,
        format: RunProperties
    ) throws {
        let actualIndex = try paragraphStorageIndex(for: paragraphIndex)
        guard case .paragraph(var paragraph) = body.children[actualIndex] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        try validateEditableParagraph(paragraph, start: start, end: end)
        if isTrackChangesEnabled() {
            try trackFormatTextRange(
                at: paragraphIndex,
                start: start,
                end: end,
                format: format
            )
            return
        }
        if start == end {
            return
        }
        paragraph = try paragraphBySplittingRuns(paragraph, at: [start, end])

        let alignedSegments = segments(for: paragraph)
        paragraph.runs = mergedRuns(alignedSegments.map { segment in
            guard segment.start < end, segment.end > start else {
                return segment.run
            }

            var updatedRun = segment.run
            updatedRun.properties = mergedProperties(base: updatedRun.properties, override: format)
            return updatedRun
        })
        body.children[actualIndex] = .paragraph(paragraph)
    }

    private func paragraphStorageIndex(for paragraphIndex: Int) throws -> Int {
        try visibleParagraphStorageIndex(for: paragraphIndex)
    }

    private func validateEditableParagraph(_ paragraph: Paragraph, start: Int, end: Int) throws {
        if !paragraph.hyperlinks.isEmpty {
            throw InlineEditingError.unsupportedParagraphContent("paragraph contains hyperlinks")
        }

        let visibleRuns = paragraph.trackedRuns?.compactMap { trackedRun in
            trackedRun.isDeleted ? nil : trackedRun.run
        } ?? paragraph.runs
        let paragraphLength = visibleRuns.reduce(into: 0) { partialResult, run in
            partialResult += run.text.count
        }

        guard start >= 0, end >= start, end <= paragraphLength else {
            throw InlineEditingError.invalidRange(start: start, end: end, length: paragraphLength)
        }

        let visibleSegments = visibleRuns.enumerated().reduce(into: [InlineRunSegment]()) { partialResult, item in
            let (index, run) = item
            let startOffset = partialResult.last?.end ?? 0
            let endOffset = startOffset + run.text.count
            partialResult.append(InlineRunSegment(runIndex: index, start: startOffset, end: endOffset, run: run))
        }

        for segment in visibleSegments where !segment.isTextOnly {
            if start == end {
                if start > segment.start, start < segment.end {
                    throw InlineEditingError.unsupportedParagraphContent("edit would split a non-text run")
                }
            } else if start < segment.end, end > segment.start {
                throw InlineEditingError.unsupportedParagraphContent("edit would split a non-text run")
            }
        }
    }

    private func paragraphBySplittingRuns(_ paragraph: Paragraph, at boundaries: [Int]) throws -> Paragraph {
        var updatedParagraph = paragraph
        let sortedBoundaries = Array(Set(boundaries)).sorted()

        for boundary in sortedBoundaries {
            let currentSegments = segments(for: updatedParagraph)
            guard let segment = currentSegments.first(where: { boundary > $0.start && boundary < $0.end }) else {
                continue
            }

            guard segment.isTextOnly else {
                throw InlineEditingError.unsupportedParagraphContent("edit would split a non-text run")
            }

            let offset = boundary - segment.start
            let splitRuns = splitTextRun(segment.run, at: offset)

            updatedParagraph.runs.remove(at: segment.runIndex)
            updatedParagraph.runs.insert(contentsOf: splitRuns, at: segment.runIndex)
        }

        return updatedParagraph
    }

    private func splitTextRun(_ run: Run, at offset: Int) -> [Run] {
        guard offset > 0, offset < run.text.count else {
            return [run]
        }

        let splitIndex = stringIndex(in: run.text, offset: offset)
        let leftText = String(run.text[..<splitIndex])
        let rightText = String(run.text[splitIndex...])

        return [
            Run(text: leftText, properties: run.properties),
            Run(text: rightText, properties: run.properties),
        ]
    }

    private func stringIndex(in text: String, offset: Int) -> String.Index {
        text.index(text.startIndex, offsetBy: offset)
    }

    private func segments(for paragraph: Paragraph) -> [InlineRunSegment] {
        var cursor = 0
        return paragraph.runs.enumerated().map { index, run in
            let start = cursor
            cursor += run.text.count
            return InlineRunSegment(runIndex: index, start: start, end: cursor, run: run)
        }
    }

    private func inheritedRunProperties(forRangeStart start: Int, end: Int, in segments: [InlineRunSegment]) -> RunProperties {
        if start < end, let segment = segments.first(where: { start < $0.end && end > $0.start }) {
            return segment.run.properties
        }

        if start == end, let lastSegment = segments.last(where: { $0.end == start }) {
            return lastSegment.run.properties
        }

        if let segment = segments.first(where: { start >= $0.start && start < $0.end }) {
            return segment.run.properties
        }

        if let segment = segments.first(where: { $0.start == start }) {
            return segment.run.properties
        }

        return RunProperties()
    }

    private func mergedProperties(base: RunProperties, override: RunProperties) -> RunProperties {
        var merged = base
        if override.bold { merged.bold = true }
        if override.italic { merged.italic = true }
        if let underline = override.underline { merged.underline = underline }
        if override.strikethrough { merged.strikethrough = true }
        if let fontSize = override.fontSize { merged.fontSize = fontSize }
        if let fontName = override.fontName { merged.fontName = fontName }
        if let color = override.color { merged.color = color }
        if override.clearHighlight {
            merged.highlight = nil
        } else if let highlight = override.highlight {
            merged.highlight = highlight
        }
        if let verticalAlign = override.verticalAlign { merged.verticalAlign = verticalAlign }
        if let characterSpacing = override.characterSpacing { merged.characterSpacing = characterSpacing }
        if let textEffect = override.textEffect { merged.textEffect = textEffect }
        if let rawXML = override.rawXML { merged.rawXML = rawXML }
        return merged
    }

    private func mergedRuns(_ runs: [Run]) -> [Run] {
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
}
