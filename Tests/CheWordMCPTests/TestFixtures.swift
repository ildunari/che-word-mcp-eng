import Foundation
import OOXMLSwift

enum TestFixtures {
    static func makeRun(
        _ text: String,
        bold: Bool = false,
        italic: Bool = false,
        underline: UnderlineType? = nil,
        color: String? = nil,
        rawXML: String? = nil,
        drawing: Drawing? = nil
    ) -> Run {
        var properties = RunProperties()
        properties.bold = bold
        properties.italic = italic
        properties.underline = underline
        properties.color = color
        properties.rawXML = rawXML

        var run = Run(text: text, properties: properties)
        run.rawXML = rawXML
        run.drawing = drawing
        return run
    }

    static func makeParagraph(
        runs: [Run],
        style: String? = nil,
        hyperlinks: [Hyperlink] = []
    ) -> Paragraph {
        var paragraph = Paragraph(runs: runs)
        paragraph.properties.style = style
        paragraph.hyperlinks = hyperlinks
        return paragraph
    }

    static func makeDocument(paragraphs: [Paragraph]) -> WordDocument {
        var document = WordDocument()
        for paragraph in paragraphs {
            document.appendParagraph(paragraph)
        }
        return document
    }
}
