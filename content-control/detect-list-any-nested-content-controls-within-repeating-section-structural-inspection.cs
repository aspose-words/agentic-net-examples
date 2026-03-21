using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a repeating section content control directly into the document body.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "MyRepeatingSection",
            Tag = "RepeatingSectionTag"
        };
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Inside the repeating section, add a paragraph to host nested controls.
        Paragraph para = new Paragraph(doc);
        repeatingSection.AppendChild(para);
        builder.MoveTo(para);

        // Add a nested rich text content control.
        StructuredDocumentTag nestedRichText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline)
        {
            Title = "NestedRichText",
            Tag = "RichTextTag"
        };
        para.AppendChild(nestedRichText);
        builder.MoveTo(nestedRichText);
        builder.Writeln("Sample text inside nested rich text control.");

        // Add another nested plain text content control.
        StructuredDocumentTag nestedPlain = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NestedPlainText",
            Tag = "PlainTextTag"
        };
        para.AppendChild(nestedPlain);
        builder.MoveTo(nestedPlain);
        builder.Writeln("Plain text inside nested plain text control.");

        // Retrieve all structured document tags (content controls) in the document.
        NodeCollection allSdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control that is a repeating section.
        foreach (StructuredDocumentTag repeating in allSdtNodes
                     .OfType<StructuredDocumentTag>()
                     .Where(sdt => sdt.SdtType == SdtType.RepeatingSection))
        {
            Console.WriteLine($"Repeating Section found. Title: '{repeating.Title}', Tag: '{repeating.Tag}'");

            // Get all nested content controls inside this repeating section (deep search).
            NodeCollection nestedControls = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true);

            foreach (StructuredDocumentTag nested in nestedControls
                         .OfType<StructuredDocumentTag>()
                         .Where(sdt => sdt != repeating))
            {
                Console.WriteLine($"  Nested Control - Type: {nested.SdtType}, Title: '{nested.Title}', Tag: '{nested.Tag}'");
            }

            Console.WriteLine(); // Blank line for readability between sections.
        }

        // Optionally, save the document after inspection.
        doc.Save("OutputDocument.docx");
    }
}
