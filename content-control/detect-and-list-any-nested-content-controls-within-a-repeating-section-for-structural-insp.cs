using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Build a repeating section with nested content controls.
        // -------------------------------------------------

        // Create the outer repeating section (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "OuterRepeatingSection",
            Tag = "outerRepeating"
        };
        // Insert the block‑level SDT into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Create a repeating section item inside the repeating section.
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block)
        {
            Title = "Item",
            Tag = "item"
        };
        repeatingSection.AppendChild(repeatingItem);

        // Add a paragraph to hold inline content controls.
        Paragraph paragraph = new Paragraph(doc);
        repeatingItem.AppendChild(paragraph);

        // Plain text content control (inline).
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainTextControl",
            Tag = "plain"
        };
        paragraph.AppendChild(plainTextSdt);
        Run plainRun = new Run(doc, "Plain text content");
        plainTextSdt.AppendChild(plainRun);

        // Rich text content control (inline).
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline)
        {
            Title = "RichTextControl",
            Tag = "rich"
        };
        paragraph.AppendChild(richTextSdt);
        Run richRun = new Run(doc, "Rich text content");
        richTextSdt.AppendChild(richRun);

        // Save the sample document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RepeatingSectionNested.docx");
        doc.Save(outputPath);

        // -------------------------------------------------
        // Detect and list nested content controls inside repeating sections.
        // -------------------------------------------------

        // Retrieve all StructuredDocumentTag nodes in the document.
        var allSdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                             .OfType<StructuredDocumentTag>();

        // Iterate over each repeating section.
        foreach (var repeating in allSdtNodes.Where(s => s.SdtType == SdtType.RepeatingSection))
        {
            Console.WriteLine($"Repeating Section: Title=\"{repeating.Title}\", Tag=\"{repeating.Tag}\"");

            // Find nested content controls within this repeating section (any depth),
            // excluding the repeating section itself and its item nodes.
            var nestedControls = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                          .OfType<StructuredDocumentTag>()
                                          .Where(s => s != repeating && s.SdtType != SdtType.RepeatingSectionItem);

            foreach (var nested in nestedControls)
            {
                Console.WriteLine($"  Nested Control -> Type: {nested.SdtType}, Title: \"{nested.Title}\", Tag: \"{nested.Tag}\"");
            }
        }
    }
}
