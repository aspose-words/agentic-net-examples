using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

namespace ContentControlExport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder for convenient text insertion.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading to demonstrate outline handling (optional).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Document with Content Controls");

            // ---------- Inline plain‑text content control ----------
            // Create an inline plain‑text StructuredDocumentTag.
            StructuredDocumentTag inlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "customer-name"
            };
            // Set the displayed text of the content control.
            inlineSdt.RemoveAllChildren();
            inlineSdt.AppendChild(new Run(doc, "Contoso Ltd."));

            // Insert the inline content control into the current paragraph.
            builder.InsertNode(inlineSdt);
            builder.Writeln(); // Move to a new paragraph.

            // ---------- Block‑level rich‑text content control ----------
            // Create a block‑level rich‑text StructuredDocumentTag.
            StructuredDocumentTag blockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "SectionContent",
                Tag = "section-content"
            };
            // Add a paragraph with some text inside the block content control.
            Paragraph blockParagraph = new Paragraph(doc);
            blockParagraph.AppendChild(new Run(doc, "This is a block‑level rich text content control."));
            blockSdt.AppendChild(blockParagraph);

            // Append the block content control to the document body.
            doc.FirstSection.Body.AppendChild(blockSdt);

            // ---------- Save the document as XPS preserving content control boundaries ----------
            // Create XpsSaveOptions; default settings preserve the structure of content controls.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Define output file path in the current working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControls.xps");

            // Save the document to XPS format.
            doc.Save(outputPath, xpsOptions);

            // Inform that the operation completed (optional console output).
            Console.WriteLine($"Document saved to XPS: {outputPath}");
        }
    }
}
