using System;
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

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading – this will appear in the XPS outline.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Document with Content Controls");

            // Insert a block‑level RichText content control.
            StructuredDocumentTag blockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "BlockContent",
                Tag = "block-content"
            };
            // Add a paragraph inside the block content control.
            Paragraph blockParagraph = new Paragraph(doc);
            blockParagraph.AppendChild(new Run(doc, "This is text inside a block‑level content control."));
            blockSdt.AppendChild(blockParagraph);
            // Append the block content control to the document body.
            doc.FirstSection.Body.AppendChild(blockSdt);

            // Insert a blank line.
            builder.Writeln();

            // Insert an inline PlainText content control.
            StructuredDocumentTag inlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "customer-name"
            };
            inlineSdt.RemoveAllChildren();
            inlineSdt.AppendChild(new Run(doc, "John Doe"));
            // Append the inline content control to the current paragraph.
            builder.CurrentParagraph.AppendChild(inlineSdt);
            builder.Writeln(); // Move to a new paragraph after the inline control.

            // Save the document as DOCX (optional, for verification).
            doc.Save("ContentControls.docx");

            // Export the document to XPS while preserving content control boundaries.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();
            // Example: limit outline to heading levels 2 (optional).
            xpsOptions.OutlineOptions.HeadingsOutlineLevels = 2;

            doc.Save("ContentControls.xps", xpsOptions);
        }
    }
}
