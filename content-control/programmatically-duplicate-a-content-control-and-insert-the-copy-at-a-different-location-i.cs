using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlDuplicationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an introductory paragraph.
            builder.Writeln("Paragraph before the original content control.");

            // Create a block‑level RichText content control.
            StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "OriginalControl",
                Tag = "original"
            };

            // Add a paragraph with some text inside the content control.
            Paragraph sdtParagraph = new Paragraph(doc);
            sdtParagraph.AppendChild(new Run(doc, "This is the original content control."));
            originalSdt.AppendChild(sdtParagraph);

            // Insert the original content control into the document body.
            doc.FirstSection.Body.AppendChild(originalSdt);

            // Add another paragraph after the original content control.
            builder.Writeln("Paragraph after the original content control.");

            // Clone the original content control (deep clone with its children).
            StructuredDocumentTag clonedSdt = (StructuredDocumentTag)originalSdt.Clone(true);
            clonedSdt.Title = "ClonedControl";
            clonedSdt.Tag = "cloned";

            // Insert the cloned content control after the first paragraph in the document.
            // The first paragraph is the one added by the first WriteLine call.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            firstParagraph.ParentNode.InsertAfter(clonedSdt, firstParagraph);

            // Save the resulting document.
            doc.Save("DuplicatedContentControl.docx");
        }
    }
}
