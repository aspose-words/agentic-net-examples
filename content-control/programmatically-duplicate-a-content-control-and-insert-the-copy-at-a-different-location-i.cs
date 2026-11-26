using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace DuplicateContentControlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add an initial paragraph.
            Paragraph introPara = new Paragraph(doc);
            introPara.AppendChild(new Run(doc, "Document with original content control:"));
            doc.FirstSection.Body.AppendChild(introPara);

            // Create a block-level rich text content control.
            StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
            originalSdt.Title = "OriginalControl";
            originalSdt.Tag = "original";

            // Add a paragraph inside the content control.
            Paragraph sdtParagraph = new Paragraph(doc);
            sdtParagraph.AppendChild(new Run(doc, "This is the original content control text."));
            originalSdt.AppendChild(sdtParagraph);

            // Insert the original content control into the document body.
            doc.FirstSection.Body.AppendChild(originalSdt);

            // Add another paragraph after which the duplicate will be placed.
            Paragraph afterPara = new Paragraph(doc);
            afterPara.AppendChild(new Run(doc, "Document after the original content control:"));
            doc.FirstSection.Body.AppendChild(afterPara);

            // Locate the original content control by its title.
            StructuredDocumentTag foundSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .FirstOrDefault(s => s.Title == "OriginalControl");

            if (foundSdt != null)
            {
                // Clone the found content control (deep clone).
                StructuredDocumentTag clonedSdt = (StructuredDocumentTag)foundSdt.Clone(true);
                clonedSdt.Title = "ClonedControl";
                clonedSdt.Tag = "cloned";

                // Insert the cloned content control after the second paragraph.
                // The second paragraph is at index 1 (0‑based) in the body.
                Paragraph insertionPoint = doc.FirstSection.Body.Paragraphs[1];
                insertionPoint.ParentNode.InsertAfter(clonedSdt, insertionPoint);
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DuplicatedContentControl.docx");
            doc.Save(outputPath);
        }
    }
}
