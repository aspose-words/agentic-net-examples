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

        // Ensure the document has at least two paragraphs.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        Paragraph secondParagraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(secondParagraph);

        // Create an inline plain‑text content control (SDT) in the first paragraph.
        StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        originalSdt.Title = "OriginalControl";
        originalSdt.Tag = "original";
        originalSdt.RemoveAllChildren();
        originalSdt.AppendChild(new Run(doc, "Original content"));
        firstParagraph.AppendChild(originalSdt);

        // Locate the content control we just created (by title).
        StructuredDocumentTag? foundSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(sdt => sdt.Title == "OriginalControl");

        if (foundSdt != null)
        {
            // Clone the found content control (deep clone).
            StructuredDocumentTag clonedSdt = (StructuredDocumentTag)foundSdt.Clone(true);

            // Optionally modify the clone (e.g., change title/tag to indicate it's a copy).
            clonedSdt.Title = "ClonedControl";
            clonedSdt.Tag = "cloned";

            // Insert the cloned content control into the second paragraph.
            secondParagraph.AppendChild(clonedSdt);
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DuplicatedContentControl.docx");
        doc.Save(outputPath);
    }
}
