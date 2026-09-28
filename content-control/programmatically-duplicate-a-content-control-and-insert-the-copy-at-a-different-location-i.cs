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

        // Ensure the document has at least one section and a body.
        Section section = doc.FirstSection ?? doc.AppendChild(new Section(doc));
        Body body = section.Body ?? section.AppendChild(new Body(doc));

        // Add the first paragraph where the original content control will be placed.
        Paragraph firstParagraph = new Paragraph(doc);
        firstParagraph.AppendChild(new Run(doc, "This is the original paragraph. "));
        body.AppendChild(firstParagraph);

        // Create a plain‑text content control (inline) with a title and tag.
        StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        originalSdt.Title = "SampleControl";
        originalSdt.Tag = "sample-tag";
        originalSdt.RemoveAllChildren();
        originalSdt.AppendChild(new Run(doc, "Original Content"));
        // Insert the content control into the first paragraph.
        firstParagraph.AppendChild(originalSdt);

        // Add a second paragraph where the duplicated content control will be inserted.
        Paragraph secondParagraph = new Paragraph(doc);
        secondParagraph.AppendChild(new Run(doc, "This is the target paragraph for the duplicate. "));
        body.AppendChild(secondParagraph);

        // Locate the original content control by its title.
        StructuredDocumentTag? foundSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(sdt => sdt.Title == "SampleControl");

        if (foundSdt != null)
        {
            // Clone the original content control (deep clone).
            StructuredDocumentTag clonedSdt = (StructuredDocumentTag)foundSdt.Clone(true);
            // Optionally modify the cloned control (e.g., change its title to avoid duplicates).
            clonedSdt.Title = "SampleControlCopy";
            clonedSdt.Tag = "sample-tag-copy";

            // Insert the cloned content control into the second paragraph.
            secondParagraph.AppendChild(clonedSdt);
        }
        else
        {
            throw new InvalidOperationException("Original content control not found.");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DuplicatedContentControl.docx");
        doc.Save(outputPath);
    }
}
