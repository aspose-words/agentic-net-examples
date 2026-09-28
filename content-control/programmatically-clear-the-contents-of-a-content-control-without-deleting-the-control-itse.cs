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

        // Insert a paragraph with some introductory text.
        Paragraph intro = new Paragraph(doc);
        intro.AppendChild(new Run(doc, "Document with a content control:"));
        doc.FirstSection.Body.AppendChild(intro);

        // Create a plain‑text content control (inline) with a title and tag.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "SampleControl";
        sdt.Tag = "sample-control";

        // Add initial content to the control.
        sdt.RemoveAllChildren(); // Ensure it is empty before adding content.
        sdt.AppendChild(new Run(doc, "Initial content"));

        // Append the content control to a new paragraph.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(sdt);
        doc.FirstSection.Body.AppendChild(para);

        // Save the document that contains the populated content control.
        doc.Save("original.docx");

        // Load the document again (simulating a separate operation).
        Document loadedDoc = new Document("original.docx");

        // Find the content control by its title.
        StructuredDocumentTag targetSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(tag => tag.Title == "SampleControl");

        if (targetSdt == null)
            throw new InvalidOperationException("Content control not found.");

        // Clear the contents of the content control without removing the control itself.
        targetSdt.RemoveAllChildren();

        // Optionally, you could add placeholder text or leave it empty.
        // Here we leave it empty to demonstrate a cleared control.

        // Save the resulting document.
        loadedDoc.Save("cleared.docx");
    }
}
