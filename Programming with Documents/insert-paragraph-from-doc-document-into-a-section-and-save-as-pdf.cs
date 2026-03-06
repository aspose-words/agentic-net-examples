using System;
using Aspose.Words;

class InsertParagraphIntoSection
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.docx");

        // Create a new destination document.
        Document dstDoc = new Document();

        // Add a new empty section to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.Sections.Add(newSection);

        // Ensure the new section has a body to hold paragraphs.
        Body body = newSection.Body ?? new Body(dstDoc);
        if (newSection.Body == null)
            newSection.AppendChild(body);

        // Get the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node into the destination document.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the new section's body.
        body.AppendChild(importedParagraph);

        // Save the resulting document as PDF.
        dstDoc.Save("Result.pdf", SaveFormat.Pdf);
    }
}
