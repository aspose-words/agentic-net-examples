using System;
using Aspose.Words;

class InsertParagraphIntoSection
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank document that will receive the paragraph.
        Document dstDoc = new Document();

        // Optionally add a new empty section to the destination document.
        // (The blank document already contains one section, but this shows explicit usage.)
        Section dstSection = new Section(dstDoc);
        dstDoc.AppendChild(dstSection);

        // Ensure the section has a body to hold paragraphs.
        if (dstSection.Body == null)
            dstSection.AppendChild(new Body(dstDoc));

        // Get the paragraph you want to insert from the source document.
        // Here we take the first paragraph of the first section.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document.
        // This clones the node and translates styles/lists as needed.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the target section.
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a DOTM (Word macro-enabled template).
        dstDoc.Save("Result.dotm");
    }
}
