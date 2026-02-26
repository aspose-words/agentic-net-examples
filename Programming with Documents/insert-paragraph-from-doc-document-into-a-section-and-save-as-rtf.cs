using System;
using Aspose.Words;

class InsertParagraphFromDoc
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Create a new section in the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Every section must have a body; create and attach it.
        Body body = new Body(dstDoc);
        newSection.AppendChild(body);

        // Retrieve the paragraph you want to copy from the source document.
        // Here we take the first paragraph of the first section.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document.
        // The ImportNode method clones the node and resolves any style or list references.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the new section.
        body.AppendChild(importedParagraph);

        // Save the resulting document as RTF.
        dstDoc.Save("Result.rtf");
    }
}
