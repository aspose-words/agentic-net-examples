using System;
using Aspose.Words;

class InsertParagraphFromDoc
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("SourceDocument.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Create a new section in the destination document.
        Section newSection = new Section(dstDoc);
        // Append the new section to the document's node collection.
        dstDoc.AppendChild(newSection);

        // Ensure the new section has a body (required to hold paragraphs).
        Body body = new Body(dstDoc);
        newSection.AppendChild(body);

        // Get the first paragraph from the source document (adjust as needed).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document's node hierarchy.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the new section's body.
        body.AppendChild(importedParagraph);

        // Save the resulting document as DOCX.
        dstDoc.Save("ResultDocument.docx");
    }
}
