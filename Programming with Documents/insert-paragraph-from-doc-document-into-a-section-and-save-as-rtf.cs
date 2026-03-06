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

        // Ensure the destination has at least one section (required for a valid document).
        dstDoc.EnsureMinimum();

        // Retrieve the paragraph you want to insert.
        // Here we take the first paragraph of the first section of the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document.
        // The second argument (true) performs a deep clone of the node.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Create a new section in the destination document where the paragraph will be placed.
        Section targetSection = new Section(dstDoc);
        dstDoc.Sections.Add(targetSection);

        // Append the imported paragraph to the body of the target section.
        targetSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as RTF.
        dstDoc.Save("Result.rtf");
    }
}
