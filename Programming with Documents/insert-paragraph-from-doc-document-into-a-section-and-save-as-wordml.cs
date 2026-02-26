using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphExample
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be inserted.
        Document srcDoc = new Document("Source.doc");

        // Create a new destination document.
        Document dstDoc = new Document();

        // Create a new section in the destination document.
        Section dstSection = new Section(dstDoc);
        dstDoc.AppendChild(dstSection);

        // Ensure the section has a body (required for a valid section).
        dstSection.EnsureMinimum();

        // Get the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the destination section.
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document in WORDML (XML) format.
        dstDoc.Save("Result.xml", SaveFormat.WordML);
    }
}
