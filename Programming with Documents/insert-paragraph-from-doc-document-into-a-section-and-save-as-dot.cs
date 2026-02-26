using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC file that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank document that will become the DOT template.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section, body, and paragraph.
        dstDoc.EnsureMinimum();

        // Get the first section of the destination document.
        Section dstSection = dstDoc.Sections[0];

        // Get the body of the destination section where the paragraph will be inserted.
        Body dstBody = dstSection.Body;

        // Create a NodeImporter to handle style and list translation between documents.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Copy each paragraph from the source document's first section into the destination body.
        foreach (Paragraph srcParagraph in srcDoc.FirstSection.Body.Paragraphs)
        {
            // Import the paragraph node into the destination document.
            Node importedParagraph = importer.ImportNode(srcParagraph, true);

            // Append the imported paragraph to the destination body.
            dstBody.AppendChild(importedParagraph);
        }

        // Save the resulting document as a DOT (Word template) file.
        dstDoc.Save("Result.dot");
    }
}
