using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be copied
        Document srcDoc = new Document("Source.doc");

        // Get the first paragraph from the source document (adjust index as needed)
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new destination document (blank)
        Document dstDoc = new Document();

        // Use the first (and only) section of the destination document
        Section dstSection = dstDoc.FirstSection;

        // Import the paragraph node into the destination document's node collection
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the end of the destination section's body
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a DOC file
        dstDoc.Save("Result.doc");
    }
}
