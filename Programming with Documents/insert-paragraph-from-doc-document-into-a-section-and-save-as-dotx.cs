using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section, body and paragraph.
        dstDoc.EnsureMinimum();

        // Select the paragraph to insert. Here we take the first paragraph of the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document's node collection.
        // The second argument (true) performs a deep clone of the node.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Get the first (and only) section of the destination document.
        Section destSection = dstDoc.FirstSection;

        // Insert the imported paragraph after the last paragraph of the destination section's body.
        Paragraph lastParagraph = destSection.Body.LastParagraph;
        destSection.Body.InsertAfter(importedParagraph, lastParagraph);

        // Save the resulting document as a DOTX template.
        dstDoc.Save("Result.dotx", SaveFormat.Dotx);
    }
}
