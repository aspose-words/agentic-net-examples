using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC file that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Retrieve the paragraph you want to insert.
        // Here we take the first paragraph of the first section.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new blank document that will receive the paragraph.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Ensure the document has at least one section, body and paragraph.

        // Get the target section where the paragraph will be inserted.
        Section dstSection = dstDoc.FirstSection;

        // Import the paragraph node into the destination document.
        // The import copies the node and resolves any style or list conflicts.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true, ImportFormatMode.KeepSourceFormatting);

        // Append the imported paragraph to the end of the target section's body.
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a macro‑enabled DOCM file.
        dstDoc.Save("Result.docm", SaveFormat.Docm);
    }
}
