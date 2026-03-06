using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC file that contains the paragraph to be copied.
        Document srcDoc = new Document("SourceDocument.doc");

        // Create a new blank document that will receive the paragraph.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Guarantees at least one section.

        // Get the first paragraph from the source document (adjust index if needed).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document.
        // This handles style, list and other formatting translation.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph into a specific section of the destination document.
        // Here we use the first (and only) section, but any section from dstDoc.Sections can be used.
        Section targetSection = dstDoc.FirstSection;
        targetSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a DOCM file.
        dstDoc.Save("ResultDocument.docm");
    }
}
