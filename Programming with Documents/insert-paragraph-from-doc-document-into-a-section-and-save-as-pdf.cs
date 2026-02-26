using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSavePdf
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Use the first (and only) section that already exists in the blank document.
        Section dstSection = dstDoc.FirstSection;

        // Get the paragraph we want to copy from the source document (first paragraph of the first section).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the destination section.
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as PDF.
        dstDoc.Save("Result.pdf", SaveFormat.Pdf);
    }
}
