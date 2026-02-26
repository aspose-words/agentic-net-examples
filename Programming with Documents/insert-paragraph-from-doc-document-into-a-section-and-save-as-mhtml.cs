using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveMhtml
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Create a new (blank) destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section.
        // A blank document already contains a Section, Body and Paragraph.
        Section dstSection = dstDoc.FirstSection;

        // Get the first paragraph from the source document.
        // Adjust the index if you need a different paragraph.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.Paragraphs[0];

        // Import the paragraph node into the destination document.
        // The NodeImporter handles style and list translation.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the destination section's body.
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as MHTML (Web archive) format.
        dstDoc.Save("Result.mht", SaveFormat.Mhtml);
    }
}
