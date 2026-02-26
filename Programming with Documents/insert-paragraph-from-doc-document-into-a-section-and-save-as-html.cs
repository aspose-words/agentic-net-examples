using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveHtml
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.docx");

        // Create a new (blank) destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section/body/paragraph.
        dstDoc.EnsureMinimum();

        // Retrieve the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.Paragraphs[0];

        // Import the paragraph node into the destination document.
        // KeepSourceFormatting preserves the original formatting of the paragraph.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the end of the first section's body in the destination document.
        Section dstSection = dstDoc.FirstSection;
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as HTML.
        dstDoc.Save("Result.html", SaveFormat.Html);
    }
}
