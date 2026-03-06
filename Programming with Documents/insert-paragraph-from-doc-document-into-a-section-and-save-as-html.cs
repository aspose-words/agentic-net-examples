using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveHtml
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be copied.
        Document srcDoc = new Document("SourceDocument.docx");

        // Retrieve the first paragraph from the source document.
        // Adjust the index if a different paragraph is required.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.Paragraphs[0];

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Import the source paragraph into the destination document.
        // The NodeImporter handles style and list translation between documents.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the end of the first section's body.
        // The destination document already contains a paragraph, so we insert after it.
        dstDoc.FirstSection.Body.InsertAfter(importedParagraph, dstDoc.FirstSection.Body.LastParagraph);

        // Save the resulting document as HTML.
        dstDoc.Save("Result.html", SaveFormat.Html);
    }
}
