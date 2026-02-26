using System;
using Aspose.Words;

class InsertParagraphAndSaveAsMarkdown
{
    static void Main()
    {
        // Load the source DOCX document that contains the paragraph to be copied.
        Document srcDoc = new Document("SourceDocument.docx");

        // Retrieve the first paragraph from the source document (adjust index if needed).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new blank destination document. It already contains one empty section and body.
        Document dstDoc = new Document();

        // Import the source paragraph into the destination document, preserving its formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the first (and only) section.
        dstDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as Markdown.
        dstDoc.Save("ResultDocument.md", SaveFormat.Markdown);
    }
}
