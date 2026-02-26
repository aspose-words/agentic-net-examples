using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsXps
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be inserted.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();
        // Ensure the document has at least one section, body and paragraph.
        dstDoc.EnsureMinimum();

        // Get the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the source paragraph into the destination document.
        // Keep the source formatting during the import.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the first section in the destination document.
        dstDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Prepare XPS save options (default options are sufficient for this task).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the resulting document as an XPS file.
        dstDoc.Save("Result.xps", xpsOptions);
    }
}
