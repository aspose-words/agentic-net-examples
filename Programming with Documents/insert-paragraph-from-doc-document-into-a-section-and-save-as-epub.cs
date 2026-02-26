using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveEpub
{
    static void Main()
    {
        // Paths to the source DOC file and the output EPUB file.
        string sourceDocPath = @"C:\Input\SourceDocument.docx";
        string outputEpubPath = @"C:\Output\Result.epub";

        // Load the source document that contains the paragraph to be inserted.
        Document srcDoc = new Document(sourceDocPath);

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section.
        // (A new Document already contains a default section with a body and a paragraph.)
        Section dstSection = dstDoc.FirstSection;

        // Get the paragraph we want to copy from the source document.
        // For example, take the first paragraph of the first section.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Prepare a NodeImporter to import nodes from srcDoc into dstDoc.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the paragraph node (deep clone) so it can be inserted into the destination.
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the target section.
        dstSection.Body.AppendChild(importedParagraph);

        // Configure EPUB save options (optional: set encoding, split criteria, export properties).
        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            Encoding = System.Text.Encoding.UTF8,
            DocumentSplitCriteria = DocumentSplitCriteria.None,
            ExportDocumentProperties = true
        };

        // Save the destination document as an EPUB file.
        dstDoc.Save(outputEpubPath, epubSaveOptions);
    }
}
