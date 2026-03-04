using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveEpub
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Create a new destination document (blank) and obtain its first section.
        Document dstDoc = new Document();
        Section dstSection = dstDoc.FirstSection; // The blank document already has one section.

        // Ensure the section has a body (it does by default) – we will insert into its body.
        Body dstBody = dstSection.Body;

        // Retrieve the paragraph we want to copy from the source document.
        // Here we take the first paragraph; adjust the index as needed.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node from the source document into the destination document.
        // The NodeImporter handles style and list translation.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph into the destination section's body.
        // AppendChild adds it at the end of the existing content.
        dstBody.AppendChild(importedParagraph);

        // Prepare EPUB save options (optional: set split criteria, encoding, etc.).
        HtmlSaveOptions epubOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            ExportDocumentProperties = true,
            // By default EPUB splits by heading paragraphs; change if desired.
            DocumentSplitCriteria = DocumentSplitCriteria.None
        };

        // Save the resulting document as an EPUB file.
        dstDoc.Save("Result.epub", epubOptions);
    }
}
