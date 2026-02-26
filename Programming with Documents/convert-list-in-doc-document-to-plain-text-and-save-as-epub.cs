using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertListToPlainTextEpub
{
    static void Main()
    {
        // Paths to the source DOC file and the resulting EPUB file.
        string sourcePath = "InputDocument.doc";
        string outputPath = "PlainTextOutput.epub";

        // Load the original DOC document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure list labels are up‑to‑date so that the plain‑text conversion reflects correct numbering.
        sourceDoc.UpdateListLabels();

        // Extract the entire document as plain text (lists will be represented by their labels).
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document that will contain only the plain‑text content.
        Document plainTextDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(plainTextDoc);

        // Write the extracted plain text into the new document.
        builder.Writeln(plainText);

        // Configure EPUB save options.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,   // Target format.
            Encoding = Encoding.UTF8,       // Use UTF‑8 encoding.
            DocumentSplitCriteria = DocumentSplitCriteria.None, // Single part EPUB.
            ExportDocumentProperties = false
        };

        // Save the plain‑text document as an EPUB file.
        plainTextDoc.Save(outputPath, epubOptions);
    }
}
