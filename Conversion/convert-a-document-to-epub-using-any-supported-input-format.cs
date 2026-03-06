using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToEpub
{
    static void Main()
    {
        // Path to the source document (can be DOCX, PDF, HTML, etc.).
        string inputPath = @"C:\Input\sample.docx";

        // Path where the EPUB file will be saved.
        string outputPath = @"C:\Output\sample.epub";

        // Load the source document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Create save options for EPUB conversion.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Specify that the output format is EPUB.
            SaveFormat = SaveFormat.Epub,

            // Use UTF-8 encoding for the EPUB content.
            Encoding = Encoding.UTF8,

            // Export built‑in and custom document properties to the EPUB.
            ExportDocumentProperties = true,

            // Optional: split the EPUB into multiple HTML parts at heading paragraphs.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph
        };

        // Save the document as EPUB using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
