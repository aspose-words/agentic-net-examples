using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToEpub
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words, e.g., .docx, .pdf, .html)
        string inputPath = @"C:\Docs\SampleDocument.docx";

        // Path where the resulting EPUB file will be saved
        string outputPath = @"C:\Docs\SampleDocument.epub";

        // Load the source document
        Document doc = new Document(inputPath);

        // Configure EPUB save options
        HtmlSaveOptions epubOptions = new HtmlSaveOptions
        {
            // Set the target format to EPUB
            SaveFormat = SaveFormat.Epub,

            // Use UTF-8 encoding (optional, default is UTF-8 without BOM)
            Encoding = Encoding.UTF8,

            // Split the EPUB into separate HTML parts at each heading paragraph
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Export built‑in and custom document properties into the EPUB package
            ExportDocumentProperties = true
        };

        // Save the document as EPUB using the configured options
        doc.Save(outputPath, epubOptions);
    }
}
