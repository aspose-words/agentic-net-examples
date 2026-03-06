using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToEpubConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\source.docx";

        // Path where the resulting EPUB file will be saved.
        string outputFile = @"C:\Docs\result.epub";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Configure save options for EPUB output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Use UTF-8 encoding (without BOM) for the EPUB content.
            Encoding = Encoding.UTF8,

            // Export built‑in and custom document properties into the EPUB.
            ExportDocumentProperties = true,

            // Split the EPUB into separate HTML parts at heading paragraphs.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph
        };

        // Save the document as a single EPUB file using the configured options.
        doc.Save(outputFile, saveOptions);
    }
}
