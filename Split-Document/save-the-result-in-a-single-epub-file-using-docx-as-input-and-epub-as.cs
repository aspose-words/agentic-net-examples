using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path where the resulting EPUB file will be saved.
        string outputPath = @"C:\Docs\output.epub";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Create EPUB save options.
        // The constructor sets the format to EPUB.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub);

        // Use UTF‑8 encoding for the EPUB content.
        saveOptions.Encoding = Encoding.UTF8;

        // Optional: export built‑in and custom document properties.
        saveOptions.ExportDocumentProperties = true;

        // Optional: split the EPUB into separate HTML parts at heading paragraphs.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

        // Save the document as an EPUB file using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
