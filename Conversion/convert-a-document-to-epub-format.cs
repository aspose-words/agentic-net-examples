using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToEpub
{
    static void Main()
    {
        // Input DOCX file path.
        string inputPath = @"C:\Docs\Input.docx";

        // Output EPUB file path.
        string outputPath = @"C:\Docs\Output.epub";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create save options for EPUB conversion.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveFormat = SaveFormat.Epub;          // Set target format to EPUB.
        saveOptions.Encoding = Encoding.UTF8;             // Use UTF‑8 encoding (optional).
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph; // Split by headings (optional).
        saveOptions.ExportDocumentProperties = true;      // Export document properties (optional).

        // Save the document as EPUB using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
