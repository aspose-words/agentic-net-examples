using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (DOCX, DOC, etc.).
        string inputPath = "input.docx";

        // Path where the EPUB file will be saved.
        string outputPath = "output.epub";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create save options for EPUB conversion.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveFormat = SaveFormat.Epub;          // Specify EPUB format.
        saveOptions.Encoding = Encoding.UTF8;             // Use UTF‑8 encoding.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph; // Optional: split by headings.
        saveOptions.ExportDocumentProperties = true;      // Optional: include document properties.

        // Save the document as EPUB using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
