using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToEpub
{
    static void Main()
    {
        // Path to the source document (DOCX, DOC, etc.).
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the EPUB file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.epub";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create a SaveOptions instance suitable for EPUB using the factory method.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Epub);

        // The factory returns an HtmlSaveOptions object for EPUB.
        // Cast to HtmlSaveOptions to configure EPUB‑specific settings.
        if (saveOptions is HtmlSaveOptions epubOptions)
        {
            // Use UTF‑8 encoding for the EPUB content.
            epubOptions.Encoding = Encoding.UTF8;

            // Split the EPUB into separate HTML parts at heading paragraphs.
            epubOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Export built‑in and custom document properties into the EPUB.
            epubOptions.ExportDocumentProperties = true;
        }

        // Save the document as EPUB using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
