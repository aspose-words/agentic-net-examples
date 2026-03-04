using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOTM (macro‑enabled template) file.
        string inputPath = @"C:\Docs\Template.dotm";

        // Path where the resulting EPUB file will be written.
        string outputPath = @"C:\Docs\Template.epub";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Configure EPUB save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveFormat = SaveFormat.Epub;          // Specify EPUB format.
        saveOptions.Encoding = Encoding.UTF8;              // Use UTF‑8 encoding.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph; // Optional split.
        saveOptions.ExportDocumentProperties = true;      // Export built‑in and custom properties.

        // Save the document as EPUB using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
