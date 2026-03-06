using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure Markdown save options (optional).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Example option: embed images directly as Base64 strings.
            ExportImagesAsBase64 = true
        };

        // Save the document as Markdown using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
