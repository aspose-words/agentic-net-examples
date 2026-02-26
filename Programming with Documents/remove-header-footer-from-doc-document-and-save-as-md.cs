using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsMarkdown
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting Markdown file will be saved.
        string outputPath = "output.md";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Configure Markdown save options to omit headers and footers.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Do not export any headers or footers.
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None,
            // Explicitly set the format to Markdown (optional, but clarifies intent).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
