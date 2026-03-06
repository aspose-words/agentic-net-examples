using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source ODT document.
        string inputPath = @"C:\Docs\SourceDocument.odt";

        // Path where the Markdown file will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.md";

        // Load the ODT document.
        Document doc = new Document(inputPath);

        // Configure Markdown save options to export tables as raw HTML,
        // which preserves complex table structures that cannot be represented
        // by pure Markdown syntax.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the document as Markdown using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
