using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source ODT document.
        string odtPath = @"C:\Docs\SourceDocument.odt";

        // Path for the resulting Markdown file.
        string markdownPath = @"C:\Docs\ResultDocument.md";

        // Load the ODT document.
        Document doc = new Document(odtPath);

        // Configure Markdown save options to export tables as raw HTML,
        // which preserves complex table structures that cannot be represented
        // by pure Markdown syntax.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the document as Markdown using the configured options.
        doc.Save(markdownPath, saveOptions);
    }
}
