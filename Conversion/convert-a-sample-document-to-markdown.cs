using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path where the Markdown file will be saved.
        string outputPath = @"C:\Docs\Sample.md";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Configure options for saving as Markdown.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Export tables that cannot be represented in pure Markdown as raw HTML.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

            // Export links using reference style.
            LinkExportMode = MarkdownLinkExportMode.Reference,

            // Preserve empty paragraphs as empty lines.
            EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

            // Explicitly set the target format (optional, defaults to Markdown).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document in Markdown format using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
