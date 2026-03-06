using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Sample Markdown content to be loaded.
        string markdown = "# Sample Document\r\n\r\nThis is **bold** text.\r\n\r\n- Item 1\r\n- Item 2\r\n";

        // Load the Markdown text into an Aspose.Words Document using MarkdownLoadOptions.
        using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(markdown)))
        {
            // Preserve empty lines while loading.
            MarkdownLoadOptions loadOptions = new MarkdownLoadOptions
            {
                PreserveEmptyLines = true
            };

            Document doc = new Document(stream, loadOptions);

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // No raw HTML export.
                ExportAsHtml = MarkdownExportAsHtml.None,
                // Export lists using standard Markdown syntax.
                ListExportMode = MarkdownListExportMode.MarkdownSyntax,
                // Export empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine
            };

            // Define output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "output.md");

            // Save the document as a Markdown file using the specified options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
