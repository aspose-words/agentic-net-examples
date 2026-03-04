using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class MarkdownRenderer
{
    static void Main()
    {
        // Sample markdown text to be processed.
        string markdown = "# Sample Document\r\n\r\nThis is a **bold** paragraph with a table:\r\n\r\n| Header1 | Header2 |\r\n|---------|---------|\r\n| Cell1   | Cell2   |\r\n";

        // Load the markdown into an Aspose.Words Document using MarkdownLoadOptions.
        // PreserveEmptyLines ensures that blank lines in the source are kept.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(markdown)))
        {
            MarkdownLoadOptions loadOptions = new MarkdownLoadOptions
            {
                PreserveEmptyLines = true
            };

            Document doc = new Document(stream, loadOptions);

            // Configure how the document will be saved back to Markdown.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export tables as raw HTML to retain complex structures.
                ExportAsHtml = MarkdownExportAsHtml.Tables,

                // Export list items using standard Markdown syntax.
                ListExportMode = MarkdownListExportMode.MarkdownSyntax,

                // Export empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Do not embed the Aspose.Words generator name in the output.
                ExportGeneratorName = false
            };

            // Define the output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "RenderedDocument.md");

            // Save the document using the configured MarkdownSaveOptions.
            doc.Save(outputPath, saveOptions);
        }
    }
}
