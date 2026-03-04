using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (DOCX, DOC, etc.).
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Path where the Markdown file will be saved.
            string outputPath = @"C:\Docs\OutputDocument.md";

            // Load the source document using the standard Document constructor.
            Document doc = new Document(inputPath);

            // Create Markdown save options and configure custom settings.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export tables as raw HTML instead of Markdown tables.
                ExportAsHtml = MarkdownExportAsHtml.Tables,

                // Export list items using standard Markdown syntax.
                ListExportMode = MarkdownListExportMode.MarkdownSyntax,

                // Set image resolution to 300 DPI.
                ImageResolution = 300,

                // Export empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Ensure the format is explicitly set to Markdown.
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document to Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
