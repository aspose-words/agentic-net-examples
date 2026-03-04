using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    public class MarkdownConverter
    {
        // Converts a Word document to Markdown format.
        public void ConvertToMarkdown(string inputFilePath, string outputFilePath)
        {
            // Load the source document.
            Document doc = new Document(inputFilePath);

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (optional, as the class defaults to this).
                SaveFormat = SaveFormat.Markdown,

                // Example: export tables that cannot be represented in pure Markdown as raw HTML.
                // ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Example: preserve empty paragraphs as empty lines.
                // EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Example: export list items using Markdown syntax.
                // ListExportMode = MarkdownListExportMode.MarkdownSyntax
            };

            // Save the document as a Markdown file.
            doc.Save(outputFilePath, saveOptions);
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            // Define input and output paths.
            string inputPath = @"C:\Docs\Sample.docx";      // Replace with actual source file path.
            string outputPath = @"C:\Docs\Sample.md";      // Desired Markdown output path.

            // Perform conversion.
            MarkdownConverter converter = new MarkdownConverter();
            converter.ConvertToMarkdown(inputPath, outputPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
