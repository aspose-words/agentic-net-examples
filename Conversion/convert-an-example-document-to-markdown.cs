using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    public class MarkdownConverter
    {
        /// <summary>
        /// Converts a Word document to Markdown using Aspose.Words.
        /// </summary>
        /// <param name="inputPath">Full path to the source .doc/.docx file.</param>
        /// <param name="outputPath">Full path where the .md file will be saved.</param>
        public void ConvertToMarkdown(string inputPath, string outputPath)
        {
            // Load the source document from the file system.
            Document doc = new Document(inputPath);

            // Initialize Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

            // Example option: export OfficeMath objects as images.
            saveOptions.OfficeMathExportMode = MarkdownOfficeMathExportMode.Image;

            // Additional optional settings (uncomment as needed):
            // Export tables as raw HTML.
            // saveOptions.ExportAsHtml = MarkdownExportAsHtml.Tables;

            // Export empty paragraphs as empty lines.
            // saveOptions.EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine;

            // Save the document in Markdown format.
            doc.Save(outputPath, saveOptions);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point of the console application.
        /// </summary>
        static void Main(string[] args)
        {
            // Simple argument validation.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: MarkdownConverter <input.docx> <output.md>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                var converter = new MarkdownConverter();
                converter.ConvertToMarkdown(inputPath, outputPath);
                Console.WriteLine($"Document converted successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
