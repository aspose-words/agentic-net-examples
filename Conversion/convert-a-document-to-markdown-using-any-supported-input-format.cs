using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ConvertToMarkdownExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Simple argument handling: expect input and output paths.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: ConvertToMarkdownExample <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                ConvertToMarkdown(inputPath, outputPath);
                Console.WriteLine($"Document converted successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts a document from any supported format to Markdown.
        /// </summary>
        /// <param name="inputPath">Full path to the source document.</param>
        /// <param name="outputPath">Full path where the Markdown file will be saved.</param>
        private static void ConvertToMarkdown(string inputPath, string outputPath)
        {
            // Load the source document. Aspose.Words automatically detects the format.
            Document doc = new Document(inputPath);

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown
                // OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex // optional
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
