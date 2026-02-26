using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DotToMarkdownConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Provide the full path to the source .dot file and the desired output .md file.
            string sourceDotPath = @"C:\Docs\Template.dot";
            string targetMarkdownPath = @"C:\Docs\Converted.md";

            ConvertDotToMarkdown(sourceDotPath, targetMarkdownPath);
        }

        /// <summary>
        /// Converts a Microsoft Word template (.dot) to a Markdown (.md) document.
        /// </summary>
        /// <param name="inputPath">Full path to the source .dot file.</param>
        /// <param name="outputPath">Full path where the Markdown file will be saved.</param>
        static void ConvertDotToMarkdown(string inputPath, string outputPath)
        {
            // Load the .dot template into an Aspose.Words Document.
            Document doc = new Document(inputPath);

            // Configure save options for Markdown output.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (optional, but ensures correctness).
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
