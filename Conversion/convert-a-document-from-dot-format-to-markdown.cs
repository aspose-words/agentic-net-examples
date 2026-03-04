using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentConversion
{
    /// <summary>
    /// Provides functionality to convert a Microsoft Word DOT template to Markdown format.
    /// </summary>
    public static class DotToMarkdownConverter
    {
        /// <summary>
        /// Converts a DOT file to a Markdown file.
        /// </summary>
        /// <param name="inputDotPath">Full path to the source .dot file.</param>
        /// <param name="outputMarkdownPath">Full path where the resulting .md file will be saved.</param>
        public static void Convert(string inputDotPath, string outputMarkdownPath)
        {
            // Load the DOT template using the Document constructor that accepts a file name.
            Document doc = new Document(inputDotPath);

            // Configure save options for Markdown output.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (optional, but ensures correct behavior).
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputMarkdownPath, saveOptions);
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            // Paths can be adjusted as needed.
            string inputDot = @"C:\Docs\Template.dot";
            string outputMd = @"C:\Docs\Result.md";

            // Perform the conversion.
            DotToMarkdownConverter.Convert(inputDot, outputMd);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
