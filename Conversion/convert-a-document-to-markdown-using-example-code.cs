using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    public class MarkdownConverter
    {
        /// <summary>
        /// Converts a Word document to Markdown format.
        /// </summary>
        /// <param name="inputPath">Full path to the source document (e.g., .docx).</param>
        /// <param name="outputPath">Full path where the Markdown file will be saved.</param>
        public static void ConvertToMarkdown(string inputPath, string outputPath)
        {
            // Load the source document.
            Document doc = new Document(inputPath);

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (optional, but clarifies intent).
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document as Markdown.
            doc.Save(outputPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            string inputFile = @"C:\Docs\SampleDocument.docx";
            string outputFile = @"C:\Docs\SampleDocument.md";

            ConvertToMarkdown(inputFile, outputFile);

            Console.WriteLine("Document converted to Markdown successfully.");
        }
    }
}
