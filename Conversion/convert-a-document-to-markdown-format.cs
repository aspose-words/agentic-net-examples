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
            string inputPath = @"MyDir\Document.docx";

            // Path where the Markdown file will be saved.
            string outputPath = @"ArtifactsDir\Document.md";

            // Load the source document.
            Document doc = new Document(inputPath);

            // Create Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (optional, but clarifies intent).
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document as Markdown using the specified options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
