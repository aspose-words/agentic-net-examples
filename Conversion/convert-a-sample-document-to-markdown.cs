using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source document (DOCX, DOC, etc.).
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Path where the Markdown file will be saved.
            string outputPath = @"C:\Docs\SampleDocument.md";

            // Load the source document using Aspose.Words Document class.
            Document doc = new Document(inputPath);

            // Create MarkdownSaveOptions to specify Markdown-specific settings.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Ensure the format is set to Markdown (required when using this options class).
                SaveFormat = SaveFormat.Markdown,

                // Example option: export empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Example option: export lists using Markdown syntax.
                ListExportMode = MarkdownListExportMode.MarkdownSyntax
            };

            // Save the document as a Markdown file using the specified options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
