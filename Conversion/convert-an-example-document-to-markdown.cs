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
            string inputPath = @"C:\Docs\Example.docx";

            // Path where the Markdown file will be saved.
            string outputPath = @"C:\Docs\Example.md";

            // Load the source document.
            Document doc = new Document(inputPath);

            // Create Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Ensure the format is set to Markdown.
                SaveFormat = SaveFormat.Markdown,

                // Optional: customize how empty paragraphs are exported.
                // EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Optional: export images as Base64 to embed them directly.
                // ExportImagesAsBase64 = true,

                // Optional: choose how OfficeMath objects are exported.
                // OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex
            };

            // Save the document as Markdown.
            doc.Save(outputPath, saveOptions);
        }
    }
}
