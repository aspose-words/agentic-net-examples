using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (any format supported by Aspose.Words, e.g., DOCX).
            string sourcePath = @"C:\Docs\SampleDocument.docx";

            // Path where the Markdown file will be saved.
            string outputPath = @"C:\Docs\SampleDocument.md";

            // Load the source document using the Document constructor (load rule).
            Document doc = new Document(sourcePath);

            // Create MarkdownSaveOptions (create rule) and configure desired options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (save rule requirement).
                SaveFormat = SaveFormat.Markdown,

                // Example option: export tables that cannot be represented in pure Markdown as raw HTML.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Example option: preserve empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine
            };

            // Save the document as Markdown using the configured options (save rule).
            doc.Save(outputPath, saveOptions);

            Console.WriteLine("Document successfully converted to Markdown.");
        }
    }
}
