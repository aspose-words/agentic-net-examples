using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    public class ExampleConversion
    {
        private readonly string _myDir;
        private readonly string _artifactsDir;

        public ExampleConversion(string myDir, string artifactsDir)
        {
            _myDir = myDir;
            _artifactsDir = artifactsDir;
        }

        public void ConvertToMarkdown()
        {
            // Load the source document (e.g., a DOCX file) from the input directory.
            Document doc = new Document(Path.Combine(_myDir, "Example.docx"));

            // Create Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Export OfficeMath as images.
                OfficeMathExportMode = MarkdownOfficeMathExportMode.Image,

                // Export tables that cannot be represented in pure Markdown as raw HTML.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Export links as reference style.
                LinkExportMode = MarkdownLinkExportMode.Reference,

                // Export underline formatting using "++".
                ExportUnderlineFormatting = true,

                // Ensure the format is set to Markdown (required when using SaveFormat property).
                SaveFormat = SaveFormat.Markdown
            };

            // Define the output file path.
            string outputPath = Path.Combine(_artifactsDir, "ExampleConverted.md");

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            // Determine input and output directories.
            // Here we use the current working directory for input and a sub‑folder "Artifacts" for output.
            string inputDir = Directory.GetCurrentDirectory();
            string outputDir = Path.Combine(inputDir, "Artifacts");

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDir);

            // Create the conversion helper and run the conversion.
            var conversion = new ExampleConversion(inputDir, outputDir);
            conversion.ConvertToMarkdown();

            Console.WriteLine($"Markdown file created at: {Path.Combine(outputDir, "ExampleConverted.md")}");
        }
    }
}
