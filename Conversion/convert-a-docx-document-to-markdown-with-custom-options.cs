using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocxToMarkdownExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = Path.Combine("Input", "SampleDocument.docx");

            // Path to the output Markdown file.
            string outputPath = Path.Combine("Output", "SampleDocument.md");

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export images as Base64 data URIs.
                ExportImagesAsBase64 = true,

                // Set a custom folder for images (used when ExportImagesAsBase64 is false).
                ImagesFolder = Path.Combine("Output", "Images"),

                // Alias used in the Markdown file for image URIs.
                ImagesFolderAlias = "http://example.com/images",

                // Export tables that cannot be represented in pure Markdown as raw HTML.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Export OfficeMath objects as LaTeX.
                OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,

                // Export links as reference blocks.
                LinkExportMode = MarkdownLinkExportMode.Reference,

                // Use pretty formatting for the generated Markdown.
                PrettyFormat = true,

                // Increase image resolution for any extracted images.
                ImageResolution = 300,

                // Ensure the save format is explicitly set to Markdown.
                SaveFormat = SaveFormat.Markdown
            };

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
