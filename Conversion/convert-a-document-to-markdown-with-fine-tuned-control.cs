using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the Markdown file will be saved.
            string outputPath = @"C:\Docs\ConvertedDocument.md";

            // Load the Word document.
            Document doc = new Document(inputPath);

            // Create Markdown save options with fine‑tuned settings.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export tables that cannot be represented in pure Markdown as raw HTML.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Export OfficeMath objects as LaTeX compatible with MarkItDown.
                OfficeMathExportMode = MarkdownOfficeMathExportMode.MarkItDown,

                // Export links as reference style.
                LinkExportMode = MarkdownLinkExportMode.Reference,

                // Export list items using plain text (labels are updated).
                ListExportMode = MarkdownListExportMode.PlainText,

                // Export empty paragraphs as empty lines.
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Export underline formatting using "++" markers.
                ExportUnderlineFormatting = true,

                // Save images as Base64 strings embedded in the Markdown file.
                ExportImagesAsBase64 = true,

                // Specify a folder for external images (if ExportImagesAsBase64 is false).
                ImagesFolder = @"C:\Docs\Images",
                ImagesFolderAlias = "images",

                // Enable pretty formatting for better readability.
                PrettyFormat = true,

                // Set the image resolution to 150 DPI.
                ImageResolution = 150,

                // Ensure page breaks are preserved.
                ForcePageBreaks = true
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
