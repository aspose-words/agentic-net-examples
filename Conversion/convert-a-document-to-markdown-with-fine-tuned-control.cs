using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document.
            Document doc = new Document(@"C:\Input\SampleDocument.docx");

            // Create a MarkdownSaveOptions instance to fine‑tune the conversion.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                // Export tables that cannot be represented in pure Markdown as raw HTML.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Export list items using Markdown syntax (automatic numbering).
                ListExportMode = MarkdownListExportMode.MarkdownSyntax,

                // Export OfficeMath objects as LaTeX compatible with MarkItDown.
                OfficeMathExportMode = MarkdownOfficeMathExportMode.MarkItDown,

                // Preserve empty paragraphs as empty lines in the output.
                // Note: EmptyParagraphExportMode is not available in older Aspose.Words versions.
                // If your version supports it, uncomment the line below and add the appropriate using.
                // EmptyParagraphExportMode = EmptyParagraphExportMode.EmptyLine,

                // Encode the output as UTF‑8 (default, shown explicitly).
                Encoding = Encoding.UTF8,

                // Do not embed images as Base64; they will be saved as separate files.
                ExportImagesAsBase64 = false,

                // Specify a folder where extracted images will be written.
                ImagesFolder = @"C:\Output\Images",
                ImagesFolderAlias = "Images", // Used to build image URIs in the Markdown file.

                // Enable pretty formatting for better readability.
                PrettyFormat = true
            };

            // Ensure the output directories exist.
            Directory.CreateDirectory(@"C:\Output");
            Directory.CreateDirectory(mdOptions.ImagesFolder);

            // Save the document as Markdown using the configured options.
            doc.Save(@"C:\Output\ConvertedDocument.md", mdOptions);
        }
    }
}
