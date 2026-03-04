using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownConversion
{
    /// <summary>
    /// Demonstrates how to convert a DOCX file to Markdown using custom <see cref="MarkdownSaveOptions"/>.
    /// </summary>
    public static class DocxToMarkdownConverter
    {
        /// <summary>
        /// Converts the specified DOCX document to a Markdown file applying custom save options.
        /// </summary>
        /// <param name="inputDocxPath">Full path to the source DOCX file.</param>
        /// <param name="outputMarkdownPath">Full path where the resulting Markdown file will be saved.</param>
        public static void Convert(string inputDocxPath, string outputMarkdownPath)
        {
            // Load the existing DOCX document.
            Document doc = new Document(inputDocxPath);

            // Create and configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export images as Base64 strings embedded directly in the Markdown.
                ExportImagesAsBase64 = true,

                // Export tables that cannot be represented in pure Markdown as raw HTML.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

                // Export OfficeMath objects as LaTeX (useful for Markdown processors that support LaTeX).
                OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,

                // Use UTF-8 encoding (default, but set explicitly for clarity).
                Encoding = Encoding.UTF8,

                // Enable pretty formatting for readability.
                PrettyFormat = true
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputMarkdownPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            string sourceDocx = @"C:\Docs\SampleDocument.docx";
            string targetMd   = @"C:\Docs\SampleDocument.md";

            Convert(sourceDocx, targetMd);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
