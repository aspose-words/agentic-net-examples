using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MarkdownConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path to the resulting Markdown file.
            string outputPath = @"C:\Docs\ConvertedDocument.md";

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Ensure that list labels are up‑to‑date before conversion.
            doc.UpdateListLabels();

            // Configure Markdown save options.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                // Export list items using Markdown syntax so that ordered lists are
                // automatically numbered when the Markdown is rendered.
                ListExportMode = MarkdownListExportMode.MarkdownSyntax,

                // (Optional) Export tables that cannot be represented in pure Markdown
                // as raw HTML to preserve their structure.
                ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, mdOptions);
        }
    }
}
