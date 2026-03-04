using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing Word document.
            Document doc = new Document("InputDocument.docx");

            // Create Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

            // Specify that the document should be saved in Markdown format.
            saveOptions.SaveFormat = SaveFormat.Markdown;

            // Example of specific options:
            // Export tables that cannot be represented in pure Markdown as raw HTML.
            saveOptions.ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables;

            // Export list items using plain text (instead of Markdown syntax).
            saveOptions.ListExportMode = MarkdownListExportMode.PlainText;

            // Export links as reference blocks.
            saveOptions.LinkExportMode = MarkdownLinkExportMode.Reference;

            // Export underline formatting using "++" markers.
            saveOptions.ExportUnderlineFormatting = true;

            // Save the document as a Markdown file using the configured options.
            doc.Save("OutputDocument.md", saveOptions);
        }
    }
}
