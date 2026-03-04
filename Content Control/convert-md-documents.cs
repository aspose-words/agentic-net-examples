using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Markdown file.
        const string markdownPath = "SourceDocument.md";

        // Path for the converted PDF file.
        const string pdfPath = "ConvertedDocument.pdf";

        // Path for the re‑saved Markdown file with custom options.
        const string markdownOutPath = "ReSavedDocument.md";

        // ---------- Load the Markdown document ----------
        // Preserve empty lines while loading.
        var loadOptions = new MarkdownLoadOptions
        {
            PreserveEmptyLines = true
        };

        // The Document constructor reads the file using the specified load options.
        Document doc = new Document(markdownPath, loadOptions);

        // ---------- Save as PDF ----------
        // Use the built‑in PDF format; no additional options are required.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // ---------- Save back to Markdown with custom options ----------
        var saveOptions = new MarkdownSaveOptions
        {
            // Export list items using plain text (useful when the original list has complex numbering).
            ListExportMode = MarkdownListExportMode.PlainText,

            // Export OfficeMath objects as LaTeX.
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,

            // Export tables that cannot be represented in pure Markdown as raw HTML.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

            // Preserve empty paragraphs as empty lines.
            EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine
        };

        // Save the document using the configured Markdown options.
        doc.Save(markdownOutPath, saveOptions);
    }
}
