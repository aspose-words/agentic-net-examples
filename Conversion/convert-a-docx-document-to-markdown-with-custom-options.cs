using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCX file path.
        string inputPath = @"C:\Docs\input.docx";

        // Output Markdown file path.
        string outputPath = @"C:\Docs\output.md";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure Markdown save options with custom settings.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Export tables that cannot be represented in pure Markdown as raw HTML.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

            // Export images as Base64 strings embedded directly in the Markdown file.
            ExportImagesAsBase64 = true,

            // Preserve empty paragraphs as empty lines in the output.
            // Note: The EmptyParagraphExportMode enum is not available in older versions of Aspose.Words.
            // If you are using a version that supports it, uncomment the line below:
            // EmptyParagraphExportMode = EmptyParagraphExportMode.EmptyLine,

            // Export OfficeMath objects as LaTeX markup.
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,

            // Use UTF‑8 encoding for the output file.
            Encoding = Encoding.UTF8,

            // Enable pretty formatting of the generated Markdown.
            PrettyFormat = true
        };

        // Save the document as Markdown using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
