using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToMarkdown
{
    static void Main()
    {
        // Path to the source PDF template.
        string pdfPath = @"C:\Input\Template.pdf";

        // Path for the generated Markdown file.
        string mdPath = @"C:\Output\Result.md";

        // Load the PDF document. Aspose.Words automatically detects the format from the file extension.
        Document doc = new Document(pdfPath);

        // Create Markdown save options.
        // You can customize options here (e.g., list export mode, image handling, etc.).
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            // Example: export lists using Markdown syntax.
            ListExportMode = MarkdownListExportMode.MarkdownSyntax,

            // Example: export tables as raw HTML when they cannot be represented in pure Markdown.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

            // Example: export OfficeMath as LaTeX.
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex
        };

        // Save the document as Markdown.
        doc.Save(mdPath, mdOptions);
    }
}
