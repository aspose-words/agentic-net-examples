using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template that contains fields or expressions.
        Document doc = new Document("Template.pdf");

        // Ensure all fields (expressions) are evaluated before exporting.
        doc.UpdateFields();

        // Configure Markdown save options.
        // Example: export OfficeMath objects as LaTeX and tables as raw HTML.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the evaluated document as a Markdown file.
        doc.Save("Result.md", mdOptions);
    }
}
