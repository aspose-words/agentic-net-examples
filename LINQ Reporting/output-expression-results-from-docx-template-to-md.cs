using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains <<[field]>> expressions.
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Create a data source object. Replace this with your actual data model.
            var dataSource = new
            {
                Title = "Report Title",
                Date = DateTime.Now,
                Value = 12345.67
            };

            // Build the report by merging the data source with the template.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used in the template to reference the data source.
            engine.BuildReport(template, dataSource, "ds");

            // Configure Markdown save options.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                // Export list items using Markdown syntax (default).
                ListExportMode = MarkdownListExportMode.MarkdownSyntax,

                // Export empty paragraphs as empty lines (default).
                EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

                // Export OfficeMath as LaTeX (optional, change as needed).
                OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,

                // Ensure the document is saved in Markdown format.
                SaveFormat = SaveFormat.Markdown
            };

            // Path for the resulting Markdown file.
            string outputPath = @"C:\Docs\Result.md";

            // Save the populated document as Markdown.
            template.Save(outputPath, mdOptions);
        }
    }
}
