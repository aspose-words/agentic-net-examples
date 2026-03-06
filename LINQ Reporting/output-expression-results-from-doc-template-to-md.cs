using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains Aspose.Words expression tags.
        string templatePath = @"MyDir\Template.docx";

        // Path where the resulting Markdown file will be saved.
        string outputPath = @"ArtifactsDir\Result.md";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Example data source – can be any .NET object, DataSet, etc.
        var dataSource = new
        {
            Name = "John Doe",
            Age = 30,
            Salary = 12345.67
        };

        // Populate the template with data.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used inside the template to reference the data source.
        engine.BuildReport(doc, dataSource, "ds");

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Ensure the format is explicitly set to Markdown.
            SaveFormat = SaveFormat.Markdown,

            // Example: export any OfficeMath as plain text (default is Text).
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Text,

            // Example: export tables as raw HTML if needed.
            // ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the populated document as Markdown.
        doc.Save(outputPath, saveOptions);
    }
}
