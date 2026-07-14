using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
        public string Empty { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Paragraph with a normal value.
            builder.Writeln("Customer Name: <<[model.Name]>>");
            // Paragraph that contains only a tag whose value will be empty.
            builder.Writeln("This paragraph will be removed: <<[model.Empty]>>");
            // Another paragraph to show that the document remains valid after removal.
            builder.Writeln("Report generated successfully.");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Prepare the data source.
            var model = new ReportModel
            {
                Name = "John Doe",
                Empty = string.Empty // This will cause the paragraph to become empty.
            };

            // Load the template for reporting.
            var document = new Document(templatePath);

            // Configure the reporting engine to remove empty paragraphs.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root object name is "model" to match the tags.
            engine.BuildReport(document, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            document.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report saved to '{outputPath}'.");
        }
    }
}
