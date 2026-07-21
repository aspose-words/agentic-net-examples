using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingIncludeExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report Title";
        public string Body { get; set; } = "This is the main body of the report generated with Aspose.Words LINQ Reporting.";
    }

    public class Program
    {
        public static void Main()
        {
            // Working directory.
            string workDir = Directory.GetCurrentDirectory();

            // -----------------------------------------------------------------
            // 1. Create a reusable header fragment (Header.docx).
            // -----------------------------------------------------------------
            string headerPath = Path.Combine(workDir, "Header.docx");
            Document headerDoc = new Document();
            DocumentBuilder headerBuilder = new DocumentBuilder(headerDoc);

            // Header content with a LINQ Reporting tag.
            headerBuilder.Writeln("<<[model.Title]>>");
            headerBuilder.Writeln("--------------------------------------------------");
            headerDoc.Save(headerPath);

            // -----------------------------------------------------------------
            // 2. Create the main template and embed the header fragment.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

            // Load the previously saved header fragment.
            Document loadedHeader = new Document(headerPath);

            // Insert the header fragment into the template.
            templateBuilder.InsertDocument(loadedHeader, ImportFormatMode.KeepSourceFormatting);
            templateBuilder.Writeln(); // Add an empty paragraph between header and body.

            // Body content with a LINQ Reporting tag.
            templateBuilder.Writeln("<<[model.Body]>>");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            ReportModel model = new ReportModel(); // Sample data.

            // Build the report using the root object name "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(workDir, "ReportOutput.docx");
            loadedTemplate.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
