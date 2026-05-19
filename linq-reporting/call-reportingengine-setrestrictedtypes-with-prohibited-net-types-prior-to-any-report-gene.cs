using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Person
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Define paths for the template and the generated report.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the data model.
            builder.Writeln("Hello <<[person.Name]>>!");

            // Save the template to disk (required before loading for report generation).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Set restricted .NET types before any report generation.
            // -----------------------------------------------------------------
            // Example: restrict access to the System.Environment type.
            ReportingEngine.SetRestrictedTypes(typeof(System.Environment));

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Prepare the data source.
            Person model = new Person { Name = "Aspose.Words" };

            // Build the report using the root object name "person".
            engine.BuildReport(loadedTemplate, model, "person");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
