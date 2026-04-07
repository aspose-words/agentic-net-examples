using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with public properties.
    public class Person
    {
        // Public properties that will be accessed from the template.
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags that reference the model's public properties.
            builder.Writeln("Person Report");
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");

            // Save the template to disk (required before loading for reporting).
            const string templatePath = "PersonReportTemplate.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Person model = new Person();

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // No need to set RestrictedMembers – the engine already limits access
            // to public members. If you need to restrict specific types, use
            // ReportingEngine.SetRestrictedTypes(params Type[] types) before building.
            // Example (commented out):
            // ReportingEngine.SetRestrictedTypes(typeof(System.Object));

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "PersonReportResult.docx";
            reportDoc.Save(outputPath);

            // Inform that the process completed.
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
