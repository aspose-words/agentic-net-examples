using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // Create a template document with a LINQ Reporting tag.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // The tag references the root object name "person".
            builder.Writeln("Hello, <<[person.Name]>>!");
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for reporting.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data source.
            Person person = new Person();

            // -------------------------------------------------
            // Temporarily disable reflection optimization for this report.
            // -------------------------------------------------
            bool originalOptimization = ReportingEngine.UseReflectionOptimization;
            ReportingEngine.UseReflectionOptimization = false;

            try
            {
                ReportingEngine engine = new ReportingEngine();
                // Build the report using the root name "person".
                engine.BuildReport(reportDoc, person, "person");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = originalOptimization;
            }

            // -------------------------------------------------
            // Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
