using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Utility class with a static method that will be called from the template.
    public static class MyUtility
    {
        // Returns a greeting message for the supplied name.
        public static string GetGreeting(string name) => $"Hello, {name}!";
    }

    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Sample property referenced in the template.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting expression tag that calls the static method.
            // The type member access syntax allows us to reference MyUtility.GetGreeting.
            builder.Writeln("<<[MyUtility.GetGreeting(Name)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel { Name = "Aspose" };

            // -----------------------------------------------------------------
            // 4. Configure and execute the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Register the utility type so its static members can be accessed in the template.
            engine.KnownTypes.Add(typeof(MyUtility));

            // Build the report using the loaded template, the data model, and the root name "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
