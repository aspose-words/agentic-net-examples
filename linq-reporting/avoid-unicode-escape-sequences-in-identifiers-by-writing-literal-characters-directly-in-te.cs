using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingUnicodeDemo
{
    // Model class with a Unicode character in the property name.
    public class ReportModel
    {
        // Property name contains the character 'é'.
        public string Café { get; set; } = "Hello from Café!";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that references the Unicode property directly.
            builder.Writeln("Customer greeting: <<[model.Café]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the final report.
            loadedTemplate.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}
