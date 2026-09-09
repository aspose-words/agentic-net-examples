using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Model class with a Unicode character in the property name.
    public class ReportModel
    {
        // Property name contains the character 'é'.
        public string Café { get; set; } = "Café au lait";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that references the Unicode property directly.
            builder.Writeln("Product description: <<[model.Café]>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            templateDoc.Save(templatePath);

            // Load the template back (required before building the report).
            Document loadedTemplate = new Document(templatePath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            ReportModel model = new ReportModel();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string reportPath = "report.docx";
            loadedTemplate.Save(reportPath);

            // Indicate completion (no interactive input).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
