using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingUnicodeExample
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
            // Register code page provider for Aspose.Words (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            ReportModel model = new();

            // Create a template document programmatically.
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Insert a paragraph with a LINQ Reporting tag that references the Unicode property directly.
            builder.Writeln("<<[model.Café]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Load the template back (simulating a separate load step).
            Document loadedTemplate = new(templatePath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}
