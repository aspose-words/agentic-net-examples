using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDateTimeExample
{
    public class ReportModel
    {
        // Initialize with a sample UTC date/time.
        public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            ReportModel model = new();

            // Create a new Word document that will serve as the template.
            Document template = new();
            DocumentBuilder builder = new(template);

            // Insert a line that formats the DateTime property as ISO 8601.
            // The expression tag {=model.CreatedDate:yyyy-MM-ddTHH:mm:ss} applies the format.
            builder.Writeln("Report generated at: {=model.CreatedDate:yyyy-MM-ddTHH:mm:ss}");

            // Save the template (optional, just for inspection).
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new();
            // The root object name must match the name used in the template tags ("model").
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "output.docx";
            template.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
