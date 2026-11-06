using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingErrorDemo
{
    // Simple data model.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order { CustomerName = "John Doe" };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("Order Report");
            builder.Writeln(); // Blank line.

            // Conditional that references a non‑existent property – will cause an evaluation error.
            builder.Writeln("<<if [order.NonExisting]>>");
            builder.Writeln("<<error>>"); // Tag to display the error message.
            builder.Writeln("<</if>>");

            // Valid conditional section.
            builder.Writeln("<<if [order.CustomerName != null]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("<</if>>");

            // Save the template.
            builder.Document.Save(templatePath);

            // Load the template for reporting.
            var doc = new Document(templatePath);

            // Configure the reporting engine to embed inline error messages.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(doc, order, "order");

            // Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);

            // Output the result status.
            Console.WriteLine($"Report generated. Success: {success}");
            Console.WriteLine($"Template: {Path.GetFullPath(templatePath)}");
            Console.WriteLine($"Report:   {Path.GetFullPath(outputPath)}");
        }
    }
}
