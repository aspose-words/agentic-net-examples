using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalExample
{
    // Data model for the report.
    public class Product
    {
        // Initialize non‑nullable reference types to avoid warnings.
        public string Name { get; set; } = "";
        public int Stock { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a simple title.
            builder.Writeln("Product Report");
            builder.Writeln();

            // Insert product name placeholder.
            builder.Writeln("Product: <<[model.Name]>>");

            // Conditional block: show "In stock" when Stock > 0.
            builder.Writeln("<<if [model.Stock > 0]>>In stock<</if>>");
            // Optional else block to demonstrate the false case.
            builder.Writeln("<<if [model.Stock <= 0]>>Out of stock<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template for report generation.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Prepare sample data.
            var product = new Product
            {
                Name = "Widget",
                Stock = 5 // Change this value to 0 to see the "Out of stock" message.
            };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, product, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}
