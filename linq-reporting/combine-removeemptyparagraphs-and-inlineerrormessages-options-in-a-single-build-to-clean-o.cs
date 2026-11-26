using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Sample data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create template document
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Header
            builder.Writeln("Report Header");
            // Empty paragraph that should be removed
            builder.Writeln("");

            // Begin foreach over Items
            builder.Writeln("<<foreach [item in Items]>>");
            // Line with a missing property 'Price' to trigger inline error message
            builder.Writeln("Item: <<[item.Name]>>, Price: <<[item.Price]>>");
            // Another empty paragraph inside the loop
            builder.Writeln("   ");
            // End foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template (simulating separate load step)
            Document doc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Name = "Apple", Quantity = 10 },
                    new Item { Name = "Banana", Quantity = 20 }
                }
            };

            // Configure reporting engine with both options
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages | ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // Output result status
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Output saved to: {outputPath}");
        }
    }
}
