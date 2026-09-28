using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Sample data model
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var order = new Order
            {
                CustomerName = "Acme Corp",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Widget A" },
                    new Item { Index = 2, Name = "Widget B" },
                    new Item { Index = 3, Name = "Widget C" }
                }
            };

            // Create a template document programmatically
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Enable reflection optimization for better performance on large collections
            ReportingEngine.UseReflectionOptimization = true;

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report
            const string outputPath = "report.docx";
            doc.Save(outputPath);
        }
    }
}
