using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingAsyncExample
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
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static async Task Main()
        {
            // Register code page provider for possible encoding needs
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            await GenerateReportAsync();
        }

        private static async Task GenerateReportAsync()
        {
            // Paths for template and output
            string templatePath = "ReportTemplate.docx";
            string outputPath = "ReportResult.docx";

            // -------------------------------------------------
            // Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a title and a placeholder for the customer name
            builder.Writeln("Order Report");
            builder.Writeln("Customer: <<[model.CustomerName]>>");
            builder.Writeln();

            // Begin foreach loop over Items
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for reporting
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // Prepare sample data
            Order order = new()
            {
                CustomerName = "Acme Corp",
                Items = new()
                {
                    new Item { Index = 1, Name = "Widget A" },
                    new Item { Index = 2, Name = "Widget B" },
                    new Item { Index = 3, Name = "Widget C" }
                }
            };

            // -------------------------------------------------
            // Build the report asynchronously (wrapped in Task.Run)
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            bool success = await Task.Run(() => engine.BuildReport(doc, order, "model"));

            // Save the generated report if successful
            if (success)
            {
                doc.Save(outputPath);
            }
        }
    }
}
