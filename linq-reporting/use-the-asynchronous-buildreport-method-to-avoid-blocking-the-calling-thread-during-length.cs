using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingAsyncExample
{
    // Sample data model.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            // Ensure code page support (required for some data sources).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output.
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // 1. Create the template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 2. Load the template (simulating a separate load step).
            Document doc = new Document(templatePath);

            // 3. Prepare sample data.
            Order order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 3 },
                    new Item { Name = "Banana", Quantity = 5 },
                    new Item { Name = "Orange", Quantity = 2 }
                }
            };

            // (Optional) Serialize to JSON just to demonstrate the required package.
            string json = JsonConvert.SerializeObject(order, Formatting.Indented);
            Console.WriteLine("Sample data in JSON format:");
            Console.WriteLine(json);

            // 4. Build the report asynchronously.
            ReportingEngine engine = new ReportingEngine();

            await BuildReportAsync(engine, doc, order, "order");

            // 5. Save the generated report.
            doc.Save(outputPath);
        }

        // Wrap the synchronous BuildReport call in a Task to make it asynchronous.
        private static Task BuildReportAsync(ReportingEngine engine, Document doc, object root, string rootName)
        {
            return Task.Run(() => engine.BuildReport(doc, root, rootName));
        }
    }
}
