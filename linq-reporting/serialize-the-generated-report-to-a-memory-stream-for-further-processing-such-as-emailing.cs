using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingMemoryStreamExample
{
    // Data model classes
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

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags.
            builder.Writeln("Order for <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            Order sampleOrder = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 3 },
                    new Item { Name = "Banana", Quantity = 5 },
                    new Item { Name = "Cherry", Quantity = 7 }
                }
            };

            // 3. Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // No special options required for this simple example.
            engine.BuildReport(template, sampleOrder, "order");

            // 4. Serialize the generated report to a memory stream.
            using (MemoryStream reportStream = new MemoryStream())
            {
                // Save the document into the stream in DOCX format.
                template.Save(reportStream, SaveFormat.Docx);

                // Reset the stream position to the beginning for further processing.
                reportStream.Position = 0;

                // Example: output the size of the generated report.
                Console.WriteLine($"Report generated. Stream length: {reportStream.Length} bytes.");

                // At this point the stream can be attached to an email, uploaded, etc.
                // For demonstration, we could write the stream to a file (optional).
                // File.WriteAllBytes("GeneratedReport.docx", reportStream.ToArray());
            }
        }
    }
}
