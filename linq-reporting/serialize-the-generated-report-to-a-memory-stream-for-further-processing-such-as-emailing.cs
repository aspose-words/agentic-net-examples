using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = "";
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
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Add a title.
            builder.Writeln("Order Report");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin a foreach loop over the collection of items.
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            Order sampleOrder = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Apple" },
                    new Item { Index = 2, Name = "Banana" },
                    new Item { Index = 3, Name = "Cherry" }
                }
            };

            // 3. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            bool success = engine.BuildReport(template, sampleOrder, "order");

            // 4. Serialize the generated report to a memory stream.
            using MemoryStream reportStream = new MemoryStream();
            template.Save(reportStream, SaveFormat.Docx);
            reportStream.Position = 0; // Reset for further processing (e.g., emailing).

            // Optional: write the stream to a file to verify the output.
            File.WriteAllBytes("GeneratedReport.docx", reportStream.ToArray());

            // Indicate completion.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Stream length: {reportStream.Length} bytes.");
        }
    }
}
