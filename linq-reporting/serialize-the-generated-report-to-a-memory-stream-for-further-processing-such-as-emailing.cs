using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

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
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Step 1: Create the template document programmatically
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title with a LINQ Reporting tag
            builder.Writeln("Order Report for: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach block that will repeat the whole table for each item
            builder.Writeln("<<foreach [item in order.Items]>>");

            // Start the table inside the foreach block
            Table table = builder.StartTable();

            // Header row
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row (repeated for each item)
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End the table
            builder.EndTable();

            // End foreach block
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // Step 2: Load the template for report generation
            var reportDoc = new Document(templatePath);

            // Step 3: Prepare sample data
            var order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new() { Index = 1, Name = "Item A" },
                    new() { Index = 2, Name = "Item B" },
                    new() { Index = 3, Name = "Item C" }
                }
            };

            // Step 4: Build the report using LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Step 5: Serialize the generated report to a memory stream
            using var memoryStream = new MemoryStream();
            reportDoc.Save(memoryStream, SaveFormat.Docx);
            memoryStream.Position = 0; // Reset position for further processing

            Console.WriteLine($"Report generated. Memory stream length: {memoryStream.Length} bytes");

            // Optional: Save the generated report to a file for verification
            var outputPath = "GeneratedReport.docx";
            reportDoc.Save(outputPath);
        }
    }
}
