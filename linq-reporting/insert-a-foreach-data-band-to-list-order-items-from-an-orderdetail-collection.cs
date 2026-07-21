using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<OrderDetail> Details { get; set; } = new();
    }

    public class OrderDetail
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Step 1: Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Static text with a placeholder for the customer's name.
            builder.Writeln("Order Report for <<[order.CustomerName]>>");
            builder.Writeln(); // empty line

            // Begin a foreach data band that iterates over Order.Details.
            builder.Writeln("<<foreach [detail in order.Details]>>");

            // Create a table header.
            var table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Row for each order detail.
            builder.InsertCell();
            builder.Writeln("<<[detail.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[detail.Quantity]>>");
            builder.InsertCell();
            builder.Writeln("<<[detail.Price]>>");
            builder.EndRow();

            // End the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "OrderTemplate.docx";
            template.Save(templatePath);

            // Step 2: Prepare sample data.
            var order = new Order
            {
                CustomerName = "John Doe",
                Details = new List<OrderDetail>
                {
                    new() { Name = "Apple", Quantity = 3, Price = 0.5 },
                    new() { Name = "Banana", Quantity = 5, Price = 0.3 },
                    new() { Name = "Cherry", Quantity = 10, Price = 0.2 }
                }
            };

            // Step 3: Load the template and build the report.
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            bool success = engine.BuildReport(reportDoc, order, "order");

            // Step 4: Save the generated report.
            const string outputPath = "OrderReport.docx";
            reportDoc.Save(outputPath);
        }
    }
}
