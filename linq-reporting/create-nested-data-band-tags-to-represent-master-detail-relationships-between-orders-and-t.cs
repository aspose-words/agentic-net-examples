using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for the Table class

namespace AsposeWordsLinqReportingExample
{
    // Root data model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();

        public ReportModel()
        {
            // Sample data for demonstration.
            Orders.Add(new Order
            {
                CustomerName = "John Doe",
                Items = new List<LineItem>
                {
                    new LineItem { Product = "Laptop", Quantity = 1 },
                    new LineItem { Product = "Mouse", Quantity = 2 }
                }
            });

            Orders.Add(new Order
            {
                CustomerName = "Jane Smith",
                Items = new List<LineItem>
                {
                    new LineItem { Product = "Desk", Quantity = 1 },
                    new LineItem { Product = "Chair", Quantity = 4 },
                    new LineItem { Product = "Lamp", Quantity = 2 }
                }
            });
        }
    }

    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<LineItem> Items { get; set; } = new();
    }

    public class LineItem
    {
        public string Product { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Outer foreach iterates over orders.
            builder.Writeln("<<foreach [order in model.Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Inner foreach iterates over line items of the current order.
            builder.Writeln("<<foreach [item in order.Items]>>");
            Table table = builder.StartTable(); // Table comes from Aspose.Words.Tables

            // Header row (once per order).
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.EndRow();

            // Data row for each line item.
            builder.InsertCell();
            builder.Writeln("<<[item.Product]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.EndRow();

            builder.EndTable();
            builder.Writeln("<</foreach>>"); // End inner foreach (line items).

            builder.Writeln("<</foreach>>"); // End outer foreach (orders).

            // Save the template to disk.
            const string templatePath = "OrderReportTemplate.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report using the data model.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            // No special options required for this simple example.
            engine.BuildReport(report, model, "model");

            // Save the generated report.
            const string reportPath = "OrderReport.docx";
            report.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}
