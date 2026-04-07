using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model representing the report root.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    // Master object – an order.
    public class Order
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; } = "";
        public List<LineItem> LineItems { get; set; } = new();
    }

    // Detail object – a line item of an order.
    public class LineItem
    {
        public string ProductName { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with nested LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders Report");
            builder.Writeln();

            // Outer foreach – iterate over the collection of orders.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Order ID: <<[order.OrderId]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");

            // Inner foreach – iterate over the line items of the current order.
            builder.Writeln("<<foreach [item in order.LineItems]>>");
            builder.Writeln("- <<[item.ProductName]>> x <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>"); // End inner foreach.

            builder.Writeln("<</foreach>>"); // End outer foreach.

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data (master‑detail relationship).
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new Order
                    {
                        OrderId = 1001,
                        CustomerName = "Alice Johnson",
                        LineItems = new List<LineItem>
                        {
                            new LineItem { ProductName = "Laptop", Quantity = 1 },
                            new LineItem { ProductName = "Mouse", Quantity = 2 }
                        }
                    },
                    new Order
                    {
                        OrderId = 1002,
                        CustomerName = "Bob Smith",
                        LineItems = new List<LineItem>
                        {
                            new LineItem { ProductName = "Desk Chair", Quantity = 1 },
                            new LineItem { ProductName = "Monitor", Quantity = 2 }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using ReportingEngine.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);
        }
    }
}
