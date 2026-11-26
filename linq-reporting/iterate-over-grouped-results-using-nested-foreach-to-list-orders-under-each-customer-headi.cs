using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedForeach
{
    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<CustomerGroup> CustomerGroups { get; set; } = new();
    }

    // Represents a customer and its related orders.
    public class CustomerGroup
    {
        public string CustomerName { get; set; } = "";
        public List<Order> Orders { get; set; } = new();
    }

    // Simple order data.
    public class Order
    {
        public int OrderId { get; set; }
        public string Product { get; set; } = "";
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var orders = new List<Order>
            {
                new() { OrderId = 1, Product = "Laptop", Amount = 1200.00m, },
                new() { OrderId = 2, Product = "Mouse", Amount = 25.50m, },
                new() { OrderId = 3, Product = "Keyboard", Amount = 45.00m, },
                new() { OrderId = 4, Product = "Monitor", Amount = 300.00m, },
                new() { OrderId = 5, Product = "Desk", Amount = 150.00m, }
            };

            // Assign customers to orders.
            var orderAssignments = new List<(string Customer, Order Order)>
            {
                ("Alice", orders[0]),
                ("Bob", orders[1]),
                ("Alice", orders[2]),
                ("Bob", orders[3]),
                ("Charlie", orders[4])
            };

            // Group orders by customer name.
            var model = new ReportModel
            {
                CustomerGroups = orderAssignments
                    .GroupBy(x => x.Customer)
                    .Select(g => new CustomerGroup
                    {
                        CustomerName = g.Key,
                        Orders = g.Select(x => x.Order).ToList()
                    })
                    .ToList()
            };

            // Create the template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Add a title.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
            builder.Writeln("Orders Report");
            builder.ParagraphFormat.ClearFormatting();

            // Begin outer foreach over customer groups.
            builder.Writeln("<<foreach [group in CustomerGroups]>>");
            builder.Writeln("Customer: <<[group.CustomerName]>>");
            builder.Writeln("");

            // Begin inner foreach over orders of the current customer.
            builder.Writeln("<<foreach [order in group.Orders]>>");
            builder.Writeln("- Order ID: <<[order.OrderId]>>, Product: <<[order.Product]>>, Amount: $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>"); // End inner foreach

            builder.Writeln("<</foreach>>"); // End outer foreach

            // Save the template.
            doc.Save(templatePath);

            // Load the template for reporting.
            var templateDoc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            templateDoc.Save(outputPath);
        }
    }
}
