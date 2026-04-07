using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model for a single order.
    public class Order
    {
        public int OrderId { get; set; }
        public int CustomerId { get; set; }
        public string Product { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Wrapper for a group of orders belonging to the same customer.
    public class CustomerGroup
    {
        public int CustomerId { get; set; }
        public List<Order> Orders { get; set; } = new();
    }

    // Root object passed to the reporting engine.
    public class ReportModel
    {
        public List<CustomerGroup> Groups { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample order data.
            List<Order> orders = new()
            {
                new Order { OrderId = 1, CustomerId = 100, Product = "Laptop", Quantity = 2 },
                new Order { OrderId = 2, CustomerId = 101, Product = "Mouse", Quantity = 5 },
                new Order { OrderId = 3, CustomerId = 100, Product = "Keyboard", Quantity = 1 },
                new Order { OrderId = 4, CustomerId = 102, Product = "Monitor", Quantity = 3 },
                new Order { OrderId = 5, CustomerId = 101, Product = "USB‑Cable", Quantity = 10 }
            };

            // 2. Group orders by CustomerId using LINQ.
            List<CustomerGroup> groups = orders
                .GroupBy(o => o.CustomerId)
                .Select(g => new CustomerGroup
                {
                    CustomerId = g.Key,
                    Orders = g.ToList()
                })
                .ToList();

            // 3. Create a template document programmatically.
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders Report");
            builder.Writeln();

            // Outer loop – iterate over each customer group.
            builder.Writeln("<<foreach [group in Groups]>>");
            builder.Writeln("Customer ID: <<[group.CustomerId]>>");
            builder.Writeln();

            // Inner loop – iterate over orders within the current group.
            builder.Writeln("<<foreach [order in group.Orders]>>");
            builder.Writeln("- Order <<[order.OrderId]>>: <<[order.Product]>> (Qty: <<[order.Quantity]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 4. Load the template and build the report.
            Document reportDoc = new Document(templatePath);
            ReportModel model = new ReportModel { Groups = groups };

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options

            // The root object name used in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // 5. Save the generated report.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
