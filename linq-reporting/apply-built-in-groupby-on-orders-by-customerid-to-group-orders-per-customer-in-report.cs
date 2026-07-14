using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Represents a single order.
    public class Order
    {
        public int OrderId { get; set; }
        public int CustomerId { get; set; }
        public string Product { get; set; } = string.Empty;
        public double Amount { get; set; }
    }

    // Represents a group of orders belonging to a specific customer.
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
            // Step 1: Prepare sample order data.
            List<Order> orders = new()
            {
                new Order { OrderId = 1, CustomerId = 101, Product = "Laptop", Amount = 1200.00 },
                new Order { OrderId = 2, CustomerId = 102, Product = "Smartphone", Amount = 800.00 },
                new Order { OrderId = 3, CustomerId = 101, Product = "Mouse", Amount = 25.50 },
                new Order { OrderId = 4, CustomerId = 103, Product = "Keyboard", Amount = 45.00 },
                new Order { OrderId = 5, CustomerId = 102, Product = "Monitor", Amount = 300.00 }
            };

            // Step 2: Group orders by CustomerId using LINQ.
            ReportModel model = new()
            {
                Groups = orders
                    .GroupBy(o => o.CustomerId)
                    .Select(g => new CustomerGroup
                    {
                        CustomerId = g.Key,
                        Orders = g.ToList()
                    })
                    .ToList()
            };

            // Step 3: Create a template document with LINQ Reporting tags.
            const string templatePath = "Template.docx";
            DocumentBuilder builder = new();
            builder.Writeln("<<foreach [group in Groups]>>");
            builder.Writeln("Customer ID: <<[group.CustomerId]>>");
            builder.Writeln("Orders:");
            builder.Writeln("<<foreach [order in group.Orders]>>");
            builder.Writeln("- Order <<[order.OrderId]>>: <<[order.Product]>> – $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // Step 4: Load the template and build the report.
            Document doc = new(templatePath);
            ReportingEngine engine = new();
            engine.BuildReport(doc, model, "model");

            // Step 5: Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);

            // Optional: Inform that the process completed (no interactive prompts required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}
