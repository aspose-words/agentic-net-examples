using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for an order.
    public class Order
    {
        public int OrderId { get; set; }
        public int CustomerId { get; set; }
        public string Product { get; set; } = "";
        public double Price { get; set; }
    }

    // Wrapper model that contains the collection of orders and the grouped view.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
        public List<OrderGroup> Groups { get; set; } = new();
    }

    // Represents a group of orders belonging to a single customer.
    public class OrderGroup
    {
        public int CustomerId { get; set; }
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Iterate over each customer group.
            builder.Writeln("<<foreach [group in model.Groups]>>");
            builder.Writeln("Customer ID: <<[group.CustomerId]>>");
            // Iterate over the orders within the current group.
            builder.Writeln("<<foreach [order in group.Orders]>>");
            builder.Writeln("- Order ID: <<[order.OrderId]>>, Product: <<[order.Product]>>, Price: $<<[order.Price]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // 3. Prepare sample data.
            var orders = new List<Order>
            {
                new Order { OrderId = 1, CustomerId = 1, Product = "Apple",  Price = 1.20 },
                new Order { OrderId = 2, CustomerId = 1, Product = "Banana", Price = 0.80 },
                new Order { OrderId = 3, CustomerId = 2, Product = "Carrot", Price = 0.50 },
                new Order { OrderId = 4, CustomerId = 2, Product = "Doughnut", Price = 1.50 },
                new Order { OrderId = 5, CustomerId = 3, Product = "Eggplant", Price = 1.00 }
            };

            // Group orders by CustomerId using LINQ.
            var groups = orders
                .GroupBy(o => o.CustomerId)
                .Select(g => new OrderGroup
                {
                    CustomerId = g.Key,
                    Orders = g.ToList()
                })
                .ToList();

            var model = new ReportModel
            {
                Orders = orders,
                Groups = groups
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. Save the generated report.
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);
        }
    }
}
