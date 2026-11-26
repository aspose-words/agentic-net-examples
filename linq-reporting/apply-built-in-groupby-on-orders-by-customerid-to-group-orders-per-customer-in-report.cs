using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int OrderId { get; set; }
    public int CustomerId { get; set; }
    public string CustomerName { get; set; } = "";
    public decimal Amount { get; set; }
}

public class CustomerGroup
{
    public int CustomerId { get; set; }
    public string CustomerName { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

public class ReportModel
{
    public List<CustomerGroup> CustomerGroups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var orders = new List<Order>
        {
            new Order { OrderId = 1, CustomerId = 101, CustomerName = "Alice",   Amount = 120.50m },
            new Order { OrderId = 2, CustomerId = 102, CustomerName = "Bob",     Amount = 75.00m },
            new Order { OrderId = 3, CustomerId = 101, CustomerName = "Alice",   Amount = 200.00m },
            new Order { OrderId = 4, CustomerId = 103, CustomerName = "Charlie", Amount = 50.25m },
            new Order { OrderId = 5, CustomerId = 102, CustomerName = "Bob",     Amount = 180.75m }
        };

        // Group orders by CustomerId.
        var model = new ReportModel
        {
            CustomerGroups = orders
                .GroupBy(o => new { o.CustomerId, o.CustomerName })
                .Select(g => new CustomerGroup
                {
                    CustomerId = g.Key.CustomerId,
                    CustomerName = g.Key.CustomerName,
                    Orders = g.ToList()
                })
                .ToList()
        };

        // Create template document.
        string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [c in CustomerGroups]>>");
        builder.Writeln("Customer: <<[c.CustomerName]>> (ID: <<[c.CustomerId]>>)");
        builder.Writeln("<<foreach [o in c.Orders]>>");
        builder.Writeln("- Order ID: <<[o.OrderId]>>, Amount: <<[o.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template and build report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
