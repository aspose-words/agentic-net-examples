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
    public decimal Amount { get; set; }
}

public class CustomerGroup
{
    public int CustomerId { get; set; }
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
        // Sample data
        List<Order> orders = new()
        {
            new Order { OrderId = 1, CustomerId = 101, Amount = 250.00m },
            new Order { OrderId = 2, CustomerId = 102, Amount = 150.50m },
            new Order { OrderId = 3, CustomerId = 101, Amount = 99.99m },
            new Order { OrderId = 4, CustomerId = 103, Amount = 300.00m },
            new Order { OrderId = 5, CustomerId = 102, Amount = 75.25m }
        };

        // Group orders by CustomerId using built‑in LINQ GroupBy
        ReportModel model = new()
        {
            CustomerGroups = orders
                .GroupBy(o => o.CustomerId)
                .Select(g => new CustomerGroup
                {
                    CustomerId = g.Key,
                    Orders = g.ToList()
                })
                .ToList()
        };

        // Create template document programmatically
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Orders grouped by Customer:");
        builder.Writeln("<<foreach [group in CustomerGroups]>>");
        builder.Writeln("Customer ID: <<[group.CustomerId]>>");
        builder.Writeln("<<foreach [ord in group.Orders]>>");
        builder.Writeln(" - Order ID: <<[ord.OrderId]>>, Amount: $<<[ord.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new(templatePath);

        // Build the report
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
