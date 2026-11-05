using System;
using System.Collections.Generic;
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

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { OrderId = 1, CustomerId = 100, CustomerName = "Alice", Amount = 250.00m },
                new Order { OrderId = 2, CustomerId = 101, CustomerName = "Bob",   Amount = 150.50m },
                new Order { OrderId = 3, CustomerId = 100, CustomerName = "Alice", Amount = 99.99m },
                new Order { OrderId = 4, CustomerId = 102, CustomerName = "Carol", Amount = 300.00m },
                new Order { OrderId = 5, CustomerId = 101, CustomerName = "Bob",   Amount = 75.25m }
            }
        };

        // Create template.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Group orders by CustomerId using LINQ GroupBy inside the template.
        builder.Writeln("<<foreach [g in model.Orders.GroupBy(o => o.CustomerId)]>>");
        builder.Writeln("Customer ID: <<[g.Key]>>");
        builder.Writeln("Customer Name: <<[g.First().CustomerName]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in g]>>");
        builder.Writeln("- Order ID: <<[o.OrderId]>>, Amount: $<<[o.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
