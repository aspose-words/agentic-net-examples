using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
    public string Status { get; set; } = "";
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Pending Orders:");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("- <<[order.Id]>>: <<[order.CustomerName]>> (<<[order.Status]>>)");
        builder.Writeln("<</foreach>>");
        doc.Save(templatePath);

        // Load the template document.
        var template = new Document(templatePath);

        // Sample data with various order statuses.
        var orders = new List<Order>
        {
            new() { Id = 1, CustomerName = "Alice", Status = "Pending" },
            new() { Id = 2, CustomerName = "Bob", Status = "Shipped" },
            new() { Id = 3, CustomerName = "Charlie", Status = "Pending" },
            new() { Id = 4, CustomerName = "Diana", Status = "Cancelled" }
        };

        // Apply Where to filter only pending orders.
        var pendingOrders = orders.Where(o => o.Status == "Pending").ToList();

        // Prepare the root model for the report.
        var model = new ReportModel { Orders = pendingOrders };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
