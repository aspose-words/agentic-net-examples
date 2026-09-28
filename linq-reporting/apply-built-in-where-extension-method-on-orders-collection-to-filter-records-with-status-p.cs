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
        // Sample data with various statuses.
        List<Order> orders = new()
        {
            new Order { Id = 1, CustomerName = "Alice",   Status = "Pending",   Amount = 120.50m },
            new Order { Id = 2, CustomerName = "Bob",     Status = "Completed", Amount = 75.00m },
            new Order { Id = 3, CustomerName = "Charlie", Status = "Pending",   Amount = 200.00m },
            new Order { Id = 4, CustomerName = "Diana",   Status = "Shipped",   Amount = 50.25m }
        };

        // Apply LINQ Where to keep only pending orders.
        List<Order> pendingOrders = orders.Where(o => o.Status == "Pending").ToList();

        // Wrap the filtered collection in a model for the reporting engine.
        ReportModel model = new() { Orders = pendingOrders };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Pending Orders Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Status: <<[order.Status]>>");
        builder.Writeln("Amount: $<<[order.Amount]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new(templatePath);

        // Build the report using the filtered data.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
