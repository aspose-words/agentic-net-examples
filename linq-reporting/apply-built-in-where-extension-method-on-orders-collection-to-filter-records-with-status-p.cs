using System;
using System.Collections.Generic;
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
        // Sample data.
        var orders = new List<Order>
        {
            new Order { Id = 1, CustomerName = "Alice",   Status = "Pending" },
            new Order { Id = 2, CustomerName = "Bob",     Status = "Shipped" },
            new Order { Id = 3, CustomerName = "Charlie", Status = "Pending" },
            new Order { Id = 4, CustomerName = "Diana",   Status = "Cancelled" }
        };

        // Apply LINQ Where to keep only pending orders.
        var pendingOrders = orders.Where(o => o.Status == "Pending").ToList();

        // Prepare the model for the reporting engine.
        var model = new ReportModel { Orders = pendingOrders };

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders with status 'Pending':");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Id: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Status: <<[order.Status]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the filtered data.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
