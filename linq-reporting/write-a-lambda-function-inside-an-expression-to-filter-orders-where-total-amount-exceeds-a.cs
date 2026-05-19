using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
    public decimal TotalAmount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice", TotalAmount = 75m },
                new Order { Id = 2, CustomerName = "Bob",   TotalAmount = 150m },
                new Order { Id = 3, CustomerName = "Carol", TotalAmount = 250m },
                new Order { Id = 4, CustomerName = "Dave",  TotalAmount = 45m }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders with total amount > 100:");
        // Use a lambda expression inside the foreach tag to filter the collection.
        builder.Writeln("<<foreach [order in Orders.Where(o => o.TotalAmount > 100)]>>");
        builder.Writeln("- <<[order.Id]>>: <<[order.CustomerName]>> - $<<[order.TotalAmount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
