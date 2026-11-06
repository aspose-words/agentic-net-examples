using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; } = 0;
    public string CustomerName { get; set; } = "";
    public decimal Total { get; set; } = 0m;
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
                new Order { Id = 1, CustomerName = "Alice", Total = 120.50m },
                new Order { Id = 2, CustomerName = "Bob",   Total = 75.00m },
                new Order { Id = 3, CustomerName = "Carol", Total = 210.25m }
            }
        };

        // Create the template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Order Summary");
        builder.Writeln("Total of all orders: <<[model.Orders.Sum(o => o.Total)]>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = Path.Combine("output", "Report.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        doc.Save(outputPath);
    }
}
