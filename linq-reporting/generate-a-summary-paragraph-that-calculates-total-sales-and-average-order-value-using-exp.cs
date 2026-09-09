using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Amount = 120.50m },
                new Order { Amount = 75.00m },
                new Order { Amount = 200.25m }
            }
        };
        // Compute aggregates.
        model.TotalSales = model.Orders.Sum(o => o.Amount);
        model.AverageOrderValue = model.Orders.Average(o => o.Amount);

        // Create the template document programmatically.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Summary Report");
        builder.Writeln("Total Sales: <<[model.TotalSales]>>");
        builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Data model exposed to the template.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    public decimal TotalSales { get; set; }
    public decimal AverageOrderValue { get; set; }
}

// Simple order entity.
public class Order
{
    public decimal Amount { get; set; }
}
