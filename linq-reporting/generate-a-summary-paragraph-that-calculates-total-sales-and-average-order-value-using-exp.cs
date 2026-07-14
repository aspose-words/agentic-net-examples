using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public double Amount { get; set; }
}

public class ReportModel
{
    // Collection of orders.
    public List<Order> Orders { get; set; } = new();

    // Calculated total sales.
    public double TotalSales => Orders.Sum(o => o.Amount);

    // Calculated average order value.
    public double AverageOrderValue => Orders.Any() ? Orders.Average(o => o.Amount) : 0;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Orders =
            {
                new Order { Amount = 120.50 },
                new Order { Amount = 75.00 },
                new Order { Amount = 200.25 },
                new Order { Amount = 50.75 }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Sales Summary");
        builder.Writeln("Total Sales: <<[model.TotalSales]>>");
        builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("SalesReport.docx");
    }
}
