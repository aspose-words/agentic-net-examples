using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Customer name for the order.
    public string CustomerName { get; set; } = string.Empty;

    // Amount of the order.
    public double Amount { get; set; }
}

public class ReportModel
{
    // Collection of orders.
    public List<Order> Orders { get; set; } = new();

    // Total sales calculated from the orders.
    public double TotalSales => Orders.Sum(o => o.Amount);

    // Average order value calculated from the orders.
    public double AverageOrderValue => Orders.Any() ? Orders.Average(o => o.Amount) : 0;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "SalesSummaryTemplate.docx";
        string reportPath = "SalesSummaryReport.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Sales Summary");
        builder.Writeln("Total Sales: <<[model.TotalSales]>>");
        builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for reporting.
        // -------------------------------------------------
        Document reportDocument = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample data.
        // -------------------------------------------------
        ReportModel model = new ReportModel();
        model.Orders.Add(new Order { CustomerName = "Alice", Amount = 120.5 });
        model.Orders.Add(new Order { CustomerName = "Bob", Amount = 250.0 });
        model.Orders.Add(new Order { CustomerName = "Charlie", Amount = 75.25 });

        // -------------------------------------------------
        // 4. Build the report using LINQ Reporting Engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple scenario.
        engine.BuildReport(reportDocument, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDocument.Save(reportPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
    }
}
