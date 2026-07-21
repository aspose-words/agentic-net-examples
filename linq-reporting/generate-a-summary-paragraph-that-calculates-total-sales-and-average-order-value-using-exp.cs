using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public decimal Amount { get; set; }

    public Order(decimal amount)
    {
        Amount = amount;
    }
}

public class ReportModel
{
    // Collection of orders.
    public List<Order> Orders { get; set; } = new();

    // Total sales calculated from the Orders collection.
    public decimal TotalSales => Orders.Sum(o => o.Amount);

    // Average order value; returns 0 when there are no orders to avoid division by zero.
    public decimal AverageOrderValue => Orders.Count > 0 ? Orders.Average(o => o.Amount) : 0m;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "SalesTemplate.docx");
        string reportPath   = Path.Combine(Environment.CurrentDirectory, "SalesReport.docx");

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Sales Summary");
        builder.Writeln();

        // Insert expression tags that will be replaced by the reporting engine.
        builder.Writeln("Total Sales: <<[model.TotalSales]>>");
        builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new ReportModel();
        model.Orders.Add(new Order(120.50m));
        model.Orders.Add(new Order(75.00m));
        model.Orders.Add(new Order(210.30m));
        model.Orders.Add(new Order(55.20m));

        // -------------------------------------------------
        // 3. Load the template and build the report.
        // -------------------------------------------------
        Document docToReport = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple scenario.
        engine.BuildReport(docToReport, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        docToReport.Save(reportPath);
    }
}
