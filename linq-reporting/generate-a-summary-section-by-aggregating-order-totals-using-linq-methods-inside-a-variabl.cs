using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public decimal Total { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln("----------------");

        // Repeat each order.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>> - Total: $<<[order.Total]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Summary:");
        builder.Writeln("Total Orders: <<[Orders.Count]>>");
        // Aggregate order totals using LINQ inside a tag expression.
        builder.Writeln("Grand Total: $<<[Orders.Sum(o => o.Total)]>>");

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Orders.Add(new Order { CustomerName = "Alice", Total = 120.50m });
        model.Orders.Add(new Order { CustomerName = "Bob", Total = 75.00m });
        model.Orders.Add(new Order { CustomerName = "Charlie", Total = 210.25m });

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("OrderReport.docx");
    }
}
