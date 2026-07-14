using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public Order(string customerName, decimal total)
    {
        CustomerName = customerName;
        Total = total;
    }

    public string CustomerName { get; set; } = "";
    public decimal Total { get; set; }
}

public class ReportModel
{
    public ReportModel()
    {
        Orders = new List<Order>();
    }

    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportModel
        {
            Orders = new()
            {
                new Order("Alice", 120.50m),
                new Order("Bob",   75.00m),
                new Order("Carol", 210.30m)
            }
        };

        // Build the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Summary");
        // LINQ aggregation inside a variable tag expression.
        builder.Writeln("Total Amount: <<[model.Orders.Sum(o => o.Total)]>>");

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("Report.docx");
    }
}
