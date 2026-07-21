using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int OrderId { get; set; }
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { OrderId = 1, CustomerName = "Alice",   Total = 120.50m },
                new Order { OrderId = 2, CustomerName = "Bob",     Total =  85.75m },
                new Order { OrderId = 3, CustomerName = "Charlie", Total = 210.00m }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("------------------------------");

        // Table with order details.
        builder.Writeln("<<foreach [order in model.Orders]>>");
        var table = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Order ID");
        builder.InsertCell(); builder.Writeln("Customer");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        builder.InsertCell(); builder.Writeln("<<[order.OrderId]>>");
        builder.InsertCell(); builder.Writeln("<<[order.CustomerName]>>");
        builder.InsertCell(); builder.Writeln("<<[order.Total]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Summary:");
        // LINQ aggregation inside a variable tag expression.
        builder.Writeln("Total of all orders: <<[model.Orders.Sum(o => o.Total)]>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
