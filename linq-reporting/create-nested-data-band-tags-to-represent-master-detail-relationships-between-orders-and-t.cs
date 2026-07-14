using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create the LINQ Reporting template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Master (orders) band.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("");

        // Detail (line items) band.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.LineItems]>>");
        builder.Writeln("- <<[item.ProductName]>>  Qty: <<[item.Quantity]>>  Price: $<<[item.UnitPrice]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("");
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        const string templatePath = "ReportTemplate.docx";
        template.Save(templatePath);

        // 2. Prepare sample data.
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    CustomerName = "John Doe",
                    OrderDate = new DateTime(2023, 5, 21),
                    LineItems = new List<LineItem>
                    {
                        new LineItem { ProductName = "Apple", Quantity = 5, UnitPrice = 0.60m },
                        new LineItem { ProductName = "Banana", Quantity = 3, UnitPrice = 0.40m }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    CustomerName = "Jane Smith",
                    OrderDate = new DateTime(2023, 5, 22),
                    LineItems = new List<LineItem>
                    {
                        new LineItem { ProductName = "Orange", Quantity = 4, UnitPrice = 0.55m },
                        new LineItem { ProductName = "Grapes", Quantity = 2, UnitPrice = 2.00m },
                        new LineItem { ProductName = "Mango", Quantity = 1, UnitPrice = 1.50m }
                    }
                }
            }
        };

        // 3. Build the report using the ReportingEngine.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.BuildReport(report, model, "model");

        // 4. Save the generated report.
        const string outputPath = "ReportResult.docx";
        report.Save(outputPath);
    }
}

// Wrapper class that matches the root data source name used in BuildReport.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Master object.
public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public DateTime OrderDate { get; set; }
    public List<LineItem> LineItems { get; set; } = new();
}

// Detail object.
public class LineItem
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
