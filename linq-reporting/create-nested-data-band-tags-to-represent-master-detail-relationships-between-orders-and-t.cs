using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "OrderReportTemplate.docx";
        const string outputPath = "OrderReportResult.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Master band – iterate over orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Line Items:");

        // Detail band – iterate over line items of the current order.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Simple table for each line item.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[item.ProductName]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>"); // End of line‑items foreach.
        builder.Writeln("<</foreach>>"); // End of orders foreach.

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model (master‑detail relationship).
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    CustomerName = "Alice Johnson",
                    Items = new List<LineItem>
                    {
                        new LineItem { ProductName = "Laptop", Quantity = 1 },
                        new LineItem { ProductName = "Mouse", Quantity = 2 }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    CustomerName = "Bob Smith",
                    Items = new List<LineItem>
                    {
                        new LineItem { ProductName = "Desk Chair", Quantity = 1 },
                        new LineItem { ProductName = "Monitor", Quantity = 2 },
                        new LineItem { ProductName = "Keyboard", Quantity = 1 }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object name in the template is "Orders", so we pass the model
        // and reference it with the name "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model definitions (public, non‑nullable, initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = "";
    public List<LineItem> Items { get; set; } = new();
}

public class LineItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
}
