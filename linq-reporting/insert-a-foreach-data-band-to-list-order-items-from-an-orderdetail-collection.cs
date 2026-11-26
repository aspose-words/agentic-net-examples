using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class OrderDetail
{
    public int Index { get; set; }
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<OrderDetail> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Order order = new()
        {
            CustomerName = "Acme Corp",
            Items =
            {
                new OrderDetail { Index = 1, ProductName = "Widget", Quantity = 5 },
                new OrderDetail { Index = 2, ProductName = "Gadget", Quantity = 3 },
                new OrderDetail { Index = 3, ProductName = "Doohickey", Quantity = 7 }
            }
        };

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Header with customer name.
        builder.Writeln("Order Report for <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin foreach data band for order items.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Qty");
        builder.EndRow();

        // Table row for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.ProductName]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        // End table and foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for report generation.
        Document report = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, order, "order");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }
}
