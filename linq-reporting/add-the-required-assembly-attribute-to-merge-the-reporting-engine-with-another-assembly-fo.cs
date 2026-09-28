using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Order order = new()
        {
            CustomerName = "John Doe",
            Items = new()
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Cherry", Quantity = 7 }
            }
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");

        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        string outputPath = "OrderReport.docx";
        template.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
