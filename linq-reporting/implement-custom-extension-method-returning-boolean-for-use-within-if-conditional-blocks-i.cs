using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        Order sampleOrder = new()
        {
            Id = 12345,
            CustomerName = "John Doe",
            TotalAmount = 1500.75m,
            Items = new()
            {
                new OrderItem { Name = "Laptop", Price = 1200.00m },
                new OrderItem { Name = "Mouse", Price = 25.50m },
                new OrderItem { Name = "Keyboard", Price = 75.25m }
            }
        };

        // Create template document programmatically
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total Amount: $<<[order.TotalAmount]>>");
        builder.Writeln();
        builder.Writeln("<<if [order.IsHighValue()]>>");
        builder.Writeln("** This is a high‑value order! **");
        builder.Writeln("<</if>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new(templatePath);

        // Build the report using the custom method in the conditional
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(reportDoc, sampleOrder, "order");

        // Save the generated report
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Data model classes
public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public decimal TotalAmount { get; set; }
    public List<OrderItem> Items { get; set; } = new();

    // Method used in the template conditional
    public bool IsHighValue()
    {
        // Define high value as total amount greater than 1000
        return TotalAmount > 1000m;
    }
}

public class OrderItem
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
