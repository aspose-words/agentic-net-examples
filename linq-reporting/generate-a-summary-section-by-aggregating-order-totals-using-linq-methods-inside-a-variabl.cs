using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        ReportModel model = new()
        {
            Orders = new()
            {
                new Order { CustomerName = "Alice", Total = 120.50m },
                new Order { CustomerName = "Bob", Total = 75.00m },
                new Order { CustomerName = "Charlie", Total = 210.30m }
            }
        };

        // Create template document
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln();

        // List each order
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>> - Total: $<<[order.Total]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Summary section using LINQ aggregation inside a tag expression
        builder.Writeln("Summary:");
        builder.Writeln("Total Orders: <<[Orders.Count]>>");
        builder.Writeln("Grand Total: $<<[Orders.Sum(o => o.Total)]>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Total { get; set; }
}
