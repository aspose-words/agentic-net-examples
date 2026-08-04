using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var report = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice Johnson", Amount = 250.75m },
                new Order { Id = 2, CustomerName = "Bob Smith", Amount = 120.00m },
                new Order { Id = 3, CustomerName = "Carol Davis", Amount = 560.40m }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln();
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: $<<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(templateDoc, report, "report");

        // Save the generated report.
        var outputPath = "ReportOutput.docx";
        templateDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Wrapper class for the root data source.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Sample business object.
public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}
