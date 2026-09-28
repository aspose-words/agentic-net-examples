using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string Id { get; set; } = "";
    public decimal Amount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document with LINQ Reporting tags
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Sales Summary:");
        builder.Writeln("Total Sales: <<[Orders.Sum(o => o.Amount)]>>");
        builder.Writeln("Average Order Value: <<[Orders.Average(o => o.Amount)]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        var doc = new Document(templatePath);

        // Sample data
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Id = "001", Amount = 120.50m },
                new Order { Id = "002", Amount = 75.00m },
                new Order { Id = "003", Amount = 200.00m }
            }
        };

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);
    }
}
