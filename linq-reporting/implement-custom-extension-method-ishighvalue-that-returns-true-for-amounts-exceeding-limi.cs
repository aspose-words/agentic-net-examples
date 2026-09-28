using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string Description { get; set; } = string.Empty;
    public decimal Amount { get; set; }

    // Returns true when the order amount exceeds the supplied limit.
    public bool IsHighValue(decimal limit) => Amount > limit;
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach over the Orders collection.
        builder.Writeln("<<foreach [order in model.Orders]>>");
        // Output order description and amount.
        builder.Writeln("Order: <<[order.Description]>> Amount: <<[order.Amount]>>");
        // Use the custom method to test for high‑value orders.
        builder.Writeln("<<if [order.IsHighValue(1000)]>>");
        builder.Writeln(" - High value order!");
        builder.Writeln("<</if>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data.
        ReportModel model = new()
        {
            Orders = new()
            {
                new Order { Description = "Standard Item", Amount = 250m },
                new Order { Description = "Premium Item", Amount = 1500m },
                new Order { Description = "Basic Item", Amount = 75m }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
