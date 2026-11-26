using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public decimal Amount { get; set; }

    // Extension method replacement – must be an instance method for LINQ Reporting.
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
        // Prepare sample data
        var model = new ReportModel
        {
            Orders = new()
            {
                new Order { Id = 1, Amount = 750m },
                new Order { Id = 2, Amount = 1250m },
                new Order { Id = 3, Amount = 500m },
                new Order { Id = 4, Amount = 2000m }
            }
        };

        // Create template document
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // LINQ Reporting tags
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>   Amount: $<<[order.Amount]>>");
        builder.Writeln("<<if [order.IsHighValue(1000)]>>   **High Value**<</if>>");
        builder.Writeln("<</foreach>>");

        // Save template to disk
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load template for reporting
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build report using the model (root name "model")
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
