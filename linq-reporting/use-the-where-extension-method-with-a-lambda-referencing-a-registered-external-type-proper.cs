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
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create template document
        string templatePath = Path.Combine(outputDir, "ReportTemplate.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Orders Report ===");
        // Use model.Settings to expose external settings in the template
        builder.Writeln("Filter threshold (MinAmount): <<[model.Settings.MinAmount]>>");
        builder.Writeln();
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>, Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Sample data
        List<Order> allOrders = new()
        {
            new Order { CustomerName = "Alice", Amount = 120m },
            new Order { CustomerName = "Bob", Amount = 200m },
            new Order { CustomerName = "Charlie", Amount = 350m },
            new Order { CustomerName = "Diana", Amount = 80m }
        };

        // External settings object
        Settings settings = new() { MinAmount = 150m };

        // Advanced filtering using Where with external settings property
        List<Order> filteredOrders = allOrders
            .Where(o => o.Amount > settings.MinAmount)
            .ToList();

        // Model for the report, now includes Settings
        ReportModel model = new()
        {
            Orders = filteredOrders,
            Settings = settings
        };

        // Load template and build report
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model"
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(outputDir, "ReportOutput.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    public Settings Settings { get; set; } = new();
}

// External settings class
public class Settings
{
    public decimal MinAmount { get; set; }
}
