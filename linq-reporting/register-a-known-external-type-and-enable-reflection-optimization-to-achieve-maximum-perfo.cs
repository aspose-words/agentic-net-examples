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
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data
        ReportModel model = new()
        {
            Orders = new()
        };
        for (int i = 1; i <= 10000; i++)
        {
            model.Orders.Add(new Order
            {
                Id = i,
                CustomerName = $"Customer {i}",
                Amount = Math.Round(1000 * new Random(i).NextDouble(), 2)
            });
        }

        // Build template document programmatically
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Id: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load template for reporting
        Document reportDoc = new(templatePath);

        // Enable reflection optimization and register known external type
        ReportingEngine.UseReflectionOptimization = true;
        ReportingEngine engine = new();
        engine.KnownTypes.Add(typeof(Order));

        // Build the report
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Wrapper model for the report
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Sample data class
public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public double Amount { get; set; }
}
