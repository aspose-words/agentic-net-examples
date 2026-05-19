using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public double Total { get; set; }
    public string Date { get; set; } = "";
    public string TotalFormatted => Total.ToString("C");
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    public string GeneratedOn { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Orders =
            {
                new Order
                {
                    CustomerName = "Alice Johnson",
                    Total = 1234.56,
                    Date = "2023-08-01"
                },
                new Order
                {
                    CustomerName = "Bob Smith",
                    Total = 7890.12,
                    Date = "2023-08-02"
                }
            }
        };

        // Create the template document programmatically
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("Generated on: <<[model.GeneratedOn]>>");
        builder.Writeln();

        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.Date]>>");
        builder.Writeln("Total: <<[order.TotalFormatted]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report using the model
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
