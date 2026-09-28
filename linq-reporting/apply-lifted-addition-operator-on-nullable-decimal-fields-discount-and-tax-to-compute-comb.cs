using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with nullable decimal fields.
        var order = new Order
        {
            Discount = 5.5m,   // non‑null value
            Tax = null         // null value to demonstrate lifted addition
        };

        // Create a Word template programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Discount: <<[order.Discount]>>");
        builder.Writeln("Tax: <<[order.Tax]>>");
        builder.Writeln("Combined (Discount + Tax): <<[order.Discount + order.Tax]>>");

        doc.Save(templatePath);

        // Load the template (optional, can reuse the same Document instance).
        var template = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        var reportPath = "Report.docx";
        template.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}

// Data model with nullable decimal fields.
public class Order
{
    public decimal? Discount { get; set; } = 0m;
    public decimal? Tax { get; set; } = 0m;
}
