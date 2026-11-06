using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public double DiscountPercentage { get; set; }
    public double DiscountAmount { get; set; }
    public double Total { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data
        var order = new Order
        {
            CustomerName = "John Doe",
            DiscountPercentage = 10.0,
            DiscountAmount = 15.0,
            Total = 150.0
        };

        // Create template document
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: $<<[order.Total]>>");
        // Conditional block: show discount only when percentage > 0
        builder.Writeln("<<if [order.DiscountPercentage > 0]>>Discount: <<[order.DiscountPercentage]>>% (-$<<[order.DiscountAmount]>>)<</if>>");

        // Save template
        doc.Save(templatePath);

        // Load template for report generation
        var reportDoc = new Document(templatePath);

        // Build report
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save final report
        var outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
