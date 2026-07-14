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
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    CustomerName = "Alice Johnson",
                    TotalAmount = 250.00m,
                    Discount = 25.00m
                },
                new Order
                {
                    CustomerName = "Bob Smith",
                    TotalAmount = 180.00m,
                    Discount = 0.00m
                }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("------------------------------");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total Amount: $<<[order.TotalAmount]>>");
        builder.Writeln("<<if [order.Discount > 0]>>Discount: $<<[order.Discount]>> (<<[order.Discount / order.TotalAmount * 100]>>%)<</if>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("------------------------------");

        // Save the template (optional, for inspection).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal TotalAmount { get; set; }
    public decimal Discount { get; set; }
}
