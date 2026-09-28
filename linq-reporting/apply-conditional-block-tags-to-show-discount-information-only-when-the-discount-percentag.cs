using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create the template document with LINQ Reporting tags
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a simple report that lists orders and shows discount only when > 0
        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("<<if [order.DiscountPercent > 0]>>Discount: <<[order.DiscountPercent]>>%<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load the template for report generation
        Document doc = new Document(templatePath);

        // Prepare sample data
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, DiscountPercent = 15 },
                new Order { Id = 2, DiscountPercent = 0 },
                new Order { Id = 3, DiscountPercent = 5 }
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int Id { get; set; }
    public int DiscountPercent { get; set; }
}
