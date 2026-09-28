using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Stock { get; set; }
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Stock = 10, Price = 50m },
                new Item { Name = "Banana", Stock = 0, Price = 30m },
                new Item { Name = "Cherry", Stock = 5, Price = 120m },
                new Item { Name = "Date", Stock = 3, Price = 80m }
            }
        };

        // Create template document
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Items with Stock > 0 and Price < 100:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("- <<[item.Name]>>: Stock=<<[item.Stock]>>, Price=<<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template and build report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
