using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Bread", Price = 2.49m },
                new Product { Name = "Milk", Price = 1.79m }
            }
        };

        // Create a Word document template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product List:");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln(" - <<[p.Name]>>: <<[p.Price.ToString(\"C\")]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
