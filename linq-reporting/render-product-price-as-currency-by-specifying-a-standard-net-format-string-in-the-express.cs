using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product List");
        builder.Writeln("<<foreach [p in Products]>>");

        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price.ToString(\"C\")]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Prepare data.
        ReportModel model = new()
        {
            Products = new()
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Build report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
