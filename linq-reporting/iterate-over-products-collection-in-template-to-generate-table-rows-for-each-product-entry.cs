using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Create a template document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in model.Products]>>");

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Table row for each product.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
