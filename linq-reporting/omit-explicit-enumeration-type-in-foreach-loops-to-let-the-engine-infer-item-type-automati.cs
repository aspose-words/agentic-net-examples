using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
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
        var model = new ReportModel();
        model.Products.Add(new Product { Id = 1, Name = "Apple" });
        model.Products.Add(new Product { Id = 2, Name = "Banana" });
        model.Products.Add(new Product { Id = 3, Name = "Cherry" });

        // Create a template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // The foreach tag will iterate over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product ID: <<[p.Id]>>");
        builder.Writeln("Product Name: <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the root data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");

        // Example of a C# foreach loop without an explicit type (using var).
        foreach (var product in model.Products)
        {
            Console.WriteLine($"Processed product {product.Id}: {product.Name}");
        }
    }
}
