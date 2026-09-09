using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public double Price { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel();
        model.Products.Add(new Product { Name = "Apple", Price = 4.0 });
        model.Products.Add(new Product { Name = "Banana", Price = 9.0 });
        model.Products.Add(new Product { Name = "Cherry", Price = 16.0 });

        // Create a blank template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        builder.Writeln("Price: <<[p.Price]>>");
        // Use Math.Sqrt via KnownTypes.
        builder.Writeln("Square root of price: <<[Math.Sqrt(p.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register System.Math so static methods can be used in expressions.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
