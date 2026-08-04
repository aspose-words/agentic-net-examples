using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportModel
{
    public string Title { get; set; } = "Sales Report";
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<foreach [p in model.Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        builder.Writeln("Quantity: <<[p.Quantity]>>");
        builder.Writeln("Price: $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Products.Add(new Product { Name = "Apple", Quantity = 10, Price = 0.5m });
        model.Products.Add(new Product { Name = "Banana", Quantity = 5, Price = 0.3m });
        model.Products.Add(new Product { Name = "Orange", Quantity = 8, Price = 0.6m });

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Write the generated document to a memory stream (simulating an HTTP response body).
        using MemoryStream responseBody = new MemoryStream();
        template.Save(responseBody, SaveFormat.Docx);
        Console.WriteLine($"Report generated, size: {responseBody.Length} bytes");
    }
}
