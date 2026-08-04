using System;
using System.Collections.Generic;
using System.Linq;
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
        // Sample data.
        var allProducts = new List<Product>
        {
            new() { Name = "Apple",  Price = 10m },
            new() { Name = "Banana", Price = 25m },
            new() { Name = "Cherry", Price = 30m },
            new() { Name = "Date",   Price = 15m },
            new() { Name = "Elderberry", Price = 50m }
        };

        // Combine Where, Select, and OrderBy in a single LINQ expression.
        var filteredSorted = allProducts
            .Where(p => p.Price > 20)                     // filter
            .Select(p => new Product { Name = p.Name, Price = p.Price }) // project
            .OrderByDescending(p => p.Price)             // sort
            .ToList();

        // Prepare the model for the reporting engine.
        var model = new ReportModel { Products = filteredSorted };

        // Create a Word document template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Filtered and Sorted Products:");
        // LINQ Reporting foreach tag iterating over the Products collection.
        builder.Writeln("<<foreach [p in model.Products]>>");
        builder.Writeln(" - <<[p.Name]>> : $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
