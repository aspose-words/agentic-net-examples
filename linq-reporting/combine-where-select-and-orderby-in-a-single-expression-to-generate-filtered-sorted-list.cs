using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ProductDto
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<ProductDto> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a Word template with LINQ Reporting tags
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Name]>> - $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template (demonstrates load rule)
        Document doc = new Document(templatePath);

        // Sample data source
        List<Product> products = new()
        {
            new Product { Name = "Apple", Price = 5m },
            new Product { Name = "Banana", Price = 12m },
            new Product { Name = "Cherry", Price = 15m },
            new Product { Name = "Date", Price = 8m },
            new Product { Name = "Elderberry", Price = 20m }
        };

        // Combine Where, Select, and OrderBy in a single LINQ expression
        List<ProductDto> filteredSorted = products
            .Where(p => p.Price > 10)                     // Filter
            .Select(p => new ProductDto                    // Project
            {
                Name = p.Name,
                Price = p.Price
            })
            .OrderBy(p => p.Name)                         // Sort
            .ToList();

        // Prepare the root model for the report
        ReportModel model = new()
        {
            Items = filteredSorted
        };

        // Build the report using Aspose.Words ReportingEngine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}
