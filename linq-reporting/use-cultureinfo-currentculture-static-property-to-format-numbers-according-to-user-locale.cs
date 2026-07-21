using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Product Report");
        builder.Writeln();
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        builder.Writeln("Price: <<[p.Price.ToString(\"C\")]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new() { Name = "Apple", Price = 1.23m },
                new() { Name = "Banana", Price = 0.99m },
                new() { Name = "Cherry", Price = 2.50m }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
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
