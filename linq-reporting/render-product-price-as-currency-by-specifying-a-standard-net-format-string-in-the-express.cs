using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

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
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new()
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Create a template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Price Report");
        builder.Writeln(string.Empty);
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Price: <<[p.Price.ToString(\"C\")]>>");
        builder.Writeln(string.Empty);
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var template = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        template.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
