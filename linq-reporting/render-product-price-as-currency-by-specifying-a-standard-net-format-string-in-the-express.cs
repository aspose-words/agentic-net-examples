using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

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
        // Prepare sample data
        var model = new ReportModel
        {
            Products = new()
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Create template document
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();

        builder.Writeln("Product Report");
        builder.Writeln();

        // Table with LINQ Reporting tags
        builder.Writeln("<<foreach [p in Products]>>");

        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row (template)
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price.ToString(\"C\")]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template
        builder.Document.Save(templatePath);

        // Load the template for report generation
        var doc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
