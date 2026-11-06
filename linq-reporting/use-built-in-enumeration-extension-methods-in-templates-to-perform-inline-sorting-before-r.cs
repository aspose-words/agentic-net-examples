using System;
using System.Collections.Generic;
using System.Linq;
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
        // Sample data (unsorted).
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m },
                new Product { Name = "Banana", Price = 0.80m },
                new Product { Name = "Cherry", Price = 2.50m },
                new Product { Name = "Date",   Price = 3.00m }
            }
        };

        // Create the LINQ Reporting template.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product List (sorted by Price ascending):");

        // Begin foreach loop with inline sorting.
        builder.Writeln("<<foreach [p in Products.OrderBy(p => p.Price)]>>");

        // Table that will be repeated for each product.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");
        builder.EndRow();

        // End of table for this iteration.
        builder.EndTable();

        // Close foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
