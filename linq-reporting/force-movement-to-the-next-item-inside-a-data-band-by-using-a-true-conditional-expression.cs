using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20 },
                new Product { Name = "Banana", Price = 0.80 },
                new Product { Name = "Cherry", Price = 2.50 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product list (price will be skipped by forced next):");
        builder.Writeln("<<foreach [p in Products]>>");

        // Output the product name.
        builder.Writeln("Name: <<[p.Name]>>");

        // Force movement to the next item using a true condition.
        // The <<next>> tag moves to the next record in the data band.
        builder.Writeln("<<if [true]>><<next>><</if>>");

        // This line will never be reached because of the forced <<next>> above.
        builder.Writeln("Price: <<[p.Price]>>");

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();

        // The root object name must match the name used in the template tags.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (must be public with public properties).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
