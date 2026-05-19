using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Color = "Red" },
                new Product { Name = "Banana", Color = "GoldenRod" },
                new Product { Name = "Grapes", Color = "#800080" } // purple
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product List:");
        builder.Writeln("<<foreach [p in Products]>>");
        // Dynamic text color applied via the textColor tag.
        builder.Writeln("<<textColor [p.Color]>><<[p.Name]>> <</textColor>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back (simulating a separate load step).
        var loadedDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the final document as HTML, preserving the color formatting.
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.html");
        var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        loadedDoc.Save(htmlPath, htmlOptions);
    }
}

// Wrapper class used as the root data source for the report.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Simple data model representing a product with a name and a color.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public string Color { get; set; } = string.Empty;
}
