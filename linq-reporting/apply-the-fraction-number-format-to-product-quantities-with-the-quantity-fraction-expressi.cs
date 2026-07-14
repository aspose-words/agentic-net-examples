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
                new Product { Name = "Apples",  Quantity = 1.5 },
                new Product { Name = "Oranges", Quantity = 2.25 },
                new Product { Name = "Bananas", Quantity = 0.75 }
            }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Start a foreach block to iterate over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        // Output the quantity. (If a fraction format is required, it can be applied via a custom helper method.)
        builder.Writeln("Quantity: <<[p.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, just to illustrate the process).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// Wrapper class that serves as the root data source for the report.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Simple product class with a name and a quantity.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Quantity { get; set; }
}
