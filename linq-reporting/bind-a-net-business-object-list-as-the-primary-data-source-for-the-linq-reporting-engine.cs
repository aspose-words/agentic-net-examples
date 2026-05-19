using System;
using System.Collections.Generic;
using System.IO;
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
    public static void Main(string[] args)
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // Create a Word template with LINQ Reporting tags.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Price: <<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the business object list as the data source.
        // -------------------------------------------------
        var model = new ReportModel
        {
            Products = new()
            {
                new Product { Name = "Apple",  Price = 1.20m },
                new Product { Name = "Banana", Price = 0.80m },
                new Product { Name = "Orange", Price = 1.50m }
            }
        };

        // -------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);

        // Inform the user where the report was saved.
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
