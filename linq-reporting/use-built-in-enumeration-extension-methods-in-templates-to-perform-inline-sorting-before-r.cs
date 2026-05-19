using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Products sorted by price (ascending):");
        // Inline sorting using the OrderBy LINQ extension method.
        builder.Writeln("<<foreach [p in model.Products.OrderBy(p => p.Price)]>>");
        builder.Writeln("<<[p.Name]>> - $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        ReportModel model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m },
                new Product { Name = "Banana", Price = 0.80m },
                new Product { Name = "Cherry", Price = 2.50m },
                new Product { Name = "Date",   Price = 3.00m }
            }
        };

        // -------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(reportPath);
    }
}

// -----------------------------------------------------------------
// Data model classes (public, non‑nullable properties initialized).
// -----------------------------------------------------------------
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
