using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

public class ReportModel
{
    // Initialize the collection.
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document with LINQ tags.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Header.
        builder.Writeln("Products:");

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [product in model.Products]>>");

        // Inside the loop output the product name and price formatted as currency.
        // The price is rendered using the standard .NET format string "C".
        builder.Writeln("- <<[product.Name]>>: <<[product.Price.ToString(\"C\")]>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare sample data.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new() { Name = "Apple",  Price = 1.23m },
                new() { Name = "Banana", Price = 0.99m },
                new() { Name = "Cherry", Price = 2.50m }
            }
        };

        // -------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine
        {
            // No special options are needed for this simple example.
            Options = ReportBuildOptions.None
        };

        // The root object name in the template is "model".
        bool success = engine.BuildReport(doc, model, "model");

        // Optionally, you could check the success flag if InlineErrorMessages were enabled.
        // For this example we simply proceed to save the document.

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(reportPath);
    }
}
