using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Ensure the product name is not empty.
        builder.Writeln("<<if [p.Name.Length > 0]>>");

        // Apply a custom color to the first character only.
        // The first character is wrapped in a textColor tag.
        // The remainder of the name is inserted without any color tag.
        builder.Writeln(
            "<<textColor [\"Red\"]>><<[p.Name[0]]>><</textColor>><<[p.Name.Substring(1)]>>");

        // Close the if block.
        builder.Writeln("<</if>>");

        // Append the price after a hyphen.
        builder.Writeln(" - <<[p.Price]>>");
        builder.Writeln();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new()
        {
            Products = new()
            {
                new Product { Name = "Apple", Price = 1.20m },
                new Product { Name = "Banana", Price = 0.80m },
                new Product { Name = "Cherry", Price = 2.50m }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Simple product class.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
