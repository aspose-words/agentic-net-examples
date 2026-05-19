using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");
        // Write product name.
        builder.Writeln("Product: <<[p.Name]>>");
        // Write quantity with conditional display.
        builder.Writeln(
            "Quantity: " +
            "<<if [p.Quantity > 0]>>" +
            "<<[p.Quantity]>>" +
            "<</if>>" +
            "<<if [p.Quantity == 0]>>Out of stock<</if>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Quantity = 5 },
                new Product { Name = "Banana", Quantity = 0 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// Root wrapper class referenced in the template as "model".
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Simple data model for each product.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
