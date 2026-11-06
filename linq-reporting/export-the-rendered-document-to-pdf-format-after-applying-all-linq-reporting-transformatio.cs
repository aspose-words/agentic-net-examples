using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data model
        ReportModel model = new()
        {
            Title = "Product Report",
            Products = new()
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            }
        };

        // Build the template document programmatically
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Title
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Foreach over Products
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("- <<[p.Name]>>: $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Export the rendered document to PDF
        string pdfPath = "report.pdf";
        reportDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}

// Data model classes
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
