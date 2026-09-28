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
    public static void Main()
    {
        // Prepare template document.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // Render price as currency using standard .NET format string.
        builder.Writeln("Price: <<[p.Price.ToString(\"C\")]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Create sample data.
        ReportModel model = new()
        {
            Products = new()
            {
                new Product { Name = "Apple", Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Coffee", Price = 4.75m }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
