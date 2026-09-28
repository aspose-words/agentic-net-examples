using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public double Quantity { get; set; }
    public double Price { get; set; }
    public double LineTotal { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Quantity: <<[p.Quantity]>>");
        builder.Writeln("Total: <<[p.LineTotal]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Products.Add(new Product { Name = "Apple", Quantity = 3.7, Price = 2.5 });
        model.Products.Add(new Product { Name = "Banana", Quantity = 5.2, Price = 1.2 });
        model.Products.Add(new Product { Name = "Cherry", Quantity = 2.9, Price = 3.0 });

        // Compute LineTotal using explicit cast (int) on Quantity.
        foreach (var p in model.Products)
        {
            p.LineTotal = (int)p.Quantity * p.Price;
        }

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
