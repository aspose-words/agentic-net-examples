using System;
using System.Collections.Generic;
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
        // Step 1: Create the template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [product in Products]>>");
        builder.Writeln("Name: <<[product.Name]>>");
        builder.Writeln("Price: <<[product.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for report generation.
        var doc = new Document(templatePath);

        // Step 3: Prepare sample data.
        var model = new ReportModel();
        model.Products.Add(new Product { Name = "Apple", Price = 0.99m });
        model.Products.Add(new Product { Name = "Banana", Price = 0.59m });
        model.Products.Add(new Product { Name = "Cherry", Price = 2.49m });

        // Step 4: Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
