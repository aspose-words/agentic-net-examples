using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }

    public Product(string name, decimal price)
    {
        Name = name;
        Price = price;
    }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with a table that is wrapped by a
        //    LINQ Reporting foreach block.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Product Report");

        // Open the foreach block – it will repeat the whole table for each product.
        builder.Writeln("<<foreach [product in Products]>>");

        // Start the table that will be repeated.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row – the tags will be replaced with each product's values.
        builder.InsertCell();
        builder.Writeln("<<[product.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[product.Price]>>");
        builder.EndRow();

        // Finish the table and close the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template (simulating a real‑world scenario) and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data.
        ReportModel model = new ReportModel();
        model.Products.Add(new Product("Apple", 0.99m));
        model.Products.Add(new Product("Banana", 0.59m));
        model.Products.Add(new Product("Cherry", 2.49m));

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        bool success = engine.BuildReport(doc, model, "model");

        // Optional: check the success flag when InlineErrorMessages option is used.
        Console.WriteLine($"Report build successful: {success}");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
