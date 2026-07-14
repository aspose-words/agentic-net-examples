using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public double Price { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        List<Product> products = new()
        {
            new Product { Name = "Apple", Price = 0.99 },
            new Product { Name = "Banana", Price = 0.59 },
            new Product { Name = "Cherry", Price = 2.49 }
        };

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Product List");
        builder.Writeln();

        // Begin a foreach loop over the collection named "products".
        builder.Writeln("<<foreach [p in products]>>");
        builder.Writeln("- <<[p.Name]>> : $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the overload that specifies the data source name.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, products, "products");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProductReport.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
