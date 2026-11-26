using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath,
@"<Products>
    <Product Name='Laptop' Category='Electronics' />
    <Product Name='Desk' Category='Furniture' />
    <Product Name='Smartphone' Category='electronics' />
    <Product Name='Chair' Category='Furniture' />
    <Product Name='Headphones' Category='ELECTRONICS' />
</Products>");

        // Load XML and parse into objects.
        XDocument xdoc = XDocument.Load(xmlPath);
        List<Product> allProducts = xdoc.Root!
            .Elements("Product")
            .Select(e => new Product
            {
                Name = (string?)e.Attribute("Name") ?? "",
                Category = (string?)e.Attribute("Category") ?? ""
            })
            .ToList();

        // Filter with case‑insensitive comparison using Where.
        List<Product> filteredProducts = allProducts
            .Where(p => string.Equals(p.Category, "electronics", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Prepare the model for the report.
        ReportModel model = new()
        {
            Products = filteredProducts
        };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Products in the 'Electronics' category:");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("- <<[p.Name]>> (<<[p.Category]>>)");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
