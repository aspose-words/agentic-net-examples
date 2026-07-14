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
        // Sample XML data.
        string xmlContent = @"<Products>
    <Product Name='Apple' Category='Fruit' />
    <Product Name='Banana' Category='fruit' />
    <Product Name='Carrot' Category='Vegetable' />
    <Product Name='Broccoli' Category='vegetable' />
</Products>";

        // Load XML.
        XDocument xdoc = XDocument.Parse(xmlContent);

        // Filter nodes where the Category attribute equals "fruit" (case‑insensitive).
        List<Product> filtered = xdoc.Root?
            .Elements("Product")
            .Where(e => string.Equals((string)e.Attribute("Category"), "fruit", StringComparison.OrdinalIgnoreCase))
            .Select(e => new Product
            {
                Name = (string)e.Attribute("Name") ?? "",
                Category = (string)e.Attribute("Category") ?? ""
            })
            .ToList() ?? new List<Product>();

        // Prepare the model for the report.
        ReportModel model = new()
        {
            Products = filtered
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Products with Category = \"fruit\" (case‑insensitive):");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>, Category: <<[p.Category]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
