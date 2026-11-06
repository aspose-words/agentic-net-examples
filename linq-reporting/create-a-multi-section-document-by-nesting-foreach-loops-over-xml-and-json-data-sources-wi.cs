using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Category
{
    public string Name { get; set; } = "";
}

public class Product
{
    public string CategoryName { get; set; } = "";
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create sample XML data (categories).
        // -----------------------------------------------------------------
        const string xmlPath = "categories.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Categories>
    <Category><Name>Electronics</Name></Category>
    <Category><Name>Books</Name></Category>
    <Category><Name>Clothing</Name></Category>
</Categories>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Create sample JSON data (products).
        // -----------------------------------------------------------------
        const string jsonPath = "products.json";
        var jsonContent = @"[
    { ""CategoryName"": ""Electronics"", ""Name"": ""Smartphone"", ""Price"": 699.99 },
    { ""CategoryName"": ""Electronics"", ""Name"": ""Laptop"", ""Price"": 1199.50 },
    { ""CategoryName"": ""Books"", ""Name"": ""C# Programming"", ""Price"": 39.99 },
    { ""CategoryName"": ""Clothing"", ""Name"": ""T-Shirt"", ""Price"": 19.95 },
    { ""CategoryName"": ""Clothing"", ""Name"": ""Jeans"", ""Price"": 49.99 }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Load XML into Category objects.
        // -----------------------------------------------------------------
        var xDoc = XDocument.Load(xmlPath);
        var categories = new List<Category>();
        foreach (var elem in xDoc.Root?.Elements("Category") ?? [])
        {
            var nameElem = elem.Element("Name");
            if (nameElem != null)
            {
                categories.Add(new Category { Name = nameElem.Value });
            }
        }

        // -----------------------------------------------------------------
        // Load JSON into Product objects.
        // -----------------------------------------------------------------
        var products = JsonConvert.DeserializeObject<List<Product>>(jsonContent) ?? new List<Product>();

        // -----------------------------------------------------------------
        // Assemble the root model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Categories = categories,
            Products = products
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title section.
        builder.Writeln("Multi‑Section Report");
        builder.Writeln();

        // Start a new section for the data.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Outer foreach over categories (using the root model name).
        builder.Writeln("<<foreach [cat in model.Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln();

        // Inner foreach over products, filtered by the current category.
        builder.Writeln("<<foreach [prod in model.Products]>>");
        builder.Writeln("<<if [prod.CategoryName == cat.Name]>>");
        builder.Writeln("- Product: <<[prod.Name]>> (Price: $<<[prod.Price]>>)");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model and the root name "model".
        engine.BuildReport(templateDoc, model, "model");

        // Save the final document.
        const string outputPath = "output.docx";
        templateDoc.Save(outputPath);
    }
}
