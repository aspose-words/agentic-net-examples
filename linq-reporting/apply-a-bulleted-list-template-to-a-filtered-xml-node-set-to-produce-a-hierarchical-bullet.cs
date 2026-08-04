using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Catalog>
  <Category Name=""Fruits"" Type=""A"">
    <Product Name=""Apple"" />
    <Product Name=""Banana"" />
  </Category>
  <Category Name=""Vegetables"" Type=""B"">
    <Product Name=""Carrot"" />
    <Product Name=""Lettuce"" />
  </Category>
  <Category Name=""Beverages"" Type=""A"">
    <Product Name=""Coffee"" />
    <Product Name=""Tea"" />
  </Category>
</Catalog>";

        // Write XML to a temporary file (optional, just to keep the example self‑contained).
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "catalog.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML into strongly‑typed model classes.
        Catalog catalog = LoadCatalogFromXml(xmlPath);

        // Build the template document with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [cat in catalog.Category]>>");
        builder.Writeln("<<if [cat.Type == \"A\"]>>");
        builder.Writeln("• <<[cat.Name]>>");
        builder.Writeln("<<foreach [prod in cat.Product]>>");
        builder.Writeln("   • <<[prod.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Generate the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, catalog, "catalog");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }

    private static Catalog LoadCatalogFromXml(string xmlPath)
    {
        XDocument xdoc = XDocument.Load(xmlPath);
        Catalog catalog = new Catalog();

        foreach (XElement catElem in xdoc.Root.Elements("Category"))
        {
            Category category = new Category
            {
                Name = (string)catElem.Attribute("Name") ?? string.Empty,
                Type = (string)catElem.Attribute("Type") ?? string.Empty
            };

            foreach (XElement prodElem in catElem.Elements("Product"))
            {
                Product product = new Product
                {
                    Name = (string)prodElem.Attribute("Name") ?? string.Empty
                };
                category.Product.Add(product);
            }

            catalog.Category.Add(category);
        }

        return catalog;
    }
}

// Public data model classes.
public class Catalog
{
    public List<Category> Category { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public List<Product> Product { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
}
