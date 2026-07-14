using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Xml.Linq;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data files.
        string jsonPath = "people.json";
        string xmlPath = "products.xml";

        File.WriteAllText(jsonPath, @"{
  ""Persons"": [
    { ""Name"": ""Alice"", ""Age"": 30 },
    { ""Name"": ""Bob"", ""Age"": 25 },
    { ""Name"": ""Charlie"", ""Age"": 35 }
  ]
}");
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Products>
  <Product>
    <Name>Apple</Name>
    <Price>1.20</Price>
  </Product>
  <Product>
    <Name>Banana</Name>
    <Price>0.80</Price>
  </Product>
  <Product>
    <Name>Cherry</Name>
    <Price>2.50</Price>
  </Product>
</Products>");

        // Load data into the model.
        ReportModel model = new()
        {
            Persons = JsonConvert.DeserializeObject<RootJson>(File.ReadAllText(jsonPath))?.Persons ?? new(),
            Products = LoadProductsFromXml(xmlPath)
        };

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("=== LINQ Reporting Example ===");
        builder.Writeln();

        builder.Writeln("JSON Persons:");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("- <<[person.Name]>> (Age: <<[person.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        builder.Writeln("XML Products:");
        builder.Writeln("<<foreach [product in Products]>>");
        builder.Writeln("- <<[product.Name]>> : $<<[product.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        template.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    private static List<Product> LoadProductsFromXml(string xmlPath)
    {
        XDocument doc = XDocument.Load(xmlPath);
        List<Product> products = new();
        foreach (XElement prodElem in doc.Root?.Elements("Product") ?? new List<XElement>())
        {
            string name = prodElem.Element("Name")?.Value ?? "";
            string priceText = prodElem.Element("Price")?.Value ?? "0";
            decimal price = decimal.TryParse(priceText, out var p) ? p : 0m;
            products.Add(new Product { Name = name, Price = price });
        }
        return products;
    }

    private class RootJson
    {
        public List<Person> Persons { get; set; } = new();
    }
}
