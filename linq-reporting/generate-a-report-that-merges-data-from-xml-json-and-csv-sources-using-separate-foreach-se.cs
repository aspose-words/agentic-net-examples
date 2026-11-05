using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure code page support for potential CSV encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data files.
        string jsonPath = "data.json";
        string xmlPath = "data.xml";
        string csvPath = "data.csv";

        File.WriteAllText(jsonPath, @"[
  { ""Name"": ""Alice"", ""Age"": 30 },
  { ""Name"": ""Bob"",   ""Age"": 25 }
]");

        File.WriteAllText(xmlPath, @"<Products>
  <Product><Name>Widget</Name><Price>19.99</Price></Product>
  <Product><Name>Gadget</Name><Price>29.99</Price></Product>
</Products>");

        File.WriteAllText(csvPath, @"Item,Quantity
Pen,10
Notebook,5
");

        // Load JSON data.
        var jsonCustomers = JsonConvert.DeserializeObject<List<Customer>>(File.ReadAllText(jsonPath)) ?? new();

        // Load XML data.
        var xmlDoc = XDocument.Load(xmlPath);
        var xmlProducts = xmlDoc.Root?
            .Elements("Product")
            .Select(p => new Product
            {
                Name = (string?)p.Element("Name") ?? "",
                Price = decimal.TryParse((string?)p.Element("Price"), out var v) ? v : 0m
            })
            .ToList() ?? new();

        // Load CSV data.
        var csvLines = File.ReadAllLines(csvPath);
        var csvSales = new List<Sale>();
        foreach (var line in csvLines.Skip(1)) // Skip header
        {
            var parts = line.Split(',');
            if (parts.Length >= 2)
            {
                csvSales.Add(new Sale
                {
                    Item = parts[0],
                    Quantity = int.TryParse(parts[1], out var q) ? q : 0
                });
            }
        }

        // Assemble the root model.
        var model = new ReportModel
        {
            JsonCustomers = jsonCustomers,
            XmlProducts = xmlProducts,
            CsvSales = csvSales
        };

        // Build the LINQ Reporting template.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("=== JSON Customers ===");
        builder.Writeln("<<foreach [c in JsonCustomers]>>");
        builder.Writeln("Name: <<[c.Name]>>   Age: <<[c.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        builder.Writeln("=== XML Products ===");
        builder.Writeln("<<foreach [p in XmlProducts]>>");
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        builder.Writeln("=== CSV Sales ===");
        builder.Writeln("<<foreach [s in CsvSales]>>");
        builder.Writeln("Item: <<[s.Item]>>   Quantity: <<[s.Quantity]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Customer> JsonCustomers { get; set; } = new();
    public List<Product> XmlProducts { get; set; } = new();
    public List<Sale> CsvSales { get; set; } = new();
}

public class Customer
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class Sale
{
    public string Item { get; set; } = "";
    public int Quantity { get; set; }
}
