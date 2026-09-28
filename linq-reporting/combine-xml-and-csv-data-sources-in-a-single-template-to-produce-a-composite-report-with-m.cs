using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data.
        string xmlPath = "products.xml";
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""utf-8""?>
<Products>
    <Product>
        <Name>Apple</Name>
        <Price>0.50</Price>
    </Product>
    <Product>
        <Name>Banana</Name>
        <Price>0.30</Price>
    </Product>
    <Product>
        <Name>Orange</Name>
        <Price>0.80</Price>
    </Product>
</Products>");

        // Create sample CSV data.
        string csvPath = "sales.csv";
        File.WriteAllText(csvPath, @"Date,Product,Quantity
2023-01-01,Apple,10
2023-01-02,Banana,5
2023-01-03,Orange,8");

        // Load XML into objects.
        List<Product> products = XDocument.Load(xmlPath)
            .Descendants("Product")
            .Select(p => new Product
            {
                Name = (string)p.Element("Name") ?? string.Empty,
                Price = decimal.TryParse((string)p.Element("Price"), out var pr) ? pr : 0m
            })
            .ToList();

        // Load CSV into objects.
        List<Sale> sales = new();
        string[] csvLines = File.ReadAllLines(csvPath);
        if (csvLines.Length > 1)
        {
            string[] headers = csvLines[0].Split(',');
            for (int i = 1; i < csvLines.Length; i++)
            {
                string[] fields = csvLines[i].Split(',');
                if (fields.Length != headers.Length) continue;

                sales.Add(new Sale
                {
                    Date = fields[0],
                    Product = fields[1],
                    Quantity = int.TryParse(fields[2], out var q) ? q : 0
                });
            }
        }

        // Create the template document.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Composite Report");
        builder.Writeln("-----------------");
        builder.Writeln();

        // XML section.
        builder.Writeln("Products (from XML):");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("- <<[p.Name]>> : $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // CSV section.
        builder.Writeln("Sales (from CSV):");
        builder.Writeln("<<foreach [s in Sales]>>");
        builder.Writeln("- <<[s.Date]>> : <<[s.Product]>> x <<[s.Quantity]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Prepare the model.
        ReportModel model = new()
        {
            Products = products,
            Sales = sales
        };

        // Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

public class Sale
{
    public string Date { get; set; } = string.Empty;
    public string Product { get; set; } = string.Empty;
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
    public List<Sale> Sales { get; set; } = new();
}
