using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a large XML data set.
        string xmlPath = Path.Combine(outputDir, "products.xml");
        CreateSampleXml(xmlPath, 200);

        // 2. Parse XML with progress reporting.
        ReportData data = new();
        Action<int, int> progressCallback = (processed, total) =>
        {
            Console.WriteLine($"Parsing XML: {processed}/{total} items processed.");
        };
        ParseXmlWithProgress(xmlPath, data, progressCallback);

        // 3. Create a LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // 4. Load the template and build the report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, data, "data");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);
        Console.WriteLine($"Report generated at: {reportPath}");
    }

    private static void CreateSampleXml(string path, int count)
    {
        XDocument doc = new(
            new XElement("Products",
                Enumerable.Range(1, count).Select(i =>
                    new XElement("Product",
                        new XElement("Id", i),
                        new XElement("Name", $"Product {i}"),
                        new XElement("Price", (i * 1.23).ToString("F2")),
                        new XElement("Description", $"Description for product {i}.")
                    ))
            )
        );
        doc.Save(path);
    }

    private static void ParseXmlWithProgress(string xmlPath, ReportData data, Action<int, int> progress)
    {
        XDocument doc = XDocument.Load(xmlPath);
        var productElements = doc.Root?.Elements("Product") ?? Enumerable.Empty<XElement>();
        int total = productElements.Count();
        int processed = 0;

        foreach (var elem in productElements)
        {
            Product p = new()
            {
                Id = (int)elem.Element("Id")!,
                Name = (string)elem.Element("Name")!,
                Price = decimal.Parse((string)elem.Element("Price")!),
                Description = (string)elem.Element("Description")!
            };
            data.Products.Add(p);
            processed++;
            progress?.Invoke(processed, total);
        }
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Product Report");
        builder.Writeln("Total Products: <<[data.Products.Count]>>");
        builder.Writeln("<<foreach [p in data.Products]>>");
        builder.Writeln("Id: <<[p.Id]>>, Name: <<[p.Name]>>, Price: $<<[p.Price]>>");
        builder.Writeln("Description: <<[p.Description]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

public class ReportData
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public string Description { get; set; } = string.Empty;
}
