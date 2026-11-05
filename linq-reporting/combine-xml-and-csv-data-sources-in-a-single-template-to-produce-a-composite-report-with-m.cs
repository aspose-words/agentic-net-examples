using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Paths for sample files.
        const string xmlPath = "Customers.xml";
        const string csvPath = "Products.csv";
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // Create sample XML file.
        File.WriteAllText(xmlPath,
@"<Customers>
    <Customer>
        <Name>John Doe</Name>
        <City>New York</City>
    </Customer>
    <Customer>
        <Name>Jane Smith</Name>
        <City>London</City>
    </Customer>
</Customers>");

        // Create sample CSV file (header + rows).
        File.WriteAllText(csvPath,
@"Name,Price
Laptop,1200
Smartphone,800");

        // Load customers from XML.
        var customers = XDocument.Load(xmlPath)
            .Root?
            .Elements("Customer")
            .Select(x => new Customer
            {
                Name = (string?)x.Element("Name") ?? string.Empty,
                City = (string?)x.Element("City") ?? string.Empty
            })
            .ToList() ?? new List<Customer>();

        // Load products from CSV.
        var products = File.ReadAllLines(csvPath)
            .Skip(1) // Skip header.
            .Select(line => line.Split(','))
            .Where(parts => parts.Length >= 2)
            .Select(parts => new Product
            {
                Name = parts[0],
                Price = parts[1]
            })
            .ToList();

        // Assemble the root model.
        var model = new ReportModel
        {
            Customers = customers,
            Products = products
        };

        // -------------------- Create template document --------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Customers section.
        builder.Writeln("Customer List");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name: <<[c.Name]>>, City: <<[c.City]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Products section – table inside a foreach loop.
        builder.Writeln("Product List");
        builder.Writeln("<<foreach [p in Products]>>");

        Table table = builder.StartTable();
        // Header row.
        builder.InsertCell();
        builder.Writeln("Product Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row (repeated for each product).
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");
        // ------------------------------------------------------------------

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save(outputPath);
    }
}

// Root model containing both data sources.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
    public List<Product> Products { get; set; } = new();
}

// Simple customer POCO.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string City { get; set; } = string.Empty;
}

// Simple product POCO.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public string Price { get; set; } = string.Empty;
}
