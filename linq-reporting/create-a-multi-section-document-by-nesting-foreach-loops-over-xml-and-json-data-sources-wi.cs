using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample XML file.
        const string xmlPath = "customers.xml";
        File.WriteAllText(
            xmlPath,
            @"<Customers>
  <Customer>
    <Id>C001</Id>
    <Name>Alice</Name>
  </Customer>
  <Customer>
    <Id>C002</Id>
    <Name>Bob</Name>
  </Customer>
</Customers>",
            Encoding.UTF8);

        // Create sample JSON file.
        const string jsonPath = "orders.json";
        File.WriteAllText(
            jsonPath,
            @"[
  { ""CustomerId"": ""C001"", ""Product"": ""Widget"", ""Quantity"": 5 },
  { ""CustomerId"": ""C001"", ""Product"": ""Gadget"", ""Quantity"": 2 },
  { ""CustomerId"": ""C002"", ""Product"": ""Thingamajig"", ""Quantity"": 1 }
]",
            Encoding.UTF8);

        // Load XML data into model.
        var xmlDoc = XDocument.Load(xmlPath);
        var customers = new List<Customer>();
        foreach (var elem in xmlDoc.Root?.Elements("Customer") ?? new List<XElement>())
        {
            customers.Add(new Customer
            {
                Id = elem.Element("Id")?.Value ?? string.Empty,
                Name = elem.Element("Name")?.Value ?? string.Empty
            });
        }

        // Load JSON data into model.
        var orders = JsonConvert.DeserializeObject<List<Order>>(File.ReadAllText(jsonPath)) ?? new List<Order>();

        // Prepare the root model.
        var model = new ReportModel
        {
            Customers = customers,
            Orders = orders
        };

        // Create the template document programmatically.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Customers and Their Orders ===");
        builder.Writeln("<<foreach [cust in Customers]>>");
        builder.Writeln("Customer: <<[cust.Name]>> (ID: <<[cust.Id]>>)");
        builder.Writeln("Orders:");
        // Use the root model name to access Orders collection.
        builder.Writeln("<<foreach [ord in model.Orders]>>");
        builder.Writeln("<<if [ord.CustomerId == cust.Id]>>- <<[ord.Product]>> x <<[ord.Quantity]>> <</if>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        const string outputPath = "report.docx";
        engine.BuildReport(doc, model, "model");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
    public List<Order> Orders { get; set; } = new();
}

public class Customer
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}

public class Order
{
    public string CustomerId { get; set; } = string.Empty;
    public string Product { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
