using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

#nullable enable

public class Customer
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Product
{
    public string Name { get; set; } = "";
    public double Price { get; set; }
}

public class Order
{
    public int OrderId { get; set; }
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
    public List<Product> Products { get; set; } = new();
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a folder for sample data.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        // ---------- Create sample XML ----------
        string xmlPath = Path.Combine(dataFolder, "customers.xml");
        string xmlContent = @"<Customers>
  <Customer>
    <Name>John Doe</Name>
    <Age>30</Age>
  </Customer>
  <Customer>
    <Name>Jane Smith</Name>
    <Age>25</Age>
  </Customer>
</Customers>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // ---------- Create sample JSON ----------
        string jsonPath = Path.Combine(dataFolder, "products.json");
        string jsonContent = @"[
  { ""Name"": ""Laptop"", ""Price"": 1200.50 },
  { ""Name"": ""Mouse"", ""Price"": 25.99 }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // ---------- Create sample CSV ----------
        string csvPath = Path.Combine(dataFolder, "orders.csv");
        string csvContent = @"OrderId,Product,Quantity
1,Laptop,2
2,Mouse,5";
        File.WriteAllText(csvPath, csvContent, Encoding.UTF8);

        // ---------- Load data into model ----------
        ReportModel model = new ReportModel();

        // Load XML
        XDocument xDoc = XDocument.Load(xmlPath);
        model.Customers = xDoc.Root!
            .Elements("Customer")
            .Select(x => new Customer
            {
                Name = (string?)x.Element("Name") ?? "",
                Age = (int?)x.Element("Age") ?? 0
            })
            .ToList();

        // Load JSON
        string jsonString = File.ReadAllText(jsonPath, Encoding.UTF8);
        model.Products = JsonConvert.DeserializeObject<List<Product>>(jsonString) ?? new List<Product>();

        // Load CSV
        string[] csvLines = File.ReadAllLines(csvPath, Encoding.UTF8);
        if (csvLines.Length > 1)
        {
            model.Orders = csvLines
                .Skip(1) // skip header
                .Select(line => line.Split(','))
                .Where(parts => parts.Length == 3)
                .Select(parts => new Order
                {
                    OrderId = int.TryParse(parts[0], out int id) ? id : 0,
                    Product = parts[1],
                    Quantity = int.TryParse(parts[2], out int qty) ? qty : 0
                })
                .ToList();
        }

        // ---------- Create template document ----------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Customers section
        builder.Writeln("Customers:");
        builder.Writeln("<<foreach [c in Customers]>>");
        Table custTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Age");
        builder.EndRow();

        builder.InsertCell(); builder.Writeln("<<[c.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[c.Age]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Products section
        builder.Writeln("\nProducts:");
        builder.Writeln("<<foreach [p in Products]>>");
        Table prodTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Price");
        builder.EndRow();

        builder.InsertCell(); builder.Writeln("<<[p.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Price]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Orders section
        builder.Writeln("\nOrders:");
        builder.Writeln("<<foreach [o in Orders]>>");
        Table orderTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Order ID");
        builder.InsertCell(); builder.Writeln("Product");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.EndRow();

        builder.InsertCell(); builder.Writeln("<<[o.OrderId]>>");
        builder.InsertCell(); builder.Writeln("<<[o.Product]>>");
        builder.InsertCell(); builder.Writeln("<<[o.Quantity]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "report_template.docx");
        template.Save(templatePath);

        // ---------- Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(template, model, "model");

        // Save the final report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report_output.docx");
        template.Save(outputPath);
    }
}
