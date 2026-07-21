#nullable enable
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

public class Program
{
    public static void Main()
    {
        // Enable code page provider for CSV encoding support
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare sample data files (XML, JSON, CSV)
        // -----------------------------------------------------------------
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        string xmlPath = Path.Combine(dataFolder, "persons.xml");
        File.WriteAllText(xmlPath,
@"<Persons>
    <Person><Name>John Doe</Name><Age>30</Age></Person>
    <Person><Name>Jane Smith</Name><Age>25</Age></Person>
</Persons>");

        string jsonPath = Path.Combine(dataFolder, "products.json");
        File.WriteAllText(jsonPath,
@"[
    { ""Name"": ""Apple"", ""Price"": 1.20 },
    { ""Name"": ""Banana"", ""Price"": 0.80 },
    { ""Name"": ""Cherry"", ""Price"": 2.50 }
]");

        string csvPath = Path.Combine(dataFolder, "orders.csv");
        File.WriteAllText(csvPath,
@"Id,Description,Amount
1,Item A,10
2,Item B,20
3,Item C,15");

        // -----------------------------------------------------------------
        // 2. Load data into strongly‑typed model
        // -----------------------------------------------------------------
        ReportData model = new()
        {
            Persons = LoadPersons(xmlPath),
            Products = LoadProducts(jsonPath),
            Orders = LoadOrders(csvPath)
        };

        // -----------------------------------------------------------------
        // 3. Create a Word template with LINQ Reporting tags
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate.docx");
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Title
        builder.Writeln("=== LINQ Reporting Example ===");
        builder.Writeln();

        // Persons (XML source) – foreach section
        builder.Writeln("Persons:");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln(" - <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Products (JSON source) – foreach section
        builder.Writeln("Products:");
        builder.Writeln("<<foreach [pr in Products]>>");
        builder.Writeln(" * <<[pr.Name]>> – $<<[pr.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Orders (CSV source) – foreach section inside a table
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in Orders]>>");
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[o.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[o.Description]>>");
        builder.InsertCell();
        builder.Writeln("<<[o.Amount]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Build the report
        // -----------------------------------------------------------------
        Document report = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportOutput.docx");
        report.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }

    // -----------------------------------------------------------------
    // Helper methods to load data from files
    // -----------------------------------------------------------------
    private static List<Person> LoadPersons(string path)
    {
        XDocument xdoc = XDocument.Load(path);
        return xdoc.Root?
            .Elements("Person")
            .Select(x => new Person
            {
                Name = (string?)x.Element("Name") ?? string.Empty,
                Age = (int?)x.Element("Age") ?? 0
            })
            .ToList() ?? new List<Person>();
    }

    private static List<Product> LoadProducts(string path)
    {
        string json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<List<Product>>(json) ?? new List<Product>();
    }

    private static List<Order> LoadOrders(string path)
    {
        var lines = File.ReadAllLines(path);
        var orders = new List<Order>();
        foreach (var line in lines.Skip(1)) // Skip header
        {
            var parts = line.Split(',');
            if (parts.Length != 3) continue;
            orders.Add(new Order
            {
                Id = int.TryParse(parts[0], out var id) ? id : 0,
                Description = parts[1],
                Amount = decimal.TryParse(parts[2], out var amt) ? amt : 0m
            });
        }
        return orders;
    }
}

// -----------------------------------------------------------------
// Data model classes
// -----------------------------------------------------------------
public class ReportData
{
    public List<Person> Persons { get; set; } = new();
    public List<Product> Products { get; set; } = new();
    public List<Order> Orders { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}

public class Order
{
    public int Id { get; set; }
    public string Description { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}
