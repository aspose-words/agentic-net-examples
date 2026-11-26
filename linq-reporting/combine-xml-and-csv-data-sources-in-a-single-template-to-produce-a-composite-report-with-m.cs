using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data files.
        string xmlPath = "customers.xml";
        string csvPath = "orders.csv";
        CreateSampleXml(xmlPath);
        CreateSampleCsv(csvPath);

        // Load data from XML.
        List<Customer> customers = LoadCustomersFromXml(xmlPath);

        // Load data from CSV.
        List<OrderCsvRow> orders = LoadOrdersFromCsv(csvPath);

        // Combine data into a single model.
        ReportModel model = new ReportModel
        {
            Customers = customers,
            Orders = orders
        };

        // Create the template document.
        string templatePath = "template.docx";
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine("output", "CompositeReport.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {outputPath}");
    }

    private static void CreateSampleXml(string path)
    {
        var customers = new CustomersRoot
        {
            CustomerList = new List<Customer>
            {
                new Customer { Name = "John Doe", Email = "john.doe@example.com" },
                new Customer { Name = "Jane Smith", Email = "jane.smith@example.com" }
            }
        };

        XmlSerializer serializer = new XmlSerializer(typeof(CustomersRoot));
        using (var writer = new StreamWriter(path, false, Encoding.UTF8))
        {
            serializer.Serialize(writer, customers);
        }
    }

    private static void CreateSampleCsv(string path)
    {
        var lines = new List<string>
        {
            "OrderId,Product,Quantity",
            "1001,Widget A,5",
            "1002,Widget B,3",
            "1003,Widget C,7"
        };
        File.WriteAllLines(path, lines, Encoding.UTF8);
    }

    private static List<Customer> LoadCustomersFromXml(string path)
    {
        XmlSerializer serializer = new XmlSerializer(typeof(CustomersRoot));
        using (var stream = File.OpenRead(path))
        {
            var root = (CustomersRoot)serializer.Deserialize(stream)!;
            return root.CustomerList ?? new List<Customer>();
        }
    }

    private static List<OrderCsvRow> LoadOrdersFromCsv(string path)
    {
        var orders = new List<OrderCsvRow>();
        var lines = File.ReadAllLines(path, Encoding.UTF8);
        if (lines.Length < 2)
            return orders; // No data.

        // Assume first line is header.
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length != 3)
                continue;

            orders.Add(new OrderCsvRow
            {
                OrderId = parts[0].Trim(),
                Product = parts[1].Trim(),
                Quantity = int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty) ? qty : 0
            });
        }

        return orders;
    }

    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Composite Report");
        builder.Writeln();

        // Customers section.
        builder.Writeln("Customers:");
        builder.Writeln("<<foreach [cust in Customers]>>");
        builder.Writeln("- <<[cust.Name]>> (<<[cust.Email]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Orders section.
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [ord in Orders]>>");
        builder.Writeln("- Order ID: <<[ord.OrderId]>>, Product: <<[ord.Product]>>, Qty: <<[ord.Quantity]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

// Data model classes.

public class CustomersRoot
{
    [XmlArray("Customers")]
    [XmlArrayItem("Customer")]
    public List<Customer> CustomerList { get; set; } = new();
}

public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}

public class OrderCsvRow
{
    public string OrderId { get; set; } = string.Empty;
    public string Product { get; set; } = string.Empty;
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
    public List<OrderCsvRow> Orders { get; set; } = new();
}
