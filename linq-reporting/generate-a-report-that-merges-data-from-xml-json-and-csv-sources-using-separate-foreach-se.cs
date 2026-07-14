using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Prepare sample data files.
        // -----------------------------------------------------------------
        string xmlPath = "people.xml";
        string jsonPath = "products.json";
        string csvPath = "orders.csv";

        // XML file with a list of persons.
        File.WriteAllText(xmlPath,
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<People>
    <Person>
        <Name>John Doe</Name>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
    </Person>
</People>");

        // JSON file with a list of products.
        var products = new List<Product>
        {
            new Product { Name = "Laptop", Price = 1299.99 },
            new Product { Name = "Smartphone", Price = 799.5 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(products, Formatting.Indented));

        // CSV file with a header row (Id,Description) and two orders.
        File.WriteAllText(csvPath, "Id,Description\r\n1,First order\r\n2,Second order");

        // -----------------------------------------------------------------
        // Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("=== XML Persons ===");
        builder.Writeln("<<foreach [p in xml]>>");
        builder.Writeln("- <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln("\n=== JSON Products ===");
        builder.Writeln("<<foreach [pr in json]>>");
        builder.Writeln("- <<[pr.Name]>> : $<<[pr.Price]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln("\n=== CSV Orders ===");
        builder.Writeln("<<foreach [o in csv]>>");
        builder.Writeln("- Order <<[o.Id]>>: <<[o.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for clarity).
        template.Save("template.docx");

        // -----------------------------------------------------------------
        // Load the template document (could also reuse the same instance).
        // -----------------------------------------------------------------
        Document doc = new Document("template.docx");

        // Create data source objects.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // CSV data source – enable header parsing so column names are recognized.
        var csvLoadOptions = new CsvDataLoadOptions { HasHeaders = true };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvLoadOptions);

        // Build the report using multiple data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc,
            new object[] { xmlData, jsonData, csvData },
            new string[] { "xml", "json", "csv" });

        // Save the final report.
        doc.Save("Report.docx");
    }

    // Simple POCO for JSON serialization.
    public class Product
    {
        public string Name { get; set; } = "";
        public double Price { get; set; }
    }
}
