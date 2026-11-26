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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new ReportData
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, Customer = "Alice", Total = 1234.56m },
                new Order { Id = 2, Customer = "Bob",   Total = 78.90m },
                new Order { Id = 3, Customer = "Carol", Total = 2500.00m }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load data from JSON.
        var jsonContent = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<ReportData>(jsonContent) ?? new ReportData();

        // Create a LINQ Reporting template.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("");
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        // Currency formatting using ToString with culture-invariant format.
        builder.Writeln("Total: <<[order.Total.ToString(\"C\")]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        templateDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

public class ReportData
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int Id { get; set; }
    public string Customer { get; set; } = "";
    public decimal Total { get; set; }
}
