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
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "orders.json");
        var sampleData = new Root
        {
            Orders = new List<Order>
            {
                new Order
                {
                    CustomerName = "Acme Corp",
                    OrderDate = new DateTime(2023, 5, 21),
                    Total = 1234.56m
                },
                new Order
                {
                    CustomerName = "Globex Inc",
                    OrderDate = new DateTime(2023, 6, 3),
                    Total = 7890.12m
                }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON into model.
        var root = JsonConvert.DeserializeObject<Root>(File.ReadAllText(jsonPath))!;

        // Create a DOCX template with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Begin foreach over orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate.ToString(\"d\")]>>");
        builder.Writeln("Total: <<[order.Total.ToString(\"C\")]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report using the root object.
        engine.BuildReport(reportDoc, root, "root");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Data model classes.
public class Root
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public DateTime OrderDate { get; set; }
    public decimal Total { get; set; }
}
