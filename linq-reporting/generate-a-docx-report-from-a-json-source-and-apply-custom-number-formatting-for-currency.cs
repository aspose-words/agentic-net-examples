using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        string jsonPath = "data.json";
        var sampleData = new ReportData
        {
            Orders = new List<Order>
            {
                new Order { CustomerName = "Alice Johnson", Total = 1234.56m },
                new Order { CustomerName = "Bob Smith", Total = 7890.12m },
                new Order { CustomerName = "Carol Davis", Total = 345.67m }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load data from JSON
        var jsonContent = File.ReadAllText(jsonPath);
        var data = JsonConvert.DeserializeObject<ReportData>(jsonContent) ?? new ReportData();

        // Create template document
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Generated on: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        builder.Writeln();

        // Begin foreach loop over Orders
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.TotalFormatted]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // Load the template (optional, can reuse the same doc)
        var templateDoc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, data, "data");

        // Save the final report
        string reportPath = "report.docx";
        templateDoc.Save(reportPath);
    }
}

public class ReportData
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public decimal Total { get; set; }

    public string TotalFormatted => Total.ToString("C", CultureInfo.CurrentCulture);
}
