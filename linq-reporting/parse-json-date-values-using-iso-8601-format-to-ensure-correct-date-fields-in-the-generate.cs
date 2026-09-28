using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Order
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create the template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate.ToString(\"yyyy-MM-dd\")]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Create sample JSON data with ISO 8601 dates.
        string jsonPath = Path.Combine(outputDir, "data.json");
        string jsonContent = @"{
  ""Orders"": [
    { ""CustomerName"": ""Alice"", ""OrderDate"": ""2023-08-15T14:30:00Z"" },
    { ""CustomerName"": ""Bob"",   ""OrderDate"": ""2023-09-01T09:15:00Z"" }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // Load and deserialize JSON data.
        string json = File.ReadAllText(jsonPath);
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json) ?? new ReportModel();

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);
    }
}
