using System;
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

        // Prepare sample JSON data with ISO 8601 date.
        string jsonPath = "order.json";
        string jsonContent = @"{
            ""OrderDate"": ""2023-08-15T14:30:00Z"",
            ""CustomerName"": ""John Doe""
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // Deserialize JSON into a strongly‑typed model.
        Order order = JsonConvert.DeserializeObject<Order>(jsonContent) ?? new Order();

        // Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Date: <<[order.OrderDate.ToString(\"yyyy-MM-dd HH:mm:ss\")]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Public data model matching the JSON structure.
public class Order
{
    public DateTime OrderDate { get; set; } = DateTime.MinValue;
    public string CustomerName { get; set; } = string.Empty;
}
