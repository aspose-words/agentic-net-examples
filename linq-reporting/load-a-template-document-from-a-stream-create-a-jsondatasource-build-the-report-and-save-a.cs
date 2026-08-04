using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data representing an order.
        string json = @"
{
    ""CustomerName"": ""John Doe"",
    ""Items"": [
        { ""Name"": ""Apple"",  ""Quantity"": 3 },
        { ""Name"": ""Banana"", ""Quantity"": 5 },
        { ""Name"": ""Orange"", ""Quantity"": 2 }
    ]
}";

        // Write JSON to a temporary file because JsonDataSource expects a file path.
        string jsonPath = Path.Combine(Path.GetTempPath(), "order.json");
        File.WriteAllText(jsonPath, json);

        // Create a template document in memory.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using MemoryStream templateStream = new MemoryStream();
        templateDoc.Save(templateStream, SaveFormat.Docx);
        templateStream.Position = 0; // Reset stream for reading.

        // Load the template document from the stream.
        Document doc = new Document(templateStream);

        // Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the JSON data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "order");

        // Save the generated report as RTF.
        doc.Save("Report.rtf", SaveFormat.Rtf);
    }
}
