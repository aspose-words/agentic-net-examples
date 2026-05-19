using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Sample JSON data.
        string json = @"{
            ""Title"": ""Sample Report"",
            ""Items"": [
                { ""Name"": ""Apple"",  ""Quantity"": 10 },
                { ""Name"": ""Banana"", ""Quantity"": 5 },
                { ""Name"": ""Cherry"", ""Quantity"": 12 }
            ]
        }";

        // Create a JsonDataSource from the JSON string using a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<[data.Title]>>");
        builder.Writeln("<<foreach [item in data.Items]>>");
        builder.Writeln(" - <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back before building the report.
        var reportDoc = new Document(templatePath);

        // Build the report using the JsonDataSource.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Write the generated report to a memory stream.
        using var outputStream = new MemoryStream();
        reportDoc.Save(outputStream, SaveFormat.Docx);

        // Optionally, display the size of the generated report.
        Console.WriteLine($"Report generated. Size: {outputStream.Length} bytes");
    }
}
