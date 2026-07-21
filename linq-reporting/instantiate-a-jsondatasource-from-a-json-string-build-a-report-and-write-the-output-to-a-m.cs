using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON data.
        const string json = @"{
  ""Title"": ""Sample Report"",
  ""Items"": [
    { ""Name"": ""Apple"",  ""Quantity"": 5 },
    { ""Name"": ""Banana"", ""Quantity"": 12 },
    { ""Name"": ""Cherry"", ""Quantity"": 7 }
  ]
}";

        // Convert the JSON string to a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build a simple template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Populate the template using the reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "model");

        // Save the generated report to a memory stream.
        using var output = new MemoryStream();
        doc.Save(output, SaveFormat.Docx);

        // Output the size of the generated report.
        Console.WriteLine($"Report generated. Size: {output.Length} bytes.");
    }
}
