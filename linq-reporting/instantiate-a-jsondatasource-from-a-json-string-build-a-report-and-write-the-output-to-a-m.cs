using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple JSON string that will serve as the data source
        string json = @"{
            ""Title"": ""Sample Report"",
            ""Items"": [
                { ""Index"": 1, ""Name"": ""First item"" },
                { ""Index"": 2, ""Name"": ""Second item"" },
                { ""Index"": 3, ""Name"": ""Third item"" }
            ]
        }";

        // Build the template document programmatically
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the title tag
        builder.Writeln("<<[json.Title]>>");
        builder.Writeln(); // empty line

        // Begin a foreach loop over the Items collection
        builder.Writeln("<<foreach [item in json.Items]>>");
        // Write each item's fields
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        // End the foreach loop
        builder.Writeln("<</foreach>>");

        // Create a JsonDataSource from the JSON string using a memory stream
        using MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        jsonStream.Position = 0; // Ensure the stream is at the beginning
        JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

        // Build the report using the ReportingEngine
        ReportingEngine engine = new ReportingEngine();
        // The root name "json" matches the tags used in the template
        engine.BuildReport(doc, jsonDataSource, "json");

        // Save the generated report to a memory stream (DOCX format)
        using MemoryStream outputStream = new MemoryStream();
        doc.Save(outputStream, SaveFormat.Docx);

        // Optionally, display the size of the generated document
        Console.WriteLine($"Report generated. Output size: {outputStream.Length} bytes.");
    }
}
