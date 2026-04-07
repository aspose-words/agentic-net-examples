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
        string json = @"{
            ""Title"": ""Sample Report"",
            ""Items"": [
                { ""Name"": ""Apple"",  ""Quantity"": 5 },
                { ""Name"": ""Banana"", ""Quantity"": 12 },
                { ""Name"": ""Cherry"", ""Quantity"": 7 }
            ]
        }";

        // Convert JSON string to a stream for JsonDataSource.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));

        // Create the JSON data source.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build a simple template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title placeholder.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data rows using a foreach tag.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        // End the table while still inside the foreach block.
        builder.EndTable();
        // Close the foreach tag after the table.
        builder.Writeln("<</foreach>>");

        // Build the report using the JSON data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "model");

        // Save the generated report to a memory stream.
        using var outputStream = new MemoryStream();
        doc.Save(outputStream, SaveFormat.Docx);
        outputStream.Position = 0; // Reset for potential further use.

        Console.WriteLine($"Report generated. Size: {outputStream.Length} bytes.");
    }
}
