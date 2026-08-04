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
        string json = @"{ ""Name"": ""John Doe"", ""Age"": 30 }";

        // Convert the JSON string to a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build a simple template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");

        // Populate the template using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, jsonDataSource, "model");

        // Write the generated report to a memory stream.
        using var outputStream = new MemoryStream();
        doc.Save(outputStream, SaveFormat.Docx);
        outputStream.Position = 0; // Reset for potential further reading.

        // Output basic information (no interactive prompts).
        Console.WriteLine($"Report generated successfully: {success}");
        Console.WriteLine($"Output size (bytes): {outputStream.Length}");
    }
}
