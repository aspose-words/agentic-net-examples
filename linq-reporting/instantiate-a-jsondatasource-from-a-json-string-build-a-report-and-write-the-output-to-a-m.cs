using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class JsonReportExample
{
    public static void Main()
    {
        // Register code page provider for required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data.
        string json = @"{
            ""CustomerName"": ""John Doe"",
            ""OrderDate"": ""2023-01-01"",
            ""Items"": [
                { ""Index"": 1, ""Name"": ""Item A"", ""Price"": 10.5 },
                { ""Index"": 2, ""Name"": ""Item B"", ""Price"": 20.0 },
                { ""Index"": 3, ""Name"": ""Item C"", ""Price"": 15.75 }
            ]
        }";

        // Load JSON from a memory stream (JsonDataSource expects a path or stream).
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        var dataSource = new JsonDataSource(jsonStream, new JsonDataLoadOptions());

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Date: <<[order.OrderDate]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>> - $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "order");

        // Write the generated report to a memory stream.
        using var outputStream = new MemoryStream();
        doc.Save(outputStream, SaveFormat.Docx);
        outputStream.Position = 0;

        // Output the size of the generated document.
        Console.WriteLine($"Report generated successfully. Size: {outputStream.Length} bytes.");
    }
}
