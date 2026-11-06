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
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data
        string json = @"{
            ""people"": [
                { ""Name"": ""Alice"", ""Age"": 30 },
                { ""Name"": ""Bob"", ""Age"": 25 }
            ]
        }";

        // Create a Word document template programmatically
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in people]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Load JSON data source from a memory stream (JsonDataSource expects a path or stream)
        using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
        {
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // Build the report
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "people");

            // Write the generated report to a memory stream
            using (MemoryStream output = new MemoryStream())
            {
                doc.Save(output, SaveFormat.Docx);
                Console.WriteLine($"Report generated, size: {output.Length} bytes");
            }
        }
    }
}
