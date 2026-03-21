using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonReport
{
    class Program
    {
        static void Main()
        {
            // Create a simple Word template in memory with reporting placeholders.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin the collection loop.
            builder.Writeln("{{#products}}");
            // Row with placeholders and an inline arithmetic expression.
            builder.Writeln("{{Name}} {{Price}} {{Discount}} {{Price * (1 - Discount/100)}}");
            // End the collection loop.
            builder.Writeln("{{/products}}");

            // JSON data as a string.
            string json = @"
            [
                { ""Name"": ""Product A"", ""Price"": 100.0, ""Discount"": 10 },
                { ""Name"": ""Product B"", ""Price"": 250.0, ""Discount"": 15 }
            ]";

            // Load JSON data from a memory stream.
            using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
            JsonDataSource jsonData = new JsonDataSource(jsonStream);

            // Build the report using the root element name "products".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonData, "products");

            // Save the generated document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "DiscountReport.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
