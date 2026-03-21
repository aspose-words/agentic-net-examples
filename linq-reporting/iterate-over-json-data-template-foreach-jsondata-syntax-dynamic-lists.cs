using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonForeachDemo
{
    class Program
    {
        static void Main()
        {
            // Create a simple Word template in memory with a <<foreach>> tag.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("<<foreach [in persons]>>");
            builder.Writeln("Name: <<[Name]>>");
            builder.Writeln("Age: <<[Age]>>");
            builder.Writeln("<</foreach>>");

            // JSON data source as a string.
            string json = @"{
                ""persons"": [
                    { ""Name"": ""John Doe"", ""Age"": 30 },
                    { ""Name"": ""Jane Smith"", ""Age"": 25 }
                ]
            }";

            // Load JSON from a memory stream.
            using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "persons");

            // Save the populated document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PersonsReport.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
