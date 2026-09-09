using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // JsonDataSource resides in this namespace

namespace AsposeWordsLinqReporting
{
    // Simple data model that matches the JSON structure.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for possible legacy encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // ---------- 1. Create a template document in memory ----------
            // The template contains LINQ Reporting tags that reference the root object "person".
            using var templateStream = new MemoryStream();
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Person Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template to the memory stream.
            templateDoc.Save(templateStream, SaveFormat.Docx);
            templateStream.Position = 0; // Reset for reading.

            // ---------- 2. Prepare JSON data source ----------
            // Sample JSON representing a single Person object.
            const string json = @"{ ""Name"": ""John Doe"", ""Age"": 30 }";
            var jsonBytes = Encoding.UTF8.GetBytes(json);
            using var jsonStream = new MemoryStream(jsonBytes);
            jsonStream.Position = 0; // Ensure the stream is at the beginning.

            // Create a JsonDataSource from the JSON stream.
            var jsonDataSource = new JsonDataSource(jsonStream);

            // ---------- 3. Load the template document from the stream ----------
            var reportDoc = new Document(templateStream);

            // ---------- 4. Build the report ----------
            var engine = new ReportingEngine();
            // The root object name used in the template is "person".
            engine.BuildReport(reportDoc, jsonDataSource, "person");

            // ---------- 5. Save the generated report as RTF ----------
            const string outputPath = "PersonReport.rtf";
            reportDoc.Save(outputPath, SaveFormat.Rtf);

            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
