using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model for JSON serialization (optional, not used directly by the engine)
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for .NET Core (required by Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create sample JSON data file.
            // -----------------------------------------------------------------
            string jsonFilePath = "people.json";
            string jsonContent = @"[
                { ""Name"": ""Alice"", ""Age"": 30 },
                { ""Name"": ""Bob"",   ""Age"": 25 },
                { ""Name"": ""Charlie"", ""Age"": 28 }
            ]";
            File.WriteAllText(jsonFilePath, jsonContent, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Build a template document containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            string templatePath = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin foreach loop over the JSON array named "persons"
            builder.Writeln("<<foreach [p in persons]>>");
            // Output each person's name and age
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the JSON data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            JsonDataSource jsonDataSource = new JsonDataSource(jsonFilePath);

            // The root name "persons" must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, jsonDataSource, "persons");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
