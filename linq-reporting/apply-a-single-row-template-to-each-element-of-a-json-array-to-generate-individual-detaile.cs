using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingJsonArrayExample
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data
            string jsonContent = @"{
                ""Persons"": [
                    { ""Name"": ""Alice"", ""Age"": 30 },
                    { ""Name"": ""Bob"",   ""Age"": 25 },
                    { ""Name"": ""Charlie"", ""Age"": 28 }
                ]
            }";

            string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
            File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

            // Deserialize JSON into the model
            Model model = JsonConvert.DeserializeObject<Model>(File.ReadAllText(jsonPath, Encoding.UTF8))!;

            // Create the LINQ Reporting template programmatically
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a title for the report
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin foreach loop over the Persons collection
            builder.Writeln("<<foreach [person in Persons]>>");

            // Each person gets its own section (heading + paragraph)
            builder.Writeln("<<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln(); // blank line between entries

            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // Load the template for report generation
            Document reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
