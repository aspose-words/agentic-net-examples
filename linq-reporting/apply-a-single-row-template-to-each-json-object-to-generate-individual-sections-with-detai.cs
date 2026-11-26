using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingExample
{
    // Data model representing a single JSON object.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public string Email { get; set; } = "";
    }

    // Wrapper root object for the report.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data.
            string jsonPath = "persons.json";
            var sampleData = new List<Person>
            {
                new() { Name = "Alice Johnson", Age = 30, Email = "alice@example.com" },
                new() { Name = "Bob Smith", Age = 45, Email = "bob@example.com" },
                new() { Name = "Carol Davis", Age = 27, Email = "carol@example.com" }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // Load JSON into the wrapper model.
            var persons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath)) ?? new();
            var model = new ReportModel { Persons = persons };

            // Create the LINQ Reporting template programmatically.
            string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Persons Report ===");
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("----");
            builder.Writeln("Name : <<[person.Name]>>");
            builder.Writeln("Age  : <<[person.Age]>>");
            builder.Writeln("Email: <<[person.Email]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("----");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string outputPath = "output.docx";
            doc.Save(outputPath);
        }
    }
}
