using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingJsonAddress
{
    // Root wrapper for JSON data.
    public class DataRoot
    {
        public List<Person> Persons { get; set; } = new();
    }

    // Model representing a person.
    public class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public string Street { get; set; } = "";
        public string City { get; set; } = "";
        public string State { get; set; } = "";
        public string Zip { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create sample JSON data file.
            string jsonContent = @"{
  ""Persons"": [
    {
      ""FirstName"": ""John"",
      ""LastName"": ""Doe"",
      ""Street"": ""123 Main St"",
      ""City"": ""Springfield"",
      ""State"": ""IL"",
      ""Zip"": ""62704""
    },
    {
      ""FirstName"": ""Jane"",
      ""LastName"": ""Smith"",
      ""Street"": ""456 Oak Ave"",
      ""City"": ""Metropolis"",
      ""State"": ""NY"",
      ""Zip"": ""10001""
    }
  ]
}";
            const string jsonPath = "data.json";
            File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

            // Create the template document programmatically.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Address Report");
            builder.Writeln("<<foreach [person in Persons]>>");
            // Build address line using separate tags (no inline concatenation).
            builder.Writeln("<<[person.FirstName]>> <<[person.LastName]>>: <<[person.Street]>>, <<[person.City]>>, <<[person.State]>> <<[person.Zip]>>");
            builder.Writeln("<</foreach>>");

            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var doc = new Document(templatePath);

            // Deserialize JSON into strongly‑typed model.
            var dataRoot = JsonConvert.DeserializeObject<DataRoot>(jsonContent) ?? new DataRoot();

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, dataRoot, "root");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
