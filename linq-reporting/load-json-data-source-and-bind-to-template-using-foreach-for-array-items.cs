using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Wrapper class to align with the root object name used in the template.
    public class PersonsWrapper
    {
        // Property name matches the root name in the JSON and template.
        public Person[] persons { get; set; } = Array.Empty<Person>();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required in .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // File paths for temporary data and documents.
            string jsonPath = "persons.json";
            string templatePath = "template.docx";
            string outputPath = "Report.docx";

            // 1. Create sample JSON data.
            string jsonContent = @"{
  ""persons"": [
    { ""Name"": ""Alice"", ""Age"": 30 },
    { ""Name"": ""Bob"",   ""Age"": 25 },
    { ""Name"": ""Carol"", ""Age"": 28 }
  ]
}";
            File.WriteAllText(jsonPath, jsonContent);

            // 2. Build the template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a heading.
            builder.Writeln("Person Report");
            builder.Writeln();

            // Insert the foreach tag that iterates over the JSON array.
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age:  <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before loading it for the report).
            templateDoc.Save(templatePath);

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Create a JsonDataSource from the JSON file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // 5. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "persons");

            // 6. Save the generated report.
            doc.Save(outputPath);

            // Optional cleanup of temporary files.
            // File.Delete(jsonPath);
            // File.Delete(templatePath);
        }
    }
}
