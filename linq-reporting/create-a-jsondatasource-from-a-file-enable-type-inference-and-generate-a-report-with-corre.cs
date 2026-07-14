using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string baseDir = Directory.GetCurrentDirectory();
            string jsonPath = Path.Combine(baseDir, "people.json");
            string templatePath = Path.Combine(baseDir, "template.docx");
            string resultPath = Path.Combine(baseDir, "report.docx");

            // Create sample JSON data.
            string jsonContent = @"[
  { ""Name"": ""Alice"", ""Age"": 30, ""BirthDate"": ""1990-01-01"" },
  { ""Name"": ""Bob"",   ""Age"": 25, ""BirthDate"": ""1995-05-15"" }
]";
            File.WriteAllText(jsonPath, jsonContent);

            // Build a template document with LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Loop over the JSON array named "persons".
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("Birth date: <<[person.BirthDate]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document loadedTemplate = new Document(templatePath);

            // Configure JSON data load options to enable type inference.
            JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
            {
                // Loose parsing allows the engine to infer numeric and date types.
                SimpleValueParseMode = JsonSimpleValueParseMode.Loose
            };

            // Create the JSON data source.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

            // Build the report using the data source named "persons".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, jsonDataSource, "persons");

            // Save the generated report.
            loadedTemplate.Save(resultPath);
        }
    }
}
