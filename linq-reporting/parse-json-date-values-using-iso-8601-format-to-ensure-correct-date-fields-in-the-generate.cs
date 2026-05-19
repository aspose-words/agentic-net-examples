using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting; // Contains ReportingEngine, JsonDataSource, JsonDataLoadOptions

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Working directory for all generated files.
            string workDir = Directory.GetCurrentDirectory();

            // File paths for the template, JSON data and the final report.
            string templatePath = Path.Combine(workDir, "template.docx");
            string jsonPath = Path.Combine(workDir, "people.json");
            string reportPath = Path.Combine(workDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create a simple Word template with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Persons Report");
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Birth Date: <<[person.BirthDate]>>");
            builder.Writeln("<</foreach>>");

            // Save the template (required before building the report).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Create a JSON file containing ISO 8601 date strings.
            // -----------------------------------------------------------------
            var people = new[]
            {
                new { Name = "Alice", BirthDate = "1990-05-15T00:00:00Z" },
                new { Name = "Bob",   BirthDate = "1985-12-01T00:00:00Z" }
            };
            string jsonContent = System.Text.Json.JsonSerializer.Serialize(
                people,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonPath, jsonContent);

            // -----------------------------------------------------------------
            // 3. Configure JSON data load options to recognize ISO 8601 formats.
            // -----------------------------------------------------------------
            JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
            {
                // Additional explicit formats (optional – ISO‑8601 is recognized by default).
                ExactDateTimeParseFormats = new List<string>
                {
                    "yyyy-MM-ddTHH:mm:ssZ",
                    "yyyy-MM-ddTHH:mm:ss"
                },
                // Ensure a root object is generated so the template can reference "persons".
                AlwaysGenerateRootObject = true
            };

            // Create the JSON data source.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

            // -----------------------------------------------------------------
            // 4. Load the template document (required before BuildReport).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 5. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None // default behavior
            };

            // The root name used in the template is "persons".
            engine.BuildReport(doc, jsonDataSource, "persons");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
