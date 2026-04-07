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
            // Ensure the output folder exists.
            const string outputDir = "Output";
            Directory.CreateDirectory(outputDir);

            // 1. Create a sample JSON file.
            const string jsonFileName = "people.json";
            string jsonContent = @"{
  ""people"": [
    { ""Name"": ""Alice"", ""Age"": 30 },
    { ""Name"": ""Bob"" }
  ]
}";
            File.WriteAllText(jsonFileName, jsonContent);

            // 2. Create a JsonDataSource from the JSON file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonFileName);

            // 3. Build the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a heading.
            builder.Writeln("People Report");
            builder.Writeln("----------------");

            // Insert a foreach loop that iterates over the "people" array.
            // Missing members (e.g., Age for Bob) will be treated as null because of the AllowMissingMembers option.
            builder.Writeln("<<foreach [person in people]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // 4. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers; // Treat missing members as null.
            engine.MissingMemberMessage = "N/A"; // Optional custom text for missing plain member references.

            // 5. Build the report. No data source name is required because the template references members directly.
            engine.BuildReport(template, jsonDataSource, "");

            // 6. Save the generated report.
            string outputPath = Path.Combine(outputDir, "PeopleReport.docx");
            template.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
