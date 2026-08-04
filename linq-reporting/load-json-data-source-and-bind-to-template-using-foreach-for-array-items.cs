using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Working directory.
            string workDir = Directory.GetCurrentDirectory();

            // 1. Create sample JSON data file.
            string jsonPath = Path.Combine(workDir, "people.json");
            File.WriteAllText(jsonPath,
@"[
  { ""Name"": ""Alice"", ""Age"": 30 },
  { ""Name"": ""Bob"",   ""Age"": 25 },
  { ""Name"": ""Charlie"", ""Age"": 28 }
]");

            // 2. Build a template document with LINQ Reporting tags.
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("- <<[p.Name]>> is <<[p.Age]>> years old.");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // 3. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 4. Create a JsonDataSource from the JSON file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.
            // The root name "persons" must match the name used in the template tags.
            engine.BuildReport(reportDoc, jsonDataSource, "persons");

            // 6. Save the generated report.
            string outputPath = Path.Combine(workDir, "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
