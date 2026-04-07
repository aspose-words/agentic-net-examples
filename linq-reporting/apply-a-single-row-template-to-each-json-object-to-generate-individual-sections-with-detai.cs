using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Register code page provider for any required encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data (an array of person objects).
            string jsonPath = "people.json";
            File.WriteAllText(jsonPath,
@"[
  { ""Name"": ""Alice Johnson"", ""Age"": 30, ""Address"": ""123 Maple St, Springfield"" },
  { ""Name"": ""Bob Smith"", ""Age"": 45, ""Address"": ""456 Oak Ave, Metropolis"" },
  { ""Name"": ""Carol Davis"", ""Age"": 27, ""Address"": ""789 Pine Rd, Smalltown"" }
]");

            // Create a new blank document that will serve as the template.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a foreach block that iterates over each person in the JSON array.
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("--------------------------------------------------");
            builder.Writeln("Name   : <<[person.Name]>>");
            builder.Writeln("Age    : <<[person.Age]>>");
            builder.Writeln("Address: <<[person.Address]>>");
            builder.Writeln("--------------------------------------------------");
            builder.Writeln("<</foreach>>");

            // Save the template (optional, demonstrates load/save lifecycle).
            string templatePath = "template.docx";
            templateDoc.Save(templatePath);

            // Load the template back (simulating a separate load step).
            Document doc = new Document(templatePath);

            // Create a JSON data source from the file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("persons") must match the name used in the foreach tag.
            engine.BuildReport(doc, jsonDataSource, "persons");

            // Save the generated report.
            string outputPath = "Report_Output.docx";
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
