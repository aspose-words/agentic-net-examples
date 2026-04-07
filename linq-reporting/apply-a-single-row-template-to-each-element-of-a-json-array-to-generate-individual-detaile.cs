using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Enable legacy encodings (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths in the current working directory.
            string jsonPath = Path.Combine(Environment.CurrentDirectory, "people.json");
            string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // Sample JSON array of people.
            string jsonContent = @"
[
    { ""Name"": ""Alice Johnson"", ""Age"": 30, ""Bio"": ""Software engineer with 8 years of experience."" },
    { ""Name"": ""Bob Smith"", ""Age"": 45, ""Bio"": ""Project manager specializing in agile methodologies."" },
    { ""Name"": ""Carol Lee"", ""Age"": 27, ""Bio"": ""UX designer passionate about user-centered design."" }
]";
            File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

            // Create a Word document that will serve as the template.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("People Report");
            builder.Writeln("--------------------");

            // LINQ Reporting tags.
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("Bio: <<[person.Bio]>>");
            builder.Writeln(""); // Blank line between entries.
            builder.Writeln("<</foreach>>");

            // Save the template (required before loading it for reporting).
            templateDoc.Save(templatePath);

            // Load JSON data source.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Load the template document.
            Document doc = new Document(templatePath);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "persons");

            // Save the generated report.
            doc.Save(outputPath);

            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
