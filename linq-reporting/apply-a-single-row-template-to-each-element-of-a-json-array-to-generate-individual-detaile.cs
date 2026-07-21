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
            // Register code page provider for any legacy encodings used by Aspose.Words.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Define file paths in the current working directory.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "Template.docx");
            string jsonPath = Path.Combine(workDir, "Data.json");
            string outputPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a simple JSON file that contains an array of person objects.
            // -----------------------------------------------------------------
            string jsonContent = @"[
    { ""Name"": ""John Doe"", ""Age"": 30, ""Address"": ""123 Main St"" },
    { ""Name"": ""Jane Smith"", ""Age"": 25, ""Address"": ""456 Oak Ave"" },
    { ""Name"": ""Bob Johnson"", ""Age"": 40, ""Address"": ""789 Pine Rd"" }
]";
            File.WriteAllText(jsonPath, jsonContent);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            //    The template uses a foreach tag to iterate over the JSON array.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Start the foreach block – each iteration will create a separate section.
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("Address: <<[person.Address]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the JSON data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // The root name used in the template tags is "persons".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, jsonDataSource, "persons");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
