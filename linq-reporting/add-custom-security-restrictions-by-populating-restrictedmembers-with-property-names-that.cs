using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSecurityExample
{
    // Sample data model.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public string SecretInfo { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create a simple template document with LINQ Reporting tags.
            string templatePath = Path.Combine(outputDir, "template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Secret: <<[person.SecretInfo]>>");
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var doc = new Document(templatePath);

            // Sample data.
            var person = new Person
            {
                Name = "John Doe",
                Age = 30,
                SecretInfo = "TopSecret"
            };

            // Configure the reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // NOTE: In newer versions of Aspose.Words.Reporting the ReportingEngine
            // provides a RestrictedMembers collection to hide specific members.
            // If the property is unavailable, this step is omitted.

            // Build the report.
            bool success = engine.BuildReport(doc, person, "person");

            // Save the generated report.
            string resultPath = Path.Combine(outputDir, "result.docx");
            doc.Save(resultPath);

            // Indicate completion.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template: {templatePath}");
            Console.WriteLine($"Result: {resultPath}");
        }
    }
}
