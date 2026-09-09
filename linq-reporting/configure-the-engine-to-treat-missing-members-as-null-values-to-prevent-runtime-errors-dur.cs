using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with only a Name property.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        // Note: Age property is intentionally omitted to demonstrate missing‑member handling.
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags. The Age tag does not exist in the Person class.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>"); // Missing member.

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template (simulating a separate load step).
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            Person person = new Person { Name = "John Doe" };

            // 4. Configure the ReportingEngine to treat missing members as null.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "N/A" // Optional custom message for plain missing references.
            };

            // 5. Build the report. The root object name is "person".
            bool success = engine.BuildReport(doc, person, "person");

            // 6. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Inform the user.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
