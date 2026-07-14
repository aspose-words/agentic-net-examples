using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used in the report.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags that reference the data model.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document (required before building the report).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure restricted members for security.
            //    Here we restrict access to System.Type members (e.g., GetType()).
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members to avoid exceptions if a tag references a restricted member.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Optional: customize the message shown for missing members.
            engine.MissingMemberMessage = "[Restricted]";

            // Prepare the data source.
            Person person = new Person { Name = "Alice Smith", Age = 42 };

            // Build the report. The root object name in the template is "person".
            engine.BuildReport(loadedTemplate, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(outputDir, "Report.docx");
            loadedTemplate.Save(resultPath);

            // The example finishes without requiring user interaction.
        }
    }
}
