using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a foreach loop that iterates over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>   Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure restricted members for security.
            //    Here we restrict access to System.Type members (e.g., GetType()).
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            // -----------------------------------------------------------------
            // 4. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob",   Age = 45 },
                    new Person { Name = "Carol", Age = 27 }
                }
            };

            // -----------------------------------------------------------------
            // 5. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers; // Use explicit assignment.
            engine.MissingMemberMessage = "N/A"; // Message for any restricted or missing members.

            // The root object name in the template is "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
        }
    }
}
