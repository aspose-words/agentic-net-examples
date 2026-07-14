using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictMembersExample
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write tags that reference the Person object's members.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine to restrict access to the Person type.
            //    This is the security measure requested (equivalent to <<restrictMembers>>).
            // -----------------------------------------------------------------
            // Restrict all members of the Person type.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // Allow missing members so that restricted members are rendered as empty strings
            // instead of throwing an exception.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            engine.MissingMemberMessage = string.Empty; // Optional: customize missing member output.

            // -----------------------------------------------------------------
            // 4. Build the report using a Person instance as the data source.
            // -----------------------------------------------------------------
            Person person = new Person { Name = "John Doe", Age = 42 };

            // The root object name used in the template is "person".
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);

            // The example finishes without waiting for user input.
        }
    }
}
