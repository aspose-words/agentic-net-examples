using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedMembersDemo
{
    // Custom attribute to mark exposable properties
    [AttributeUsage(AttributeTargets.Property)]
    public class ExposeAttribute : Attribute { }

    // Sample data model
    public class Person
    {
        [Expose]
        public string Name { get; set; } = "John Doe";

        public int Age { get; set; } = 30;

        [Expose]
        public string Email { get; set; } = "john.doe@example.com";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // Create a simple template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert tags that reference the Person object.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("Email: <<[person.Email]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and configure the reporting engine.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine
            {
                // Show inline error messages if a tag cannot be resolved.
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // -----------------------------------------------------------------
            // (Optional) Demonstrate how one could restrict members if the API
            // were available. The current Aspose.Words version used in this
            // environment does not expose the RestrictedMembers class, so this
            // section is omitted to keep the example compilable.
            // -----------------------------------------------------------------
            // Example (commented out):
            // var restricted = new RestrictedMembers();
            // foreach (PropertyInfo prop in typeof(Person).GetProperties())
            // {
            //     if (prop.IsDefined(typeof(ExposeAttribute), false))
            //         restricted.AddAllowedMember(typeof(Person), prop.Name);
            // }
            // engine.RestrictedMembers = restricted;

            // -----------------------------------------------------------------
            // Prepare the data source.
            // -----------------------------------------------------------------
            Person person = new Person();

            // -----------------------------------------------------------------
            // Build the report.
            // -----------------------------------------------------------------
            bool success = engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // -----------------------------------------------------------------
            // Indicate the result (no interactive input).
            // -----------------------------------------------------------------
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
