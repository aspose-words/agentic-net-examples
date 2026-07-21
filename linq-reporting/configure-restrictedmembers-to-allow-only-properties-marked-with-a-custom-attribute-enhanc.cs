using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Custom attribute to mark properties that are allowed to be exposed in reports.
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ReportableAttribute : Attribute { }

    // Original data model with several properties. Only those marked with [Reportable] will be exposed.
    public class Person
    {
        [Reportable]
        public string Name { get; set; } = string.Empty;

        public int Age { get; set; }

        public string Secret { get; set; } = string.Empty;
    }

    // Wrapper model that contains only the properties marked with [Reportable] from the original model.
    public class PersonReport
    {
        public string Name { get; set; } = string.Empty;

        public PersonReport(Person source)
        {
            // Copy only the allowed property.
            Name = source.Name;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the Aspose.Words license is not required for this example.

            // 1. Create a simple template document with LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The template tries to access both allowed and disallowed members.
            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>"); // Age is not marked as Reportable.
            builder.Writeln("Secret: <<[person.Secret]>>"); // Secret is not marked as Reportable.

            // 2. Restrict the original Person type so its members cannot be accessed directly.
            // This must be done before any report is built.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // 3. Prepare sample data.
            Person samplePerson = new Person
            {
                Name = "John Doe",
                Age = 42,
                Secret = "TopSecret"
            };

            // Wrap the data so only allowed properties are exposed.
            PersonReport wrapper = new PersonReport(samplePerson);

            // 4. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members to avoid exceptions for restricted members.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Optional: provide a friendly message for missing members.
                MissingMemberMessage = "[Data not available]"
            };

            // 5. Build the report using the wrapper as the root object.
            engine.BuildReport(doc, wrapper, "person");

            // 6. Save the generated report.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PersonReport.docx");
            doc.Save(outputPath);
        }
    }
}
