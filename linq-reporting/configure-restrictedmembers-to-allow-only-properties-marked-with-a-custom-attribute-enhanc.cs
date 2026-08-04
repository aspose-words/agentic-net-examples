using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingRestrictedMembers
{
    // Custom attribute to mark properties that are allowed to be accessed in templates.
    [AttributeUsage(AttributeTargets.Property)]
    public class ExposeAttribute : Attribute { }

    // Original data model with mixed exposure.
    public class Person
    {
        [Expose] public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Secret { get; set; } = string.Empty;
    }

    // Wrapper exposing only properties marked with [Expose].
    public class PersonWrapper
    {
        private readonly Person _person;

        public PersonWrapper(Person person) => _person = person;

        // Only expose properties that have the ExposeAttribute.
        public string Name => _person.Name;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("People Report");
            builder.Writeln("----------------");
            // Use a foreach loop to iterate over the collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            // This property is not exposed; it will be treated as missing.
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // 2. Load the template (demonstrates the load step).
            Document doc = new Document(templatePath);

            // 3. Prepare sample data.
            List<Person> people = new()
            {
                new Person { Name = "Alice", Age = 30, Secret = "Loves cats" },
                new Person { Name = "Bob", Age = 45, Secret = "Speaks French" },
                new Person { Name = "Charlie", Age = 28, Secret = "Plays guitar" }
            };

            // Wrap each Person so that only exposed members are reachable.
            List<PersonWrapper> wrappedPeople = new();
            foreach (var p in people)
                wrappedPeople.Add(new PersonWrapper(p));

            // 4. Configure the ReportingEngine.
            // Restrict the original Person type so its members cannot be accessed directly.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members to avoid exceptions for the restricted Age property.
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Provide a friendly message for missing members.
            engine.MissingMemberMessage = "Restricted";

            // 5. Build the report.
            // The root object name must match the tag used in the template ("Persons").
            engine.BuildReport(doc, wrappedPeople, "Persons");

            // 6. Save the generated report.
            string resultPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(resultPath);
        }
    }
}
