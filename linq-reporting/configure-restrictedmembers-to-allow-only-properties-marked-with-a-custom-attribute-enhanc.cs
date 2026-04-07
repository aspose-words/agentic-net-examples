using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Custom attribute that marks properties allowed for reporting.
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ReportableAttribute : Attribute { }

    // Sample data model.
    public class Person
    {
        // This property is allowed because it has the Reportable attribute.
        [Reportable]
        public string Name { get; set; } = "John Doe";

        // This property will be blocked by the restricted‑members configuration.
        public int Age { get; set; } = 30;
    }

    // Wrapper exposing only the reportable members of Person.
    public class PersonReport
    {
        public string Name { get; set; }

        public PersonReport(Person source)
        {
            // Copy only the properties that are marked with [Reportable].
            Name = source.Name;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a simple template document with two tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age:  <<[person.Age]>>");
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template (demonstrates the required load step).
            Document doc = new Document(templatePath);

            // 3. Configure the ReportingEngine to ignore missing members.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Optional: customize the message shown for missing members.
            engine.MissingMemberMessage = string.Empty;

            // 4. Build the report using a wrapper that exposes only the allowed members.
            Person data = new Person();
            PersonReport reportData = new PersonReport(data);
            engine.BuildReport(doc, reportData, "person");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
