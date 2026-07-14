using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Custom attribute to mark properties that are allowed to be exposed in reports.
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExposeAttribute : Attribute { }

    // Original data model with some properties marked for exposure.
    public class Person
    {
        [Expose] public string Name { get; set; } = string.Empty;
        [Expose] public int Age { get; set; }
        public decimal Salary { get; set; }
    }

    // Wrapper model that only contains the exposed properties.
    public class PersonDto
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }

        public PersonDto(Person source)
        {
            Name = source.Name;
            Age = source.Age;
        }
    }

    // Root object passed to the reporting engine.
    public class ReportModel
    {
        public PersonDto Person { get; set; }

        public ReportModel(Person person)
        {
            Person = new PersonDto(person);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("Name : <<[model.Person.Name]>>");
            builder.Writeln("Age  : <<[model.Person.Age]>>");
            // This field will try to access a non‑exposed member; it will be treated as missing.
            builder.Writeln("Salary: <<[model.Person.Salary]>>");

            templateDoc.Save(templatePath);

            // 2. Load the template (simulating a separate load step).
            Document doc = new Document(templatePath);

            // 3. Restrict the original Person type so its members cannot be accessed directly.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // 4. Prepare sample data.
            Person samplePerson = new Person
            {
                Name = "John Doe",
                Age = 35, // Fixed value
                Salary = 75000m
            };

            ReportModel model = new ReportModel(samplePerson);

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "N/A"
            };

            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);
        }
    }
}
