using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model with only getter properties (read‑only).
    public class PersonReadOnly
    {
        public string Name { get; }
        public int Age { get; }

        public PersonReadOnly(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    // Original mutable model (used only to demonstrate restriction).
    public class PersonMutable
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // 2. Load the template (simulating a separate load step).
            Document doc = new Document(templatePath);

            // 3. Configure restricted members: block the mutable type so that only the read‑only type can be used.
            // This prevents templates from accessing setters on PersonMutable.
            ReportingEngine.SetRestrictedTypes(typeof(PersonMutable));

            // 4. Prepare the data source (read‑only instance).
            PersonReadOnly person = new PersonReadOnly("Alice Johnson", 35);

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, person, "person");

            // 6. Save the generated report.
            string resultPath = Path.Combine(outputDir, "Result.docx");
            doc.Save(resultPath);
        }
    }
}
