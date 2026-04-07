using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    // Custom static helper that will be accessed from the template without reflection.
    public static class MyHelper
    {
        public static string ToUpper(string value) => value?.ToUpperInvariant() ?? string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Original name: <<[person.Name]>>");
            builder.Writeln("Upper‑case name (using custom static type): <<[MyHelper.ToUpper(person.Name)]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template back (required before building the report).
            var doc = new Document(templatePath);

            // 3. Prepare the data source.
            var person = new Person { Name = "Alice Smith" };

            // 4. Configure the ReportingEngine.
            var engine = new ReportingEngine();

            // Register the custom external type so the template can call its static members.
            engine.KnownTypes.Add(typeof(MyHelper));

            // 5. Build the report.
            // The root object name must match the name used in the template tags ("person").
            engine.BuildReport(doc, person, "person");

            // 6. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
