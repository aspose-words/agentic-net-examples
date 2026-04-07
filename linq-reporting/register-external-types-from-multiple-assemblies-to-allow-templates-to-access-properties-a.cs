using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders for the template and the generated report.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a Word template programmatically and insert LINQ tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple property access.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Use a static method from an external assembly (Newtonsoft.Json) to serialize the object.
            // The ReportingEngine must know about JsonConvert, so we will register it later.
            builder.Writeln("JSON: <<[JsonConvert.SerializeObject(person)]>>");

            // Save the template to disk (required before building the report).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source (root object) and the reporting engine.
            // -----------------------------------------------------------------
            Person person = new Person();

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register external types that the template may reference.
            // - JsonConvert from Newtonsoft.Json (different assembly).
            // - Person from the current assembly (demonstrates multiple assemblies).
            engine.KnownTypes.Add(typeof(JsonConvert));
            engine.KnownTypes.Add(typeof(Person));

            // Build the report. The root name "person" must match the tag prefix used in the template.
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
