using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Valid tag – will be replaced with the person's name.
            builder.Writeln("Hello <<[person.Name]>>!");

            // Invalid tag – property 'Age' does not exist on Person.
            // With InlineErrorMessages enabled this will be shown in the output
            // and BuildReport will return false.
            builder.Writeln("Missing property: <<[person.Age]>>");

            // 2. Prepare the data source.
            Person person = new Person();

            // 3. Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // 4. Build the report. Capture the success flag.
            bool success = engine.BuildReport(doc, person, "person");

            // 5. Save the generated document.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // 6. Output the result of the build operation.
            Console.WriteLine($"Report build successful: {success}");
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
