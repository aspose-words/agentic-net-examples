using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingEngineDemo
{
    // Simple data class that will be used as the data source for the report.
    public class Person
    {
        // This property contains Markdown formatted text.
        public string Bio { get; set; }

        public Person(string bio)
        {
            Bio = bio;
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document in memory.
            // -----------------------------------------------------------------
            Document template = new Document();                     // Create a new blank document.
            DocumentBuilder builder = new DocumentBuilder(template); // Helper to add content.

            // Insert a placeholder that will be replaced with Markdown text.
            // The syntax <<[person.Bio]:markdown>> tells the ReportingEngine to
            // treat the replacement as Markdown.
            builder.Writeln("<<[person.Bio]:markdown>>");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            // Example Markdown content.
            string markdownBio = @"
# John Doe

**Software Engineer** with 10+ years of experience.

- C#
- .NET
- ASP.NET Core

> ""Code is like humor. When you have to explain it, it's bad.""
";

            // Create an instance of the data source.
            Person person = new Person(markdownBio);

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // BuildReport will replace the placeholder with the Markdown content.
            // The second parameter is the data source object.
            // The third parameter is the name used to reference the data source in the template.
            engine.BuildReport(template, person, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "PersonReport.docx");
            template.Save(outputPath);

            Console.WriteLine($"Report generated and saved to: {outputPath}");
        }
    }
}
