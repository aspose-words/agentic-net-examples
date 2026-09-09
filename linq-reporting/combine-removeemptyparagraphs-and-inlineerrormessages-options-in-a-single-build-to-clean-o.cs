using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        // Name may be null to demonstrate empty paragraph removal.
        public string? Name { get; set; }

        // This property intentionally does not exist in the template to trigger an error.
        // It is used only to show inline error messages.
        public int Age { get; set; }

        // Returns an empty string so the paragraph containing only this tag becomes empty.
        public string Empty => string.Empty;
    }

    public class Wrapper
    {
        // Collection referenced by the template.
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Begin a foreach loop over the collection "persons".
            builder.Writeln("<<foreach [person in persons]>>");

            // Paragraph that will contain a value; if Name is null the paragraph becomes "Name: ".
            builder.Writeln("Name: <<[person.Name]>>");

            // Paragraph that contains only an empty tag – it will be removed by RemoveEmptyParagraphs.
            builder.Writeln("<<[person.Empty]>>");

            // This tag references a non‑existent member and will cause an inline error message.
            builder.Writeln("Missing: <<[person.NonExisting]>>");

            // End of the foreach block.
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            var data = new Wrapper
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = null, Age = 25 }, // Name is null → empty paragraph after removal.
                    new Person { Name = "Bob", Age = 40 }
                }
            };

            // 3. Configure the ReportingEngine with both options.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The returned flag indicates whether parsing succeeded (true when InlineErrorMessages is set).
            bool success = engine.BuildReport(template, data, "persons");

            Console.WriteLine($"Report build success flag: {success}");

            // 4. Save the generated document.
            const string outputPath = "Report_Output.docx";
            template.Save(outputPath);
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
