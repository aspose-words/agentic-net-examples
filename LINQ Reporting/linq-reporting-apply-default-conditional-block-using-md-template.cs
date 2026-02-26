using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public bool IsAdult => Age >= 18; // Used in conditional blocks.
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the template document that contains LINQ Reporting tags.
            // The template can include a conditional block like:
            // <<if [person.IsAdult]>><<[person.Name]>> is an adult.<</if>>
            // <<else>><<[person.Name]>> is a minor.<</else>>
            string templatePath = Path.Combine("Templates", "PersonReportTemplate.docx");

            // Load the template document.
            Document template = new Document(templatePath);

            // Create a data source – an array of Person objects.
            Person[] people = new[]
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 15 },
                new Person { Name = "Carol", Age = 22 }
            };

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw if a tag references a non‑existent member.
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs,
                // Message to display for missing members (optional).
                MissingMemberMessage = "N/A"
            };

            // Build the report. The data source name "person" matches the tags used in the template.
            // Using the overload that accepts an array of data sources allows us to reference the collection directly.
            engine.BuildReport(template, new object[] { people }, new[] { "person" });

            // Save the generated report.
            string outputPath = Path.Combine("Output", "PersonReport.docx");
            template.Save(outputPath);
        }
    }
}
