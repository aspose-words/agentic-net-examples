using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data model that will be used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        // Example of a calculated property that can be used in a conditional block.
        public bool IsAdult => Age >= 18;
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the DOTX template that contains the conditional block.
            // The template should have syntax similar to:
            // <<if [persons.IsAdult]>><<[persons.Name]>> is an adult.<<else>><<[persons.Name]>> is a minor.<<endif>>
            string templatePath = @"C:\Templates\PersonReport.dotx";

            // Load the DOTX template into a Document object.
            Document reportDocument = new Document(templatePath);

            // Prepare a collection of Person objects that will be merged into the template.
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 15 },
                new Person { Name = "Carol", Age = 22 }
            };

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Allow missing members so that the template can contain default conditional blocks
            // without throwing an exception when a member is not present.
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // Optional: define a custom message for missing members (used when AllowMissingMembers is set).
            engine.MissingMemberMessage = "N/A";

            // Build the report by populating the template with the data source.
            // The third argument is the name used inside the template to reference the data source.
            engine.BuildReport(reportDocument, persons, "persons");

            // Save the populated document to a DOCX file.
            string outputPath = @"C:\Reports\PersonReport.docx";
            reportDocument.Save(outputPath);
        }
    }
}
