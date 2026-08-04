using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model.
    public class Person
    {
        // Name may be null to simulate missing data.
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Direct reference to a possibly missing member.
            builder.Writeln("First person name: <<[model.Persons[0].Name]>>");

            // Loop over the collection; missing members will be treated as empty strings.
            builder.Writeln("<<foreach [person in model.Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample data with a null name.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice" },
                    new Person { Name = null }, // Missing name.
                    new Person { Name = "Charlie" }
                }
            };

            // Configure the reporting engine to treat missing members as empty strings.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Optional: customize the placeholder text for missing members.
            engine.MissingMemberMessage = string.Empty;

            // Build the report. The root object name used in the template is "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
