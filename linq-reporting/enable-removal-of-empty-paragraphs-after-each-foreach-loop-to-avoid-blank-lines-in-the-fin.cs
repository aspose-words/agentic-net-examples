using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a foreach block that iterates over the Persons collection.
            // Each person's name will be written on its own line.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("<<[p.Name]>>");
            builder.Writeln("<</foreach>>");

            // Configure the reporting engine to remove paragraphs that become empty
            // after the tags are processed (e.g., the paragraph after the closing tag).
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Prepare sample data.
            ReportModel model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice" });
            model.Persons.Add(new Person { Name = "Bob" });
            model.Persons.Add(new Person { Name = "Charlie" });

            // Build the report using the model as the root data source.
            engine.BuildReport(template, model, "model");

            // Save the resulting document.
            template.Save("ReportWithNoEmptyParagraphs.docx");
        }
    }
}
