using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingFirstCharStyling
{
    // Sample data model.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Create the template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin foreach loop over Persons.
            builder.Writeln("<<foreach [p in Persons]>>");

            // Write each person's name with the first character in red.
            // First character styled with textColor, rest normal.
            builder.Writeln(
                "<<textColor [\"Red\"]>><<[p.Name.Substring(0,1)]>><</textColor>><<[p.Name.Substring(1)]>> (Age: <<[p.Age]>>)");

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            builder.Document.Save(templatePath);

            // Load the template for report generation.
            var templateDoc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save the final report.
            var outputPath = "Report.docx";
            templateDoc.Save(outputPath);
        }
    }
}
