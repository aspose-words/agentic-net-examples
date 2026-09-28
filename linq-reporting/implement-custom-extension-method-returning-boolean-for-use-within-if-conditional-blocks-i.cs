using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExtensionExample
{
    // Data model representing a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }

        // Instance method used inside the template's if condition.
        public bool IsAdult() => Age >= 18;
    }

    // Wrapper model passed as the root object to the reporting engine.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
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
                    new Person { Name = "Bob", Age = 15 },
                    new Person { Name = "Charlie", Age = 22 }
                }
            };

            // Create the template document programmatically.
            const string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write LINQ Reporting tags.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("<<if [p.IsAdult()]>> (Adult) <</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var doc = new Document(templatePath);

            // Build the report using the model.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "output.docx";
            doc.Save(outputPath);
        }
    }
}
