using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRemoveEmptyParagraphs
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Collection that will be iterated over in the template.
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Title.
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin foreach loop over the collection "Persons".
            builder.Writeln("<<foreach [person in Persons]>>");

            // Output name and age for each person.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Conditional line – will be empty for persons younger than 30.
            // The empty paragraph generated for those items will be removed.
            builder.Writeln("<<if [person.Age >= 30]>>Senior<</if>>");

            // End of foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required by the workflow).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 28 },
                    new Person { Name = "Bob",   Age = 35 },
                    new Person { Name = "Carol", Age = 22 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the model; the root name is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the final document.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
