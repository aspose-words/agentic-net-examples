using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model representing a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper class that will be passed to the reporting engine.
    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a heading.
            builder.Writeln("LINQ Reporting – Retrieve the fourth element");
            builder.Writeln();

            // Insert a LINQ Reporting tag that accesses the fourth element (index 3) of the collection.
            // The root data source name will be "model", so we reference it accordingly.
            builder.Writeln("Fourth person: <<[model.Persons.ElementAt(3).Name]>>");
            builder.Writeln("Age: <<[model.Persons.ElementAt(3).Age]>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new Model
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice",   Age = 30 },
                    new Person { Name = "Bob",     Age = 25 },
                    new Person { Name = "Charlie", Age = 28 },
                    new Person { Name = "Diana",   Age = 32 }, // Fourth element (index 3)
                    new Person { Name = "Eve",     Age = 27 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport with the root object name "model" to match the tags.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
