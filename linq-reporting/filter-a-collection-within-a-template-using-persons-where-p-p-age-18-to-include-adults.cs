using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model representing a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper class that will be passed as the root data source to the reporting engine.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting foreach tag that filters the collection to adults only.
            // The expression uses Where(p => p.Age > 18) to include only persons older than 18.
            builder.Writeln("<<foreach [p in model.Persons.Where(p => p.Age > 18)]>>");
            builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice",   Age = 25 },
                    new Person { Name = "Bob",     Age = 17 },
                    new Person { Name = "Charlie", Age = 30 },
                    new Person { Name = "Diana",   Age = 15 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
