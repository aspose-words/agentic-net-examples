using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // External type whose static members will be accessed from the template.
    public static class Utils
    {
        public static string GetGreeting(string name) => $"Hello, {name}!";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            // Output each person's name and a greeting obtained via a static method.
            builder.Writeln("<<[p.Name]>> says: <<[Utils.GetGreeting(p.Name)]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Register the external type so its static members can be used safely.
            engine.KnownTypes.Add(typeof(Utils));

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
