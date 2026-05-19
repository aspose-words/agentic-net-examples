using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with a foreach tag that iterates
            //    over the Persons collection of the root object "model".
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct foreach syntax: no type declaration, use the root object name.
            builder.Writeln("<<foreach [p in model.Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template before building the report
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required by the workflow rules)
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed

            // The root object name in the template is "model"
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}
