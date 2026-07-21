using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // External type whose static property will be used inside the LINQ expression.
    public static class ExternalHelper
    {
        // Minimum age used for filtering.
        public static int MinAge { get; } = 30;
    }

    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper class passed as the root data source.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 25 },
                    new Person { Name = "Bob",   Age = 35 },
                    new Person { Name = "Carol", Age = 45 }
                }
            };

            // 2. Create the template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Filtered persons (Age > ExternalHelper.MinAge):");
            // LINQ Reporting tag using Where with a lambda that references ExternalHelper.MinAge.
            builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > ExternalHelper.MinAge)]>>");
            builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // 3. Load the template and build the report.
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Register the external type so its static members can be accessed in the template.
            engine.KnownTypes.Add(typeof(ExternalHelper));

            // Build the report using the model as the root object named "model".
            engine.BuildReport(loadedDoc, model, "model");

            // 4. Save the generated report.
            loadedDoc.Save("Report.docx");
        }
    }
}
