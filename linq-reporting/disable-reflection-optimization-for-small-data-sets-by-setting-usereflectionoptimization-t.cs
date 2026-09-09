using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = "";
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
            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Create a template document with LINQ Reporting tags.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("People List:");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var doc = new Document(templatePath);

            // Disable reflection optimization for small data sets.
            ReportingEngine.UseReflectionOptimization = false;

            // Build the report using the model as the root data source.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
