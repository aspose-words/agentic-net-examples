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
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            doc.Save(templatePath);

            // Load the template (demonstrates separate load step).
            var template = new Document(templatePath);

            // Disable reflection optimization for this report generation.
            bool previousOptimization = ReportingEngine.UseReflectionOptimization;
            ReportingEngine.UseReflectionOptimization = false;
            try
            {
                var engine = new ReportingEngine();
                // Build the report using the model; root name must match the template tags.
                engine.BuildReport(template, model, "model");
                // Save the generated report.
                template.Save("Report.docx");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = previousOptimization;
            }
        }
    }
}
