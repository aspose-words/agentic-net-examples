using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Collection of persons – fewer than 1000 items.
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
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a temporary file (optional but follows the create‑save‑load pattern).
            string templatePath = "ReportTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template back.
            Document loadedTemplate = new Document(templatePath);

            // 3. Prepare sample data (less than 1000 records).
            ReportModel model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(new Person { Name = "Bob", Age = 45 });
            model.Persons.Add(new Person { Name = "Charlie", Age = 25 });

            // 4. Disable reflection optimization for small data sets.
            ReportingEngine.UseReflectionOptimization = false;

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this scenario.
            engine.BuildReport(loadedTemplate, model, "model");

            // 6. Save the generated report.
            string outputPath = "ReportResult.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
