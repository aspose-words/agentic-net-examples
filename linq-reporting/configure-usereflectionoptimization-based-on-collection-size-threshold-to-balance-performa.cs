using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data entity.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model that will be referenced from the template.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        // Threshold that decides whether to enable reflection optimization.
        private const int CollectionSizeThreshold = 5;

        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Persons.AddRange(new[]
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 },
                new Person { Name = "Diana", Age = 28 },
                new Person { Name = "Ethan", Age = 40 },
                new Person { Name = "Fiona", Age = 22 }
            });

            // Decide whether to use reflection optimization based on collection size.
            ReportingEngine.UseReflectionOptimization = model.Persons.Count > CollectionSizeThreshold;

            // Create a template document with LINQ Reporting tags.
            string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            Document doc = new Document(templatePath);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string reportPath = "Report.docx";
            doc.Save(reportPath);
        }

        // Generates a simple Word template containing a foreach loop over Persons.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Ensure the directory exists.
            string directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            doc.Save(filePath);
        }
    }
}
