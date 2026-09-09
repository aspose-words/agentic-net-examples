using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper class that holds the collection referenced by the template.
    public class DataRoot
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var data = new DataRoot();
            for (int i = 1; i <= 5; i++) // fewer than 1000 records
                data.Persons.Add(new Person { Name = $"Person {i}" });

            // Create a template document programmatically.
            string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            Document doc = new Document(templatePath);

            // Disable reflection optimization when the record count is less than 1000.
            ReportingEngine.UseReflectionOptimization = data.Persons.Count >= 1000;

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data");

            // Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }

        // Creates a simple LINQ Reporting template with a foreach loop.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Template tags.
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
