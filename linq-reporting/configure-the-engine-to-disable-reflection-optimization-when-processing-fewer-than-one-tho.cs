using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper class that holds a collection of Person objects.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data with fewer than 1,000 records.
            var model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(new Person { Name = "Bob", Age = 45 });
            model.Persons.Add(new Person { Name = "Charlie", Age = 25 });

            // Create a temporary folder for the template and output files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);
            string templatePath = Path.Combine(workDir, "Template.docx");
            string outputPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // Step 1: Build the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a heading.
            builder.Writeln("Person Report");
            builder.Writeln();

            // Insert a foreach block that iterates over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and configure the ReportingEngine.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Disable reflection optimization because the data set is small (< 1000 records).
            ReportingEngine.UseReflectionOptimization = false;

            var engine = new ReportingEngine();

            // Build the report using the model as the data source.
            // Using the overload without a root name allows direct access to model members.
            engine.BuildReport(doc, model);

            // Save the generated report.
            doc.Save(outputPath);

            // Inform the user where the files are located (optional, not interactive).
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report generated at: {outputPath}");
        }
    }
}
