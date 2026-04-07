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
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "";
    }

    // Wrapper model that the template will reference.
    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Set the reflection optimization globally.
            ReportingEngine.UseReflectionOptimization = true;

            // Paths for the template and generated reports.
            string templatePath = "Template.docx";
            string largeReportPath = "ReportLarge.docx";
            string smallReportPath = "ReportSmall.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple heading.
            builder.Writeln("People Report");
            builder.Writeln();

            // LINQ Reporting tags: iterate over Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before BuildReport).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Build a report using a large data source (global optimization stays true).
            // -----------------------------------------------------------------
            Document largeTemplate = new Document(templatePath);

            // Populate a large collection of persons.
            Model largeModel = new Model();
            for (int i = 1; i <= 100; i++)
            {
                largeModel.Persons.Add(new Person { Name = $"Person {i}" });
            }

            // Build the report.
            ReportingEngine largeEngine = new ReportingEngine();
            largeEngine.BuildReport(largeTemplate, largeModel, "model");
            largeTemplate.Save(largeReportPath);

            // -----------------------------------------------------------------
            // Override the reflection optimization for a small data source.
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = false;

            Document smallTemplate = new Document(templatePath);

            // Small collection of persons.
            Model smallModel = new Model
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice" },
                    new Person { Name = "Bob" }
                }
            };

            // Build the report for the small data source.
            ReportingEngine smallEngine = new ReportingEngine();
            smallEngine.BuildReport(smallTemplate, smallModel, "model");
            smallTemplate.Save(smallReportPath);

            // Reset the global setting if further processing is needed.
            ReportingEngine.UseReflectionOptimization = true;
        }
    }
}
