using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Data model classes
    public class ReportModel
    {
        public int MinAge { get; set; } = 30;
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
            // Register code page provider (required for some Aspose.Words features)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document with LINQ Reporting tags
            // -------------------------------------------------
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Title
            builder.Writeln("Filtered Persons (Age >= <<[model.MinAge]>>):");

            // Foreach with Where filter based on runtime MinAge
            builder.Writeln("<<foreach [p in model.Persons.Where(p => p.Age >= model.MinAge)]>>");
            builder.Writeln("- <<[p.Name]>> (<<[p.Age]>> years)");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template document (required before building report)
            // -------------------------------------------------
            Document loadedTemplate = new(templatePath);

            // -------------------------------------------------
            // Prepare sample data
            // -------------------------------------------------
            ReportModel model = new()
            {
                MinAge = 30,
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 28 },
                    new Person { Name = "Bob", Age = 35 },
                    new Person { Name = "Charlie", Age = 42 },
                    new Person { Name = "Diana", Age = 25 }
                }
            };

            // -------------------------------------------------
            // Build the report using Aspose.Words LINQ Reporting Engine
            // -------------------------------------------------
            ReportingEngine engine = new();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report
            loadedTemplate.Save(outputPath);
        }
    }
}
