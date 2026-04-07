using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model for a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public int AgeThreshold { get; set; }
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Report of persons older than <<[model.AgeThreshold]>> years:");
            builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > AgeThreshold)]>>");
            builder.Writeln("- <<[p.Name]>> (<<[p.Age]>> years)");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare sample data.
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                AgeThreshold = 30,
                Persons = new List<Person>
                {
                    new Person { Name = "Alice",   Age = 25 },
                    new Person { Name = "Bob",     Age = 35 },
                    new Person { Name = "Charlie", Age = 40 }
                }
            };

            // -------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(reportPath);

            // Inform the user where the report was saved.
            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}
