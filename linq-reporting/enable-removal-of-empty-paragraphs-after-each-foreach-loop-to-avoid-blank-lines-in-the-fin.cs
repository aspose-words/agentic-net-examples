using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRemoveEmptyParagraphs
{
    // Sample data model
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
            // Register code page provider (required for Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create template document programmatically
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Person Report");
            builder.Writeln(); // Empty line before the loop

            // Start foreach loop
            builder.Writeln("<<foreach [p in Persons]>>");
            // Paragraph inside loop
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Add an explicit empty paragraph after the loop (to be removed)
            builder.Writeln();

            // Save the template
            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // Load the template for reporting
            Document reportDoc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new()
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Configure reporting engine to remove empty paragraphs
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document
            string outputPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
