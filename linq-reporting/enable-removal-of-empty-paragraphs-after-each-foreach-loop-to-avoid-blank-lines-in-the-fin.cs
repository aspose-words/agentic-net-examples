using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingEmptyParagraphRemoval
{
    // Sample data model.
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
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create the template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Persons List:");
            // LINQ Reporting foreach tag.
            builder.Writeln("<<foreach [p in Persons]>>");
            // Content inside the loop.
            builder.Writeln("Name: <<[p.Name]>> Age: <<[p.Age]>>");
            // End of foreach.
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // Prepare sample data.
            ReportModel model = new()
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 40 }
                }
            };

            // Build the report using LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Remove empty paragraphs that may have been left after foreach processing.
            RemoveEmptyParagraphs(reportDoc);

            // Save the final report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(reportPath);
        }

        // Helper method to delete paragraphs that contain no visible text.
        private static void RemoveEmptyParagraphs(Document doc)
        {
            // Collect paragraphs to remove to avoid modifying the collection while iterating.
            List<Paragraph> emptyParagraphs = new();

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Trim whitespace and line breaks; an empty paragraph returns just a paragraph break.
                if (string.IsNullOrWhiteSpace(para.GetText()))
                {
                    emptyParagraphs.Add(para);
                }
            }

            foreach (Paragraph para in emptyParagraphs)
            {
                para.Remove();
            }
        }
    }
}
