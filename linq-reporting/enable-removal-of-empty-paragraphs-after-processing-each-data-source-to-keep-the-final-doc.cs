using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingRemoveEmptyParagraphs
{
    // Sample data model
    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the LINQ Reporting template programmatically
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title paragraph
            builder.Writeln("Person Report");
            builder.Writeln(); // empty paragraph for visual separation

            // Begin foreach loop over Persons collection
            builder.Writeln("<<foreach [p in Persons]>>");
            // Write the person's name; if Name is empty, this will produce an empty paragraph
            builder.Writeln("<<[p.Name]>>");
            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // 2. Load the template for report generation
            Document reportDoc = new Document(templatePath);
            DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc); // not strictly needed but kept for consistency

            // 3. Prepare sample data with some empty names to generate empty paragraphs
            Model model = new()
            {
                Persons = new()
                {
                    new Person { Name = "Alice" },
                    new Person { Name = "" },          // This will generate an empty paragraph
                    new Person { Name = "Bob" },
                    new Person { Name = null! },       // Null treated as empty string after processing
                    new Person { Name = "Charlie" }
                }
            };

            // 4. Build the report using LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. Remove empty paragraphs that may have been created during report generation
            // Collect paragraphs to remove to avoid modifying the collection while iterating
            List<Paragraph> emptyParagraphs = new();
            foreach (Paragraph para in reportDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                // GetText includes paragraph break; Trim removes whitespace and line breaks
                if (string.IsNullOrWhiteSpace(para.GetText()))
                {
                    emptyParagraphs.Add(para);
                }
            }

            foreach (Paragraph para in emptyParagraphs)
            {
                para.Remove();
            }

            // 6. Save the final cleaned report
            string reportPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(reportPath);
        }
    }
}
