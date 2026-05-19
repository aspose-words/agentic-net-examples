using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = "";
        public string Title { get; set; } = "";
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure code page support (required for some environments)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Create the template document programmatically
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Employee List:");
            // Start foreach loop over Persons collection
            builder.Writeln("<<foreach [p in Persons]>>");
            // Each iteration writes a paragraph that may become empty if both fields are empty
            builder.Writeln("<<[p.Name]>> <<[p.Title]>>");
            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template (optional, shown for clarity)
            string templatePath = Path.Combine("Output", "Template.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(templatePath)!);
            template.Save(templatePath);

            // Prepare sample data with one completely empty entry
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "John Doe", Title = "Manager" },
                    new Person { Name = "Jane Smith", Title = "Developer" },
                    new Person { Name = "", Title = "" } // This will produce an empty paragraph
                }
            };

            // Load the template (demonstrates load step)
            Document doc = new Document(templatePath);

            // Configure the reporting engine to remove empty paragraphs after processing
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report; root object name must match the name used in the template tags
            engine.BuildReport(doc, model, "model");

            // Save the final document
            string outputPath = Path.Combine("Output", "Report.docx");
            doc.Save(outputPath);
        }
    }
}
