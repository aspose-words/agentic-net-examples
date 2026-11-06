using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRemoveEmptyParagraphs
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Create template document programmatically
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Persons Report");
            builder.Writeln("<<foreach [person in Persons]>>");
            // Paragraph that may become empty if Name is empty
            builder.Writeln("<<if [person.Name != \"\"]>><<[person.Name]>> - <<[person.Age]>> <</if>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for reporting
            Document reportDoc = new Document(templatePath);

            // Sample data with an empty Name to produce an empty paragraph
            ReportModel model = new()
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "", Age = 25 }, // This will generate an empty paragraph
                    new Person { Name = "Bob", Age = 40 }
                }
            };

            // Configure ReportingEngine to remove empty paragraphs
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document
            string resultPath = Path.Combine(outputDir, "result.docx");
            reportDoc.Save(resultPath);
        }
    }
}
