using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Result.docx");

            // 1. Create the template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the collection "Persons".
            builder.Writeln("<<foreach [p in Persons]>>");

            // Write a line that will always have content.
            builder.Writeln("Name: <<[p.Name]>>");

            // Conditional block: will be empty for persons aged 30 or less.
            builder.Writeln("<<if [p.Age > 30]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("<</if>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 2. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 3. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 28 },   // Age <= 30, conditional block empty.
                    new Person { Name = "Bob",   Age = 35 },   // Age > 30, conditional block will have content.
                    new Person { Name = "Carol", Age = 22 }    // Age <= 30, conditional block empty.
                }
            };

            // 4. Configure the reporting engine to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // 5. Build the report.
            engine.BuildReport(reportDoc, model, "model");

            // 6. Save the final document.
            reportDoc.Save(resultPath);
        }
    }
}
