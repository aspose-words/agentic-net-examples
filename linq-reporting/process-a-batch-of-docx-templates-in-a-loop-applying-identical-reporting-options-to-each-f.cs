using System;
using System.Collections.Generic;
using System.IO;
using System.Text; // required for encoding registration
using Aspose.Words;
using Aspose.Words.Reporting; // for ReportingEngine and ReportBuildOptions

namespace AsposeWordsLinqReportingBatch
{
    // Simple data entity representing a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper model that contains the collection referenced by the template.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders for templates and generated reports.
            string templatesDir = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string reportsDir = Path.Combine(Directory.GetCurrentDirectory(), "Reports");
            Directory.CreateDirectory(templatesDir);
            Directory.CreateDirectory(reportsDir);

            // Create a single data model that will be used for all reports.
            ReportModel model = new()
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 45 },
                    new Person { Name = "Charlie", Age = 28 }
                }
            };

            // -----------------------------------------------------------------
            // Step 1: Create a few template files programmatically.
            // -----------------------------------------------------------------
            for (int i = 1; i <= 3; i++)
            {
                string templatePath = Path.Combine(templatesDir, $"Template{i}.docx");

                // Create a blank document and a builder to add content.
                Document templateDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(templateDoc);

                // Simple heading to identify the template.
                builder.Writeln($"Report Template {i}");

                // Insert LINQ Reporting tags.
                // The template expects a root object named "model" with a collection "Persons".
                builder.Writeln("<<foreach [p in Persons]>>");
                builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
                builder.Writeln("<</foreach>>");

                // Save the template to disk.
                templateDoc.Save(templatePath);
            }

            // -----------------------------------------------------------------
            // Step 2: Process each template, applying identical reporting options.
            // -----------------------------------------------------------------
            // Define the reporting engine once – it will be reused for every file.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Iterate over all .docx files in the Templates folder.
            foreach (string templateFile in Directory.GetFiles(templatesDir, "*.docx"))
            {
                // Load the template document.
                Document doc = new Document(templateFile);

                // Build the report using the same data model for every template.
                // The root name "model" must match the name used in the template tags.
                engine.BuildReport(doc, model, "model");

                // Save the generated report to the Reports folder.
                string reportFileName = $"Report_{Path.GetFileNameWithoutExtension(templateFile)}.docx";
                string reportPath = Path.Combine(reportsDir, reportFileName);
                doc.Save(reportPath);
            }

            // The example finishes execution without waiting for user input.
        }
    }
}
