using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Entry point of the example.
    public class ReportGenerator
    {
        public static void Main()
        {
            // Register code page provider for environments that require it.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a title.
            builder.Writeln("People Report");
            builder.Writeln();

            // Insert a foreach tag. The iteration variable is declared without an explicit type,
            // which is the supported syntax for Aspose.Words LINQ Reporting.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("- Name: <<[p.Name]>>");
            builder.Writeln("- Age : <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario).
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data model.
            // -------------------------------------------------
            ReportModel model = new()
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob",   Age = 45 },
                    new Person { Name = "Carol", Age = 27 }
                }
            };

            // -------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new();
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }

    // Root data model passed to the reporting engine.
    public class ReportModel
    {
        // Collection of Person objects to be iterated over in the template.
        public List<Person> Persons { get; set; } = new();
    }

    // Simple data entity with strongly typed properties.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
