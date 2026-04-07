using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingInlineError
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
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
            // Paths for the template and the generated report.
            string templatePath = "template.docx";
            string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin a foreach loop over the collection "Persons".
            builder.Writeln("<<foreach [p in Persons]>>");

            // Correct property.
            builder.Writeln("Name: <<[p.Name]>>");

            // This property does NOT exist on Person and will cause a syntax error.
            // The <<error>> placeholder will be replaced with the inline error message.
            builder.Writeln("Age: <<[p.Age]>> <<error>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "John Doe" },
                    new Person { Name = "Jane Smith" }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report with inline error messages enabled.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // The boolean indicates whether the template was parsed without fatal errors.
            bool success = engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Output simple status information (no interactive prompts).
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
        }
    }
}
