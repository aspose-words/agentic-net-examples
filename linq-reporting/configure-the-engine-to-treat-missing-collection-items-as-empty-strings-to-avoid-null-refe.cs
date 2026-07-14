using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model for the report.
    public class ReportModel
    {
        // Collection of persons; one entry will be null to simulate a missing item.
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // LINQ Reporting tags:
            // Iterate over the Persons collection.
            // Use an if tag to guard against null items.
            builder.Writeln("<<foreach [person in Persons]>>");
            // The condition must return a Boolean value; check for null explicitly.
            builder.Writeln("<<if [person != null]>>Name: <<[person.Name]>> Age: <<[person.Age]>> <</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source with a missing collection item.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(null); // Missing item – should be treated as empty.
            model.Persons.Add(new Person { Name = "Bob", Age = 25 });

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine to treat missing members as empty.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Enable handling of missing members.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Use an empty placeholder for missing members.
                MissingMemberMessage = string.Empty
            };

            // Build the report. The root object name is "model" and must match the tags.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);

            Console.WriteLine("Report generated successfully at:");
            Console.WriteLine(reportPath);
        }
    }
}
