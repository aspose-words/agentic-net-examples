using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        // Initialize properties to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Define the types that should be restricted in LINQ Reporting.
            //    This must be done before any ReportingEngine instance is used.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(
                typeof(System.Environment),   // Example of a prohibited type.
                typeof(System.IO.FileInfo)   // Another prohibited type.
            );

            // -----------------------------------------------------------------
            // 2. Create a template document programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            // Create a new blank document.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags that reference the data model.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template document back from disk (required by the workflow).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Prepare the data source.
            // -----------------------------------------------------------------
            Person person = new Person();

            // -----------------------------------------------------------------
            // 5. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // BuildReport overload that allows referencing the root object name ("person").
            engine.BuildReport(loadedTemplate, person, "person");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);

            // Indicate successful completion (no interactive prompts).
            Console.WriteLine($"Report generated and saved to '{Path.GetFullPath(outputPath)}'.");
        }
    }
}
