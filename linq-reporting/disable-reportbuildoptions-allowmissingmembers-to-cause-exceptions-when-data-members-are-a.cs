using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a missing member (Age) that the template will try to access.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        // Note: No Age property – this will cause the reporting engine to throw.
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Path for the template document.
            string templatePath = Path.Combine(outputDir, "Template.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a tag that references a non‑existent member (Age).
            builder.Writeln("Person age: <<[person.Age]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source (a Person instance without Age).
            // -----------------------------------------------------------------
            Person person = new Person();

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine without AllowMissingMembers.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Do NOT set ReportBuildOptions.AllowMissingMembers.
                Options = ReportBuildOptions.None
            };

            try
            {
                // Attempt to build the report. This should throw because Age is missing.
                engine.BuildReport(doc, person, "person");
                // If no exception, save the result (unlikely in this scenario).
                string resultPath = Path.Combine(outputDir, "Result.docx");
                doc.Save(resultPath);
                Console.WriteLine("Report generated successfully: " + resultPath);
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown due to the missing member.
                Console.WriteLine("Expected exception caught:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
