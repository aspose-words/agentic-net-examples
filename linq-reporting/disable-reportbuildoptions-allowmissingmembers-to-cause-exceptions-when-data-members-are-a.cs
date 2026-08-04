using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a single property.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        // Note: No Age property – this will trigger a missing‑member error.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document that contains a tag referencing a
            //    non‑existent member (Age). The tag syntax follows Aspose.Words
            //    LINQ Reporting rules.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>"); // Age does not exist.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back from disk (simulating a real‑world scenario).
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var person = new Person();

            // -----------------------------------------------------------------
            // 4. Build the report without enabling AllowMissingMembers.
            //    This should cause an exception because the template references
            //    a missing member (Age).
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            try
            {
                // The root object name used in the template tags is "person".
                engine.BuildReport(doc, person, "person");
                // If no exception occurs, save the generated report.
                doc.Save(reportPath);
                Console.WriteLine("Report generated successfully (unexpected).");
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
