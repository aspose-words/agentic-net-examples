using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Bibliography;

namespace AsposeWordsLinqReportingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Define file names for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert placeholders that reference a root object named "person".
            builder.Writeln("First Name: <<[person.First]>>");
            builder.Writeln("Middle Name: <<[person.Middle]>>");
            builder.Writeln("Last Name: <<[person.Last]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario where the
            //    template exists as an external file).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model – a single Person object.
            // -----------------------------------------------------------------
            // The Person class resides in Aspose.Words.Bibliography.
            // Constructor: Person(string last, string first, string middle)
            Person person = new Person("Doe", "John", "A.");

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // Explicit assignment as required.

            // The root object name in the template is "person", so we pass that name.
            engine.BuildReport(loadedTemplate, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
