using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Bibliography;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "PersonTemplate.docx";
            const string reportPath = "PersonReport.docx";

            // -------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert placeholders that reference the root object named "person".
            builder.Writeln("First Name: <<[person.First]>>");
            builder.Writeln("Middle Name: <<[person.Middle]>>");
            builder.Writeln("Last Name: <<[person.Last]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data source – a single Person instance.
            // -------------------------------------------------
            // Person constructor: Person(string last, string first, string middle)
            Person person = new Person("Doe", "John", "M.");

            // -------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root name in the template is "person", so we pass it as the third argument.
            engine.BuildReport(reportDoc, person, "person");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
