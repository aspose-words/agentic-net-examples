using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Public data model class.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        public string Email { get; set; } = "john.doe@example.com";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output documents.
            string templatePath = "PersonTemplate.docx";
            string outputPath = "PersonReport.docx";

            // -----------------------------------------------------------------
            // Step 1: Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Person Report");
            builder.Writeln("==============");
            builder.Writeln("Name : <<[person.Name]>>");
            builder.Writeln("Age  : <<[person.Age]>>");
            builder.Writeln("Email: <<[person.Email]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template document for report generation.
            // -----------------------------------------------------------------
            Document docForReport = new Document(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Prepare the data source (single Person object).
            // -----------------------------------------------------------------
            Person person = new Person();

            // -----------------------------------------------------------------
            // Step 4: Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(docForReport, person, "person");

            // -----------------------------------------------------------------
            // Step 5: Save the generated report.
            // -----------------------------------------------------------------
            docForReport.Save(outputPath);
        }
    }
}
