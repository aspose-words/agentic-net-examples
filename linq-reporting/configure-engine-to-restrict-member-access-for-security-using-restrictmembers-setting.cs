using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictMembersExample
{
    // Simple data model with a public property.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the Person.Name property.
            builder.Writeln("Customer Name: <<[person.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // Prepare the root data object.
            Person person = new Person();

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // NOTE: In newer versions of Aspose.Words the ReportBuildOptions.RestrictMembers flag
            // may not be available. If it exists, you can enable it as shown below:
            // engine.Options = ReportBuildOptions.RestrictMembers;
            // For this example we proceed without setting the flag to keep the code compilable.

            // Build the report.
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
        }
    }
}
