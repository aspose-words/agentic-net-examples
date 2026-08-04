using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingSecurity
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = "";
        public decimal Salary { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Salary: <<[person.Salary]>>");

            // Prepare the data source.
            Person person = new Person
            {
                Name = "John Doe",
                Salary = 12345.67m
            };

            // Restrict access to the Person type. All its members will be treated as missing.
            // This is the supported way to enforce security restrictions in Aspose.Words LINQ Reporting.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "[Hidden]"
            };

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
