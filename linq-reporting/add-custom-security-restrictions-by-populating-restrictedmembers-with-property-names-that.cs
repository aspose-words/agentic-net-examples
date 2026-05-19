using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with two properties.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public string Secret { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Secret: <<[person.Secret]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Restrict the Person type so its members cannot be accessed in the template.
            // This effectively hides the "Secret" property.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // Allow missing members to avoid exceptions when a restricted member is referenced.
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // -------------------------------------------------
            // 4. Build the report.
            // -------------------------------------------------
            Person data = new Person
            {
                Name = "John Doe",
                Secret = "TopSecret"
            };

            // The root object name used in the template is "person".
            engine.BuildReport(reportDoc, data, "person");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);

            // Output the resulting text to the console (for verification).
            Console.WriteLine("Report generated. Document text:");
            Console.WriteLine(reportDoc.GetText());
        }
    }
}
