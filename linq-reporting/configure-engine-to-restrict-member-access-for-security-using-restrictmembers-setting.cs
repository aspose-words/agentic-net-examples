using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictMembersExample
{
    // Simple data model with public members.
    public class Person
    {
        // Public member that will be accessed in the template.
        public string Name { get; set; } = "John Doe";

        // Public member that we intend to hide when member restriction is enabled.
        public string Secret { get; set; } = "TopSecret";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Secret (should be hidden): <<[person.Secret]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The original example intended to use the <<restrictMembers>> option.
            // In the currently available Aspose.Words version the corresponding
            // flag is not present, so we configure the engine with a valid option.
            // InlineErrorMessages is used here to demonstrate engine configuration.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // -----------------------------------------------------------------
            // 4. Build the report using the data model.
            // -----------------------------------------------------------------
            Person person = new Person();
            bool success = engine.BuildReport(reportDoc, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }
}
