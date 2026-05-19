using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        // Note: No property called 'Salary' – it will be referenced intentionally to cause an error.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write a simple greeting.
            builder.Writeln("Hello, <<[model.Name]>>!");

            // Conditional section that references a missing member 'Salary'.
            // The <<error>> tag will capture any evaluation failure when InlineErrorMessages is enabled.
            builder.Writeln("<<if [model.Salary > 5000]>>High salary<<error>><</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var model = new Person(); // Root object for the template.

            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages; // Enable inline error messages.

            // Build the report; the root object name must match the tag prefix used in the template.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(reportPath);

            // Output the result of the build operation.
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
