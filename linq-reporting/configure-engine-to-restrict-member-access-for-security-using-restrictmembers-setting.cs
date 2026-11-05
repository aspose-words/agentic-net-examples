using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictMembersExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Public property accessed from the template.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create a simple template document containing a LINQ Reporting tag.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello, <<[model.Name]>>!");
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Sample data model.
            var model = new ReportModel();

            // -----------------------------------------------------------------
            // Configure the reporting engine to restrict member access for security.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // The RestrictMembers flag is not available in the referenced version,
            // so we set the underlying enum value directly (value 2 corresponds to RestrictMembers).
            engine.Options = (ReportBuildOptions)2;

            // Build the report using the model as the root object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}
