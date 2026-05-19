using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Role of the current user (e.g., "Admin" or "User").
        public string Role { get; set; } = "Admin";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Paragraph visible only to administrators.
            builder.Writeln("<<if [model.Role == \"Admin\"]>>This section is visible to Admins.<</if>>");

            // Paragraph visible only to non‑administrators.
            builder.Writeln("<<if [model.Role != \"Admin\"]>>This section is visible to regular users.<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Prepare the data source.
            var model = new ReportModel
            {
                // Change this value to "User" to see the non‑admin content.
                Role = "Admin"
            };

            // Configure and execute the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // Default options.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
