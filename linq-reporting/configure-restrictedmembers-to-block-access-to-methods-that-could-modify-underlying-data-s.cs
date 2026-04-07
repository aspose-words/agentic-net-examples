using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Non‑nullable property initialized to avoid warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that will be replaced with the model's Name.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure restricted members to block access to potentially dangerous types.
            //    This must be done before any call to BuildReport.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(
                typeof(System.IO.File),
                typeof(System.IO.Directory),
                typeof(System.Environment),
                typeof(System.Diagnostics.Process));

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so that the engine does not throw if a restricted type is referenced.
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "Access Denied"
            };

            // Create the root data object.
            ReportModel model = new ReportModel();

            // Populate the report. The root name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
