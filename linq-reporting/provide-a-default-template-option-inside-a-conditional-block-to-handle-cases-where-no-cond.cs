using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalDefault
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Status value that will be evaluated in the template.
        public string Status { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The template contains two IF conditions and a default IF that acts as ELSE.
            // If none of the IF conditions evaluate to true, the default IF provides a fallback value.
            builder.Writeln(
                "Status: " +
                "<<if [model.Status == \"A\"]>>A<</if>>" +
                "<<if [model.Status == \"B\"]>>B<</if>>" +
                // Default case when Status is neither A nor B.
                "<<if [model.Status != \"A\" && model.Status != \"B\"]>>Other<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data where Status does not match any IF condition (triggers default).
            ReportModel model = new ReportModel { Status = "C" };

            // Create the reporting engine and populate the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            reportDoc.Save(outputPath);
        }
    }
}
