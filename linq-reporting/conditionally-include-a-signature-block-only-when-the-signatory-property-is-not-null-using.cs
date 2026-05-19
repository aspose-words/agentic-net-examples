using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSignatureExample
{
    // Data model for the report.
    public class ReportModel
    {
        // When null, the signature block will be omitted.
        public string? Signatory { get; set; } = null;
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

            // Simple title.
            builder.Writeln("Report Title");

            // Conditional signature block – appears only if Signatory is not null.
            // The <<if>> tag must contain a Boolean expression.
            builder.Writeln("<<if [model.Signatory != null]>>Signed by: <<[model.Signatory]>> <</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for reporting.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data where Signatory is provided.
            ReportModel model = new ReportModel
            {
                Signatory = "John Doe"
            };

            // -------------------------------------------------
            // 3. Build the report using LINQ Reporting Engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 4. Save the final report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
