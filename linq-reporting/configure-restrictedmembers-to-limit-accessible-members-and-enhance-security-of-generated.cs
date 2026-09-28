using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedMembersExample
{
    // Simple data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public string SecretInfo { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags.
            builder.Writeln("Report Title: <<[model.Title]>>");
            builder.Writeln("Secret Data: <<[model.SecretInfo]>>");

            // Save the template to disk.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Title = "Monthly Sales Summary",
                SecretInfo = "Confidential: Profit Margin 42%"
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            //    Note: The current Aspose.Words.Reporting API version does not expose
            //    a RestrictedMembers property. If it becomes available, you can set it
            //    here to block access to specific members (e.g., "SecretInfo").
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The engine will process the tags in the template.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
