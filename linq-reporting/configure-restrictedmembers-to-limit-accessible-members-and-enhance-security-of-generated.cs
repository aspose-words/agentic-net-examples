using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace LinqReportingRestrictedMembersExample
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        public string Content { get; set; } = "This is a secure LINQ Reporting example.";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "ReportTemplate.docx";
            string outputPath = "ReportOutput.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Insert a content placeholder.
            builder.Writeln("<<[model.Content]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Note: In recent versions of Aspose.Words the RestrictedMembers
            // property has been removed. If needed, configure security via
            // other mechanisms such as sandboxed environments or custom
            // reflection handling. The example proceeds without explicit
            // RestrictedMembers configuration.

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            var model = new ReportModel();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
