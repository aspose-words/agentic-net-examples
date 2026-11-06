using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedMembers
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "Aspose.Words LINQ Reporting";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a Word document template programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the model.
            builder.Writeln("Report Title: <<[model.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            var model = new ReportModel();

            // -----------------------------------------------------------------
            // 4. Configure ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Set engine options. In this example no special flags are required,
            // but the property must be assigned according to the category rules.
            engine.Options = ReportBuildOptions.None;

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
