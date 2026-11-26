using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingFallbackExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Non‑nullable property – initialized to avoid warnings.
        public string Name { get; set; } = "John Doe";

        // Nullable property that may be null.
        public string? Notes { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string outputPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[model.Name]>>");

            // Expression that may be null – use conditional tags to provide a fallback.
            builder.Writeln("Notes: <<if [model.Notes != null]>><<[model.Notes]>><</if>>");
            builder.Writeln("<<if [model.Notes == null]>>No notes provided.<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data model (Notes is null to trigger fallback).
            // -------------------------------------------------
            var model = new ReportModel
            {
                // Name is already set by default.
                Notes = null // Simulate missing data.
            };

            // -------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
