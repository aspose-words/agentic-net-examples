using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsRestrictedMembersDemo
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialise to avoid nullable warnings.
        public string Message { get; set; } = "Hello from the model!";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create a template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a simple placeholder that will be replaced by the model.
            builder.Writeln("Message: <<[model.Message]>>");

            // Attempt to call a file‑writing method – this should be blocked by the restricted types.
            builder.Writeln(@"Attempt to write file: <<[System.IO.File.WriteAllText(""test.txt"", ""Should be blocked"")]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back for reporting.
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Configure restricted types to prevent file‑system access from templates.
            // -------------------------------------------------
            // Restrict the types that expose file‑writing capabilities.
            ReportingEngine.SetRestrictedTypes(
                typeof(System.IO.File),
                typeof(System.IO.StreamWriter),
                typeof(System.IO.Directory));

            // -------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            var engine = new ReportingEngine();

            var model = new ReportModel();

            try
            {
                // The root object name ("model") must match the tag prefix used in the template.
                engine.BuildReport(doc, model, "model");
            }
            catch (Exception ex)
            {
                // If the template tries to use a restricted member, an exception will be thrown.
                // Write the error to the console and continue – the report will contain the parts that succeeded.
                Console.WriteLine($"Report generation error: {ex.Message}");
            }

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
