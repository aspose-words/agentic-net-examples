using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingErrorLogging
{
    // Simple data model used as the root object for the report.
    public class SampleModel
    {
        public string Name { get; set; } = "Test Name";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template, output document, and error log.
            string templatePath = "template.docx";
            string outputPath = "output.docx";
            string logPath = "error.log";

            // -----------------------------------------------------------------
            // 1. Create a template document with an invalid LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // This tag references a non‑existent member and will cause a syntax error.
            builder.Writeln("<<[model.NonExistentProperty]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Build the report without InlineErrorMessages.
            //    The engine will throw an exception on the syntax error.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // InlineErrorMessages disabled.

            try
            {
                // Attempt to build the report. This will fail because of the invalid tag.
                bool success = engine.BuildReport(doc, new SampleModel(), "model");

                // If, for any reason, the build succeeds, save the result.
                if (success)
                {
                    doc.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                // -----------------------------------------------------------------
                // 4. Log the syntax error details to a file.
                // -----------------------------------------------------------------
                File.WriteAllText(logPath, $"Report generation failed: {ex.Message}");
            }

            // Ensure the program finishes without waiting for user input.
        }
    }
}
