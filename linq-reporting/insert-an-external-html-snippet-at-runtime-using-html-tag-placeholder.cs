using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class used as the data source for the LINQ Reporting engine.
    public class ReportModel
    {
        // HTML snippet that will be inserted into the document at runtime.
        public string HtmlSnippet { get; set; } = "<p style=\"color:blue;\">This is <b>dynamic</b> HTML content.</p>";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title.
            builder.Writeln("LINQ Reporting – HTML Insertion Example");
            builder.Writeln();

            // Insert the <<html>> tag placeholder that will be replaced at runtime.
            // The tag references the HtmlSnippet property of the model object.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required before building the report).
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // No special options are needed for this simple scenario.
                Options = ReportBuildOptions.None
            };

            // The root object name must match the tag reference ("model").
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Optional: you could check the success flag if InlineErrorMessages were enabled.
            // For this example we simply proceed.

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
