using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // ReportingEngine namespace
using System.Text; // For encoding provider

namespace AsposeWordsLinqReportingHtml
{
    // Data model representing a record fetched from a database.
    public class ReportModel
    {
        // HTML content stored in a database field.
        public string HtmlContent { get; set; } = "<p style='color:blue;'>This is <b>HTML</b> content from a database field.</p>";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document with a LINQ Reporting HTML tag.
            // -----------------------------------------------------------------
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple title.
            builder.Writeln("Report generated with Aspose.Words LINQ Reporting");
            builder.Writeln(); // Empty line.

            // Insert the HTML expression tag. The '-html' switch tells the engine to treat the value as HTML.
            builder.Writeln("<<[model.HtmlContent] -html>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source (simulating a DB record).
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.

            // The root object name in the template is 'model', so we pass it explicitly.
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Optional: you could check 'success' if InlineErrorMessages were enabled.
            // -----------------------------------------------------------------
            // 5. Save the final document.
            // -----------------------------------------------------------------
            string outputPath = "ReportOutput.docx";
            reportDoc.Save(outputPath);

            // The program finishes without waiting for user input.
        }
    }
}
