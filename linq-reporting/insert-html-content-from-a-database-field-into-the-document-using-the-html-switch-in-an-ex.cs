using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class representing data that would normally come from a database.
    public class ReportModel
    {
        // HTML fragment stored in a database field.
        public string HtmlContent { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document with a LINQ Reporting HTML tag.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report generated with HTML content from a data source:");
            // The -html switch tells the engine to treat the expression result as HTML.
            builder.Writeln("<<[model.HtmlContent] -html>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            var model = new ReportModel
            {
                HtmlContent = "<p style='color:blue;'>This is <b>HTML</b> inserted from a database field.</p>"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // The root object name in the template is "model", matching the third argument.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the final document.
            // -----------------------------------------------------------------
            var outputPath = "ReportOutput.docx";
            reportDoc.Save(outputPath);
        }
    }
}
