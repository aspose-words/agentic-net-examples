using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtml
{
    // Model class that provides the HTML snippet to be inserted.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string HtmlSnippet { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Define file paths in the current working directory.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath   = Path.Combine(Environment.CurrentDirectory, "ReportWithHtml.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Add a title.
            builder.Writeln("LINQ Reporting – HTML Insertion Example");
            builder.Writeln();

            // Insert the LINQ Reporting tag that will be replaced by the HTML snippet.
            // The block tag <<html [model.HtmlSnippet]>> tells the engine to evaluate the
            // expression and insert its value as HTML.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template so it can be loaded later for the report generation step.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Sample HTML that could also be read from an external file.
            string sampleHtml = @"
                <h2 style='color:#2E8B57;'>Welcome to the Report</h2>
                <p>This paragraph is <b>bold</b> and this one is <i>italic</i>.</p>
                <ul>
                    <li>First item</li>
                    <li>Second item</li>
                    <li>Third item</li>
                </ul>";

            ReportModel model = new ReportModel
            {
                HtmlSnippet = sampleHtml
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "model", so we pass it explicitly.
            engine.BuildReport(report, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the final document.
            // -----------------------------------------------------------------
            report.Save(outputPath);

            // Inform the user where the files are located (no interactive input required).
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report generated and saved to: {outputPath}");
        }
    }
}
