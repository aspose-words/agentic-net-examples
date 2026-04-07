using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Text that will be displayed in the report.
        public string Status { get; set; } = string.Empty;

        // Name or HTML color code for the text color.
        public string Color { get; set; } = string.Empty;

        // Name or HTML color code for the background color.
        public string BackgroundColor { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders for the temporary template and final HTML output.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);
            string templatePath = Path.Combine(workDir, "template.docx");
            string htmlOutputPath = Path.Combine(workDir, "Report.html");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple heading.
            builder.Writeln("LINQ Reporting – Dynamic Colors Demo");

            // Text with a dynamic foreground color.
            // The <<textColor>> tag will apply the color returned by model.Color.
            builder.Writeln("<<textColor [model.Color]>>Status: <<[model.Status]>> <</textColor>>");

            // Text with a dynamic background color.
            // The <<backColor>> tag will apply the background color returned by model.BackgroundColor.
            builder.Writeln("<<backColor [model.BackgroundColor]>>This line has a dynamic background.<</backColor>>");

            // Save the template to disk so that it can be loaded later (required by the rule set).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and populate it with data using ReportingEngine.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data for the report.
            ReportModel model = new ReportModel
            {
                Status = "Active",
                Color = "Green",               // Can also be a hex code like "#008000".
                BackgroundColor = "#FFDDDD"    // Light pink background.
            };

            // Build the report. The root object name in the template is "model".
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // Default options.
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Optional: verify that the report was built successfully.
            if (!success)
            {
                Console.WriteLine("Report building failed.");
                return;
            }

            // -----------------------------------------------------------------
            // 3. Save the populated document to HTML, preserving the dynamic colors.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // The default options already preserve text and background colors.
                // Additional settings can be configured here if needed.
                ExportTextInputFormFieldAsText = true
            };

            reportDoc.Save(htmlOutputPath, htmlOptions);

            Console.WriteLine($"HTML report generated at: {htmlOutputPath}");
        }
    }
}
