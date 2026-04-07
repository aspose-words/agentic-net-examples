using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    // Simple data model used as the root object for the report.
    public class SampleModel
    {
        public string Title { get; set; } = "Report Title";
        // No property named MissingObject – this will cause a syntax error in the template.
    }

    public static void Main()
    {
        // Paths for the template, generated report and error log.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";
        const string logPath = "error.log";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Valid tag – will be replaced correctly.
        builder.Writeln("<<[model.Title]>>");

        // Invalid tag – references a non‑existent member, causing a syntax error.
        builder.Writeln("<<[model.MissingObject.Property]>>");

        // Save the template to disk (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real scenario where the
        //    template might be stored externally).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        SampleModel model = new SampleModel();

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine without InlineErrorMessages.
        //    This means syntax errors will throw an exception.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // InlineErrorMessages disabled

        try
        {
            // Attempt to build the report. The overload with the data source name
            // allows the template to reference the root object as "model".
            engine.BuildReport(doc, model, "model");

            // If no exception occurs, save the generated report.
            doc.Save(outputPath);
        }
        catch (Exception ex)
        {
            // -----------------------------------------------------------------
            // 5. Log the syntax error to a file.
            // -----------------------------------------------------------------
            File.WriteAllText(logPath, $"Report generation failed: {ex.Message}");
        }
    }
}
