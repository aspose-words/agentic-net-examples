using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string ValidProperty { get; set; } = "Sample text";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string outputPath = "ReportOutput.docx";

        // -------------------------------------------------
        // 1. Create a template document with a syntax error.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // A correct tag – will be replaced with the model value.
        builder.Writeln("<<[model.ValidProperty]>>");

        // An intentionally malformed tag to produce a syntax error.
        builder.Writeln("<<[model.MissingTag>>"); // Missing closing bracket.

        // Save the template so that it is fully created before building the report.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back (required by the rule).
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data source.
        // -------------------------------------------------
        ReportModel model = new ReportModel();

        // -------------------------------------------------
        // 4. Configure the ReportingEngine with InlineErrorMessages.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // Output the success flag.
        Console.WriteLine($"BuildReport success: {success}");

        // Save the generated report (it will contain inline error messages).
        loadedTemplate.Save(outputPath);
    }
}
