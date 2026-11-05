using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // This property exists, but the template will request a missing one.
    public string ExistingProperty { get; set; } = "Existing Value";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with a tag that references a missing member.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The tag <<[model.MissingProperty]>> does NOT exist in the Model class.
        builder.Writeln("<<[model.MissingProperty]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back for reporting.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data model (does NOT contain MissingProperty).
        Model model = new Model();

        // -----------------------------------------------------------------
        // 3. Configure ReportingEngine without AllowMissingMembers.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Disables AllowMissingMembers.

        try
        {
            // Attempt to build the report. This should throw because the member is missing.
            engine.BuildReport(loadedTemplate, model, "model");

            // If no exception, save the generated report.
            loadedTemplate.Save(reportPath);
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
        catch (Exception ex)
        {
            // Expected path: an exception indicating the missing member.
            Console.WriteLine("An exception occurred during report generation:");
            Console.WriteLine(ex.Message);
        }
    }
}
