using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Hyperlink target as an object (e.g., Uri). Initialized to avoid nullable warnings.
    public Uri LinkTarget { get; set; } = new Uri("https://www.example.com");
    // Display text for the hyperlink.
    public string LinkText { get; set; } = "Visit Example";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting link tag.
        // The first expression converts the Uri object to a string via ToString().
        builder.Writeln("<<link [model.LinkTarget.ToString()] [model.LinkText]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using the root name "model" as referenced in the template.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report.
        loadedTemplate.Save(reportPath);
    }
}
