using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Hyperlink target as a Uri object.
    public Uri Target { get; set; } = new Uri("https://example.com");

    // Display text for the hyperlink.
    public string Text { get; set; } = "Visit Example";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -------------------------------------------------
        // 1. Create a Word template with a LINQ Reporting link tag.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // The link tag expects the URI (or bookmark) and the display text.
        // The URI is provided as an object (Uri) and converted to string using ToString().
        builder.Writeln("<<link [model.Target.ToString()] [model.Text]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var model = new ReportModel
        {
            // Example target; could be any object that overrides ToString().
            Target = new Uri("https://www.aspose.com"),
            Text = "Aspose.Words Documentation"
        };

        // -------------------------------------------------
        // 3. Build the report using Aspose.Words LINQ ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}
