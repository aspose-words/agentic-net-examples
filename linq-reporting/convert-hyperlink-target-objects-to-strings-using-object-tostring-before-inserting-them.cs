using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Hyperlink target as an object (Uri) that will be converted to string.
    public Uri LinkTarget { get; set; } = new Uri("https://example.com");

    // Display text for the hyperlink.
    public string LinkText { get; set; } = "Visit Example";

    // String representation of the target, obtained via Object.ToString().
    public string LinkTargetString => LinkTarget.ToString();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting link tag.
        // The first expression is the hyperlink target (converted to string),
        // the second expression is the display text.
        builder.Writeln("<<link [model.LinkTargetString] [model.LinkText]>>");

        // Save the template to disk (required by the workflow).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
