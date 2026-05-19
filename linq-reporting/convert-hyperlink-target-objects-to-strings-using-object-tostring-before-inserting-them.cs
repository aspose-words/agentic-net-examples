using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Hyperlink target object – a Uri instance.
    public Uri UrlObject { get; set; } = new Uri("https://www.example.com");

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = "Example Site";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Visit the link below:");
        // The link tag expects a URI (or bookmark) and display text.
        // Convert the Uri object to string using Object.ToString().
        builder.Writeln("<<link [model.UrlObject.ToString()] [model.LinkText]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
