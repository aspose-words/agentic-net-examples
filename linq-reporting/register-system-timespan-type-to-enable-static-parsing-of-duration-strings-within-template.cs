using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a simple template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // The template uses a static call to TimeSpan.Parse to convert a duration string.
        // Use double quotes inside the expression to avoid parsing errors.
        builder.Writeln("Task duration: <<[TimeSpan.Parse(\"01:30:00\")]>>");

        // Save the template to disk.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (required before building the report).
        Document loadedTemplate = new Document(templatePath);

        // Create the reporting engine and register System.TimeSpan for static parsing.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. No data source is needed because the template only uses static parsing.
        // Pass a dummy object as the data source.
        engine.BuildReport(loadedTemplate, new object(), "model");

        // Save the generated report.
        string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
