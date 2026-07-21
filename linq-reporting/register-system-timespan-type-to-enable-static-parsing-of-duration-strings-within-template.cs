using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that use static parsing of a TimeSpan string.
        // Use double quotes inside the expression to avoid char literal parsing errors.
        builder.Writeln("Parsed duration: <<[TimeSpan.Parse(\"02:15:30\")]>>");
        builder.Writeln("Total minutes: <<[TimeSpan.Parse(\"02:15:30\").TotalMinutes]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Create the reporting engine and register System.TimeSpan for static method access.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. No data source is required because the template uses only static calls.
        engine.BuildReport(report, new object());

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
