using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document with a LINQ Reporting tag that calls DateTime.Now.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated at: <<[DateTime.Now]>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and configure the ReportingEngine.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Add System.DateTime to the set of known types so the template can access static members.
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report. No data source is required because the template only uses a static call.
        engine.BuildReport(loadedTemplate, new object());

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}
