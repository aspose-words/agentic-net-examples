using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create a template document with a LINQ Reporting tag.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The tag <<[DateTime.UtcNow]>> will be replaced by the current UTC time.
        builder.Writeln("Current UTC time: <<[DateTime.UtcNow]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the DateTime type so that static members can be accessed in the template.
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report. No data source is required because we only use a static member.
        // Passing a dummy object satisfies the method signature.
        engine.BuildReport(loadedTemplate, new object());

        // -------------------------------------------------
        // 3. Save the generated report.
        // -------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}
