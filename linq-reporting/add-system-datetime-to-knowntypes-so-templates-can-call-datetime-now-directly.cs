using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Path for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create a simple template that uses DateTime.Now.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for reporting.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Add System.DateTime to the set of known types so the template can access static members.
        engine.KnownTypes.Add(typeof(DateTime));

        // No data source is required for this example; an empty object is sufficient.
        engine.BuildReport(reportDoc, new object());

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive input required).
        Console.WriteLine($"Report generated successfully: {reportPath}");
    }
}
