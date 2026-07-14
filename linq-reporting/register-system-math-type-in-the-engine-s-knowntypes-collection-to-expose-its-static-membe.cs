using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define paths for the template and the final report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // -------------------------------------------------
        // 1. Create a simple template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that accesses a static member of System.Math.
        // The ReportingEngine will resolve Math.PI because we will register System.Math in KnownTypes.
        builder.Writeln("Value of PI: <<[Math.PI]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and configure the ReportingEngine.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Register System.Math so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report. No data source is required for this example.
        engine.BuildReport(loadedTemplate, new object());

        // -------------------------------------------------
        // 3. Save the generated report.
        // -------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}
