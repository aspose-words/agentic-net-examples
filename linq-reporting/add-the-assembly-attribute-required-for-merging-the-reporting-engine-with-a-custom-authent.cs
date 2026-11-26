using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample property referenced from the template.
    public string Message { get; set; } = "Hello from LINQ Reporting!";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // Create a simple Word template containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report:");
        builder.Writeln("<<[model.Message]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template back for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare the root data object that matches the tag in the template.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel();

        // -----------------------------------------------------------------
        // Build the report using the Aspose.Words LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
