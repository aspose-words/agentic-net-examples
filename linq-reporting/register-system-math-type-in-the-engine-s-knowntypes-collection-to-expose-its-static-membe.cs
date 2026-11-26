using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with a LINQ Reporting tag that uses System.Math
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("The sine of 0.5 is <<[Math.Sin(0.5)]>>.");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Create the reporting engine and register System.Math type
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(System.Math));

        // Build the report (no root data needed, but a wrapper object is required)
        EmptyModel model = new EmptyModel();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Empty wrapper class required by the reporting engine
public class EmptyModel
{
    // No members needed; the template only uses static Math members
}
