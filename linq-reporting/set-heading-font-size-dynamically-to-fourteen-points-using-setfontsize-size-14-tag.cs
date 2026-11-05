using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Heading HTML with dynamic font size of 14 points.
    public string Title { get; set; } = "<span style=\"font-size:14pt;\">Dynamic Report Title</span>";
}

public class Program
{
    public static void Main()
    {
        // Prepare paths.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // Create the template document.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a heading placeholder that will render HTML.
        builder.Writeln("<<[model.Title] -html>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
