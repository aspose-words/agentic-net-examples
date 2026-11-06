using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create the template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Start of document");
        builder.Writeln("<<if [model.ShowFirst]>>First block<</if>>");
        builder.Writeln("<<if [model.ShowSecond]>>Second block<</if>>");
        builder.Writeln("End of document");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model where all conditions are false.
        ReportModel model = new()
        {
            ShowFirst = false,
            ShowSecond = false
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    public bool ShowFirst { get; set; } = false;
    public bool ShowSecond { get; set; } = false;
}
