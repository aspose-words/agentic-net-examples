using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Url = "https://www.example.com",
            LinkText = "Visit Example"
        };

        // Create a template document with a LINQ Reporting link tag.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // Build the report using the model as the root object named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save("Report.docx");
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}
