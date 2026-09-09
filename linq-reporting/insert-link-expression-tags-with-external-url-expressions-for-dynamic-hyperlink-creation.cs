using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinkModel
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // 1. Create the template document with a LINQ Reporting link tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Visit our site: <<link [model.Url] [model.LinkText]>>");
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare the data model.
        LinkModel model = new LinkModel
        {
            Url = "https://www.example.com",
            LinkText = "Example Website"
        };

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // 5. Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}
