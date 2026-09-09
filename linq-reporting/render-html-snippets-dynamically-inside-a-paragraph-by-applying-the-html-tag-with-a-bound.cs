using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class HtmlModel
{
    // HTML snippet that will be inserted into the paragraph.
    public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph that contains a LINQ Reporting HTML tag.
        // The tag will be replaced with the value of HtmlSnippet at build time.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data source.
        HtmlModel model = new HtmlModel();

        // Create the reporting engine and generate the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final document.
        loadedTemplate.Save(reportPath);
    }
}
