using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class HtmlSnippetModel
{
    // Sample HTML that will be rendered inside the paragraph.
    public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with a LINQ Reporting tag that renders HTML.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Report generated with dynamic HTML:");
        // The -html switch tells the engine to treat the expression result as HTML.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template (simulating a separate load step).
        var loadedTemplate = new Document(templatePath);

        // 3. Prepare the data model.
        var model = new HtmlSnippetModel();

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // The root object name must match the tag reference (model).
        engine.BuildReport(loadedTemplate, model, "model");

        // 5. Save the generated report.
        const string outputPath = "ReportWithHtml.docx";
        loadedTemplate.Save(outputPath);
    }
}
