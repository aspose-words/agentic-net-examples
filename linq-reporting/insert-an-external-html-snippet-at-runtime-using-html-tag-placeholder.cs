using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // HTML snippet that will be inserted into the document at runtime.
    public string HtmlSnippet { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a title.
        builder.Writeln("Report with HTML snippet:");
        // Insert the LINQ Reporting HTML tag placeholder.
        // The tag will be replaced with the value of model.HtmlSnippet at build time.
        builder.Writeln("<<html [model.HtmlSnippet]>>");

        // 2. Prepare the data model with a realistic HTML fragment.
        ReportModel model = new ReportModel
        {
            HtmlSnippet = @"
                <p style='color:blue;'>
                    This is <b>HTML</b> content inserted at runtime.
                </p>"
        };

        // 3. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used in the template to reference the root object.
        engine.BuildReport(template, model, "model");

        // 4. Save the generated document.
        template.Save("ReportWithHtml.docx");
    }
}
