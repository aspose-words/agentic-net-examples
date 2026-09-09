using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class HtmlModel
{
    public string HtmlSnippet { get; set; }

    public HtmlModel(string htmlSnippet)
    {
        HtmlSnippet = htmlSnippet ?? string.Empty;
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a sample HTML snippet file.
        string htmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "snippet.html");
        if (!File.Exists(htmlFilePath))
        {
            const string sampleHtml = "<h1 style=\"color:Blue;\">Hello Aspose.Words!</h1>" +
                                     "<p>This is an <b>HTML</b> fragment inserted via LINQ Reporting.</p>";
            File.WriteAllText(htmlFilePath, sampleHtml);
        }

        // Read the HTML content.
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Create the data model.
        HtmlModel model = new HtmlModel(htmlContent);

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that outputs the HTML snippet using the -html switch.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(template, model, "model");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithHtml.docx");
        template.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
