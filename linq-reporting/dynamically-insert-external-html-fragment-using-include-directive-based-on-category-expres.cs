using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Category name (e.g., "News" or "Sports").
    public string Category { get; set; } = string.Empty;

    // Full path to the HTML fragment that should be included.
    public string HtmlPath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample HTML fragments that will be included later.
        // -----------------------------------------------------------------
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        string newsHtml = Path.Combine(workDir, "news.html");
        string sportsHtml = Path.Combine(workDir, "sports.html");

        File.WriteAllText(newsHtml, "<p><b>News:</b> Latest headlines go here.</p>");
        File.WriteAllText(sportsHtml, "<p><b>Sports:</b> Recent match results go here.</p>");

        // -----------------------------------------------------------------
        // 2. Create the data model. Choose a category and set the path to
        //    the corresponding HTML fragment.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Category = "News",
            HtmlPath = newsHtml // Change to sportsHtml to include the sports fragment.
        };

        // -----------------------------------------------------------------
        // 3. Build the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Show the selected category.
        builder.Writeln("Category: <<[model.Category]>>");

        // Dynamically include the external HTML fragment using the supported html tag.
        builder.Writeln("<<html [model.HtmlPath]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and generate the report using LINQ Reporting.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report; the root data source name must match the tag prefix.
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(workDir, "report.docx");
        reportDoc.Save(resultPath);
    }
}
