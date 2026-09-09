using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Category used to decide which HTML fragment to include.
    public string Category { get; set; } = string.Empty;

    // Returns the HTML fragment content based on the Category value.
    public string HtmlFragment
    {
        get
        {
            // Determine file name according to the category.
            string fileName = Category switch
            {
                "A" => "fragmentA.html",
                "B" => "fragmentB.html",
                _ => "default.html"
            };

            // Build full path relative to the current working directory.
            string path = Path.Combine(Directory.GetCurrentDirectory(), fileName);

            // If the file does not exist, return a placeholder.
            return File.Exists(path) ? File.ReadAllText(path) : $"<p>Missing fragment for category '{Category}'.</p>";
        }
    }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample HTML fragments that will be inserted later.
        // -----------------------------------------------------------------
        File.WriteAllText("fragmentA.html", "<h2>Fragment A</h2><p>This is content for category A.</p>");
        File.WriteAllText("fragmentB.html", "<h2>Fragment B</h2><p>This is content for category B.</p>");
        File.WriteAllText("default.html",   "<h2>Default Fragment</h2><p>Fallback content.</p>");

        // -----------------------------------------------------------------
        // 2. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Show the selected category in the report.
        builder.Writeln("Selected Category: <<[model.Category]>>");

        // Dynamically include the HTML fragment based on the model's property.
        // The -html switch tells the engine to treat the string as HTML.
        builder.Writeln("<<[model.HtmlFragment] -html>>");

        // Save the template to disk (required before BuildReport).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // Example model – change the Category value to see different fragments.
        ReportModel model = new ReportModel { Category = "A" };

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root object name in the template is "model".
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
