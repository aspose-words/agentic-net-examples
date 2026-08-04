using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Category selected for the report.
    public string Category { get; set; } = "News";

    // Returns the HTML fragment content based on the Category value.
    public string IncludeHtml
    {
        get
        {
            // Build the file name (e.g., "News.html").
            string fileName = $"{Category}.html";
            string fullPath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

            // If the file does not exist, return an empty string to avoid runtime errors.
            return File.Exists(fullPath) ? File.ReadAllText(fullPath) : string.Empty;
        }
    }
}

public class Program
{
    private const string TemplateFileName = "Template.docx";
    private const string OutputFileName = "Report.docx";

    public static void Main()
    {
        // Prepare sample HTML fragments.
        CreateHtmlFragment("News.html", "<h1>News Section</h1><p>Latest news content goes here.</p>");
        CreateHtmlFragment("Blog.html", "<h1>Blog Section</h1><p>Latest blog post content goes here.</p>");

        // Create the LINQ Reporting template.
        CreateTemplate();

        // Load the template document.
        Document template = new Document(TemplateFileName);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Category = "Blog" // Change to "News" to include the other fragment.
        };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save(OutputFileName);
    }

    private static void CreateHtmlFragment(string fileName, string htmlContent)
    {
        // Write the HTML fragment to a file in the current directory.
        File.WriteAllText(Path.Combine(Directory.GetCurrentDirectory(), fileName), htmlContent);
    }

    private static void CreateTemplate()
    {
        // Build a simple template that includes an external HTML fragment based on the model.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Display the selected category.
        builder.Writeln("Selected Category: <<[model.Category]>>");

        // Insert the HTML fragment content using the supported -html switch.
        builder.Writeln("<<[model.IncludeHtml] -html>>");

        // Save the template for later processing.
        doc.Save(TemplateFileName);
    }
}
