using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample HTML fragments.
        File.WriteAllText("categoryA.html", "<h2>Category A Content</h2><p>This is content for category A.</p>");
        File.WriteAllText("categoryB.html", "<h2>Category B Content</h2><p>This is content for category B.</p>");

        // Build the LINQ Reporting template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Show the selected category.
        builder.Writeln("Selected Category: <<[model.Category]>>");

        // Insert the HTML fragment returned by the model.
        // The expression evaluates to an HTML string, and the -html switch tells the engine to treat it as HTML.
        builder.Writeln("<<[model.IncludeHtml] -html>>");

        // Save the template.
        template.Save("template.docx");

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Category = "A" // Change to "B" to include the other fragment.
        };

        // Load the template and generate the report.
        Document report = new Document("template.docx");
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        report.Save("output.docx");
    }
}

// Public data model aligned with the template.
public class ReportModel
{
    // Category used in the template.
    public string Category { get; set; } = string.Empty;

    // Returns the HTML fragment content based on the Category value.
    public string IncludeHtml
    {
        get
        {
            string filePath = Category switch
            {
                "A" => "categoryA.html",
                "B" => "categoryB.html",
                _ => "categoryA.html"
            };

            // Read and return the HTML file content.
            return File.Exists(filePath) ? File.ReadAllText(filePath) : string.Empty;
        }
    }
}
