using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Category { get; set; } = string.Empty;

    public string IncludePath
    {
        get
        {
            return Category switch
            {
                "A" => Path.Combine(Directory.GetCurrentDirectory(), "categoryA.html"),
                "B" => Path.Combine(Directory.GetCurrentDirectory(), "categoryB.html"),
                _ => Path.Combine(Directory.GetCurrentDirectory(), "default.html")
            };
        }
    }

    public string HtmlFragment
    {
        get
        {
            // Load the HTML fragment file based on the category.
            return File.ReadAllText(IncludePath);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample HTML fragments.
        File.WriteAllText("categoryA.html", "<p><b>Category A</b> specific content.</p>");
        File.WriteAllText("categoryB.html", "<p><i>Category B</i> specific content.</p>");
        File.WriteAllText("default.html", "<p>Default content.</p>");

        // Create the LINQ Reporting template.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a title and the category value.
        builder.Writeln("Report for category: <<[model.Category]>>");
        // Dynamically include the HTML fragment based on the model.
        builder.Writeln("<<[model.HtmlFragment] -html>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Create the root data model.
        var model = new ReportModel { Category = "A" };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        templateDoc.Save(outputPath);
    }
}
