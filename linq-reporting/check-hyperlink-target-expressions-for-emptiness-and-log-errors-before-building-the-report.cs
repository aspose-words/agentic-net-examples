using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data with valid and invalid hyperlink targets.
        ReportModel model = new ReportModel
        {
            Items = new List<ReportItem>
            {
                new ReportItem { Title = "Aspose", Url = "https://www.aspose.com" },
                new ReportItem { Title = "EmptyLink", Url = "" },               // Invalid URL
                new ReportItem { Title = "NullLink", Url = null }               // Invalid URL
            }
        };

        // Log missing hyperlink targets before building the report.
        foreach (ReportItem item in model.Items)
        {
            if (string.IsNullOrWhiteSpace(item.Url))
                Console.WriteLine($"[Error] Hyperlink target is missing for item \"{item.Title}\".");
        }

        // Create the template document programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Build the report with inline error messages enabled.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    // Generates a simple template that iterates over Items and creates a hyperlink
    // only when the Url is not empty; otherwise it writes the title as plain text.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // If Url is non‑empty, insert a link; otherwise output the title without a link.
        builder.Writeln("<<if [item.Url != null && item.Url != \"\"]>>");
        builder.Writeln("<<link [item.Url] [item.Title]>>");
        builder.Writeln("<</if>>");

        builder.Writeln("<<if [item.Url == null || item.Url == \"\"]>>");
        builder.Writeln("[item.Title] (missing link)");
        builder.Writeln("<</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Root data model.
public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

// Individual item containing a title and a (possibly null) hyperlink target.
public class ReportItem
{
    public string Title { get; set; } = string.Empty;
    public string? Url { get; set; }
}
