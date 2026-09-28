using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Section
{
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
    public string BackgroundColor { get; set; } = "";
}

public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach over Sections.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Apply background color to the whole section (title and content).
        builder.Writeln("<<backColor [sec.BackgroundColor]>>");
        builder.Writeln("<<[sec.Title]>>");
        builder.Writeln("<<[sec.Content]>>");
        builder.Writeln("<</backColor>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Title = "Introduction",
                    Content = "This is the introduction section.",
                    BackgroundColor = "LightYellow"
                },
                new Section
                {
                    Title = "Details",
                    Content = "Detailed information goes here.",
                    BackgroundColor = "#D3E4CD"
                },
                new Section
                {
                    Title = "Conclusion",
                    Content = "Final remarks and summary.",
                    BackgroundColor = "LightBlue"
                }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        template.Save(outputPath);
    }
}
