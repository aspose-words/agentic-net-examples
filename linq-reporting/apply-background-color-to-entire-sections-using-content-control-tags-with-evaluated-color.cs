using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class SectionInfo
{
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
    public string Color { get; set; } = "";
}

public class ReportModel
{
    public List<SectionInfo> Sections { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Sections = new()
            {
                new() { Title = "Introduction", Content = "This is the introduction.", Color = "\"LightYellow\"" },
                new() { Title = "Details", Content = "Detailed information goes here.", Color = "\"LightGreen\"" },
                new() { Title = "Conclusion", Content = "Final thoughts.", Color = "\"LightBlue\"" }
            }
        };

        // Create the template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Insert a section break to start a new section for each item.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Apply background color to the whole section using backColor tag with evaluated expression.
        builder.Writeln("<<backColor [sec.Color]>>");

        // Section title.
        builder.Writeln("<<[sec.Title]>>");
        builder.Writeln();

        // Section content.
        builder.Writeln("<<[sec.Content]>>");
        builder.Writeln();

        // Close backColor tag.
        builder.Writeln("<</backColor>>");

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
