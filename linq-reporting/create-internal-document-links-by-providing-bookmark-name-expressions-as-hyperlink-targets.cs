using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Table of contents with internal links.
        builder.Writeln("Table of Contents");
        builder.Writeln("<<foreach [section in model.Sections]>>");
        builder.Writeln("<<link [section.BookmarkName] [section.Title]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln(); // empty paragraph

        // Sections with bookmarks.
        builder.Writeln("<<foreach [section in model.Sections]>>");
        builder.Writeln("<<bookmark [section.BookmarkName]>>");
        builder.Writeln("<<[section.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln(); // empty paragraph
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section { Title = "Introduction", BookmarkName = "bmIntro" },
                new Section { Title = "Chapter 1", BookmarkName = "bmChapter1" },
                new Section { Title = "Conclusion", BookmarkName = "bmConclusion" }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
