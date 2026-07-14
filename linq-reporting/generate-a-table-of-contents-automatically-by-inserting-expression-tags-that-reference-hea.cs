using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class SectionItem
{
    public string Title { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<SectionItem> Sections { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Sections = new List<SectionItem>
            {
                new() { Title = "Introduction" },
                new() { Title = "Chapter 1: Getting Started" },
                new() { Title = "Chapter 2: Advanced Topics" },
                new() { Title = "Conclusion" }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents field.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a foreach block that will generate headings.
        builder.Writeln("<<foreach [section in Sections]>>");
        // Set the paragraph style to Heading1 for each generated heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[section.Title]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Update fields so the TOC reflects the generated headings.
        doc.UpdateFields();

        // Save the final document.
        doc.Save("ReportWithToc.docx");
    }
}
