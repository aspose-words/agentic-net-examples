using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the final report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents field at the beginning.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC.

        // Begin a foreach block that will repeat for each section.
        builder.Writeln("<<foreach [section in Model.Sections]>>");

        // Insert a heading for the section title. Use Heading 1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[section.Title]>>");

        // Insert the section content as normal paragraph text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<<[section.Content]>>");
        builder.Writeln(); // Blank line between sections.

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section { Title = "Introduction", Content = "This is the introduction section." },
                new Section { Title = "Usage", Content = "Details on how to use the application." },
                new Section { Title = "Conclusion", Content = "Final thoughts and summary." }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "Model");

        // Update fields to generate the Table of Contents.
        reportDoc.UpdateFields();

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
}
