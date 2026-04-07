using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents field that will pick up headings of level 1.
        builder.InsertTableOfContents("\\o \"1-1\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a LINQ Reporting foreach block that will generate headings.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Set the paragraph style to Heading1 before writing the tag.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[section.Title]>>");

        // Reset style to Normal for the next iteration.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section { Title = "Introduction" },
                new Section { Title = "Chapter 1: Getting Started" },
                new Section { Title = "Chapter 2: Advanced Topics" },
                new Section { Title = "Conclusion" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Update fields so the TOC reflects the generated headings.
        reportDoc.UpdateFields();

        // Save the final document.
        reportDoc.Save(reportPath);
    }
}

// -----------------------------------------------------------------
// Data model classes used by the LINQ Reporting engine.
// -----------------------------------------------------------------
public class ReportModel
{
    // Initialise the collection to avoid nullable warnings.
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
}
