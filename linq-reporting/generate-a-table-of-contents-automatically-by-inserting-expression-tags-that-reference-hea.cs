using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Sections = new List<SectionInfo>
            {
                new SectionInfo { Title = "Introduction", Level = 1 },
                new SectionInfo { Title = "Background", Level = 2 },
                new SectionInfo { Title = "Purpose", Level = 2 },
                new SectionInfo { Title = "Details", Level = 1 },
                new SectionInfo { Title = "Sub‑detail A", Level = 2 },
                new SectionInfo { Title = "Sub‑detail B", Level = 2 },
                new SectionInfo { Title = "Conclusion", Level = 1 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a Table of Contents field that will pick up headings 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // -----------------------------------------------------------------
        // 2. Insert headings using LINQ Reporting tags.
        // -----------------------------------------------------------------
        // Level 1 headings.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<foreach [sec in Sections]>>");
        builder.Writeln("<<if [sec.Level == 1]>>");
        builder.Writeln("<<[sec.Title]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Level 2 headings.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("<<foreach [sec in Sections]>>");
        builder.Writeln("<<if [sec.Level == 2]>>");
        builder.Writeln("<<[sec.Title]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Level 3 headings (optional, shown for completeness).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("<<foreach [sec in Sections]>>");
        builder.Writeln("<<if [sec.Level == 3]>>");
        builder.Writeln("<<[sec.Title]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(report, model, "model");

        // Update fields so that the TOC reflects the generated headings.
        report.UpdateFields();

        // Save the final report.
        const string reportPath = "Report.docx";
        report.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<SectionInfo> Sections { get; set; } = new();
}

public class SectionInfo
{
    public string Title { get; set; } = string.Empty;
    public int Level { get; set; }
}
