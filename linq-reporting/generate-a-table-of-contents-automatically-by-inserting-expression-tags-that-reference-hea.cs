using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Section
{
    public string Title { get; set; } = "";
    public int Level { get; set; }
    public string Bookmark { get; set; } = "";
}

public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title page.
        builder.Writeln("Sample Report");
        builder.Writeln();

        // -----------------------------------------------------------------
        // 2. Insert Table of Contents placeholder using LINQ Reporting tags.
        // -----------------------------------------------------------------
        builder.Writeln("Table of Contents");
        builder.Writeln("<<foreach [sec in Sections]>>");
        // Indent based on heading level (simple 2 spaces per level).
        builder.Writeln("<<if [sec.Level == 1]>>");
        builder.Writeln("<<bookmark [sec.Bookmark]>><<[sec.Title]>> <</bookmark>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<<if [sec.Level == 2]>>");
        builder.Writeln("  <<bookmark [sec.Bookmark]>><<[sec.Title]>> <</bookmark>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Page break before the actual content.
        builder.InsertBreak(BreakType.PageBreak);

        // -----------------------------------------------------------------
        // 3. Insert headings placeholder using LINQ Reporting tags.
        // -----------------------------------------------------------------
        builder.Writeln("<<foreach [sec in Sections]>>");
        // Apply Heading 1 style for all sections (simplified).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<bookmark [sec.Bookmark]>><<[sec.Title]>> <</bookmark>>");
        // Reset style to Normal after heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Prepare sample data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section { Title = "Introduction", Level = 1, Bookmark = "Intro" },
                new Section { Title = "Background", Level = 1, Bookmark = "Background" },
                new Section { Title = "Details", Level = 1, Bookmark = "Details" },
                new Section { Title = "Subsection A", Level = 2, Bookmark = "SubA" },
                new Section { Title = "Subsection B", Level = 2, Bookmark = "SubB" },
                new Section { Title = "Conclusion", Level = 1, Bookmark = "Conclusion" }
            }
        };

        // -----------------------------------------------------------------
        // 6. Build the report using LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}
