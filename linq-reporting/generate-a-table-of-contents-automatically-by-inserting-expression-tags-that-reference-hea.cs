using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    // Data model for the report
    public class ReportModel
    {
        public List<Chapter> Chapters { get; set; } = new();
    }

    public class Chapter
    {
        public string Title { get; set; } = string.Empty;
        public List<Section> Sections { get; set; } = new();
    }

    public class Section
    {
        public string Title { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Prepare sample data
        var model = new ReportModel
        {
            Chapters =
            {
                new Chapter
                {
                    Title = "Introduction",
                    Sections =
                    {
                        new Section { Title = "Purpose" },
                        new Section { Title = "Scope" }
                    }
                },
                new Chapter
                {
                    Title = "Usage",
                    Sections =
                    {
                        new Section { Title = "Installation" },
                        new Section { Title = "Configuration" }
                    }
                }
            }
        };

        // Paths for template and output
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);
        string templatePath = Path.Combine(workDir, "Template.docx");
        string resultPath = Path.Combine(workDir, "Report.docx");

        // ---------- Create template ----------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert TOC field (will be updated after report generation)
        builder.InsertTableOfContents("\\o \"1-2\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Begin foreach over chapters
        builder.Writeln("<<foreach [chapter in Chapters]>>");

        // Chapter heading (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[chapter.Title]>>");

        // Begin foreach over sections within the current chapter
        builder.Writeln("<<foreach [section in chapter.Sections]>>");

        // Section heading (Heading 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("<<[section.Title]>>");

        // End inner foreach
        builder.Writeln("<</foreach>>");

        // End outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // ---------- Load template and build report ----------
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        // Build the report using the model; root name is "model"
        engine.BuildReport(loadedDoc, model, "model");

        // Update fields (TOC) so that entries are generated
        loadedDoc.UpdateFields();

        // Save final document
        loadedDoc.Save(resultPath);
    }
}
