using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // NOTE: The LINQ Reporting engine does not have a dedicated <<toc>> tag.
        // We will insert the Table of Contents after the report is built.

        // Begin foreach block for sections.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Section title – Heading 1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[sec.Title]>>");

        // Section content – normal style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<<[sec.Content]>>");

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section { Title = "Introduction", Content = "This is the introduction section." },
                new Section { Title = "Chapter 1", Content = "Details of chapter one go here." },
                new Section { Title = "Conclusion", Content = "Final thoughts are presented in this section." }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Insert a Table of Contents at the beginning of the generated document.
        // -----------------------------------------------------------------
        DocumentBuilder tocBuilder = new DocumentBuilder(reportDoc);
        tocBuilder.MoveToDocumentStart();
        tocBuilder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        reportDoc.UpdateFields(); // Populate the TOC entries.

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

// Section model.
public class Section
{
    public string Title { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
}
