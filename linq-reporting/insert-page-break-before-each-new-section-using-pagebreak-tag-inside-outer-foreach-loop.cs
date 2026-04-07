using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer foreach over the collection of sections.
        builder.Writeln("<<foreach [section in Model.Sections]>>");

        // Insert a real page break before each new section.
        // The break node itself will be repeated for each iteration.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln(); // Ensure the break is in its own paragraph.

        // Output the section title.
        builder.Writeln("<<[section.Title]>>");

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        var model = new ReportModel
        {
            Sections = new List<SectionInfo>
            {
                new SectionInfo { Title = "First Section" },
                new SectionInfo { Title = "Second Section" },
                new SectionInfo { Title = "Third Section" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root object name used in the template is "Model".
        engine.BuildReport(doc, model, "Model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<SectionInfo> Sections { get; set; } = new();
}

public class SectionInfo
{
    public string Title { get; set; } = string.Empty;
}
