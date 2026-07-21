using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer foreach over sections.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Insert a page break before each new section (inside the foreach loop).
        builder.InsertBreak(BreakType.PageBreak);

        // Write the section title.
        builder.Writeln("Section: <<[section.Title]>>");

        // Inner foreach over items within the section.
        builder.Writeln("<<foreach [item in section.Items]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for report generation.
        Document report = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Sections = new List<ReportSection>
            {
                new ReportSection
                {
                    Title = "First Section",
                    Items = new List<string> { "Item 1A", "Item 1B", "Item 1C" }
                },
                new ReportSection
                {
                    Title = "Second Section",
                    Items = new List<string> { "Item 2A", "Item 2B" }
                },
                new ReportSection
                {
                    Title = "Third Section",
                    Items = new List<string> { "Item 3A", "Item 3B", "Item 3C", "Item 3D" }
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<ReportSection> Sections { get; set; } = new();
}

// Section model.
public class ReportSection
{
    public string Title { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
