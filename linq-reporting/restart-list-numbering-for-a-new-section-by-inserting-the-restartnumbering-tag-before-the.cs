using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Title = "First Section",
                    Items = new List<string> { "Item A1", "Item A2", "Item A3" }
                },
                new Section
                {
                    Title = "Second Section",
                    Items = new List<string> { "Item B1", "Item B2", "Item B3", "Item B4" }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Outer foreach iterates over sections.
        builder.Writeln("<<foreach [section in model.Sections]>>");
        // Write the section title.
        builder.Writeln("<<[section.Title]>>");

        // Apply a numbered list to the upcoming paragraphs.
        builder.ListFormat.List = templateDoc.Lists.Add(ListTemplate.NumberDefault);

        // Insert the restartNum tag immediately before the inner foreach.
        // This restarts numbering for each new section.
        builder.Writeln("<<restartNum>><<foreach [item in section.Items]>><<[item]>>");
        // Close the inner foreach.
        builder.Writeln("<</foreach>>");

        // Remove list formatting after the inner list is finished.
        builder.ListFormat.RemoveNumbers();

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report; the root object name is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
