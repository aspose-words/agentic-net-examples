using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple title.
        builder.Writeln("LINQ Reporting – List Numbering Restart Example");

        // Outer loop over sections.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Section heading.
        builder.Writeln("Section: <<[section.Name]>>");

        // Numbered list of items. The <<restartNum>> tag forces numbering to restart
        // for each new section because it is placed before the inner foreach in the same paragraph.
        builder.Writeln("1. <<restartNum>><<foreach [item in section.Items]>> <<[item.Description]>> <</foreach>>");

        // End of outer loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Sections = new()
            {
                new Section
                {
                    Name = "First Section",
                    Items = new()
                    {
                        new Item { Description = "First item" },
                        new Item { Description = "Second item" },
                        new Item { Description = "Third item" }
                    }
                },
                new Section
                {
                    Name = "Second Section",
                    Items = new()
                    {
                        new Item { Description = "Alpha" },
                        new Item { Description = "Beta" }
                    }
                },
                new Section
                {
                    Name = "Third Section",
                    Items = new()
                    {
                        new Item { Description = "One" },
                        new Item { Description = "Two" },
                        new Item { Description = "Three" },
                        new Item { Description = "Four" }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        // Load the template (could also reuse the same Document instance).
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
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
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Description { get; set; } = string.Empty;
}
