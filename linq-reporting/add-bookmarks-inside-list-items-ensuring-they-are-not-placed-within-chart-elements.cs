using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;          // For ListTemplate
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts; // For ChartType

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Title = "First item", BookmarkName = "bmFirst" },
                new Item { Title = "Second item", BookmarkName = "bmSecond" },
                new Item { Title = "Third item", BookmarkName = "bmThird" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Title paragraph.
        builder.Writeln("Report with bookmarks inside list items:");

        // Start a numbered list.
        builder.ListFormat.List = template.Lists.Add(ListTemplate.NumberDefault);

        // LINQ Reporting foreach block iterating over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a bookmark that wraps the list item text.
        // The bookmark expression must return a non‑empty string.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");

        // End of foreach block.
        builder.Writeln("<</foreach>>");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Insert a chart after the list – bookmarks are **not** placed inside the chart.
        builder.Writeln("\nSample chart:");
        builder.InsertChart(ChartType.Column, 400, 300);

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var report = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root object name used in the template is "model".
        engine.BuildReport(report, model, "model");

        // Save the final report.
        report.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes – must be public with public properties.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
