using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Title = "First item", BookmarkName = "bmFirst" },
                new Item { Title = "Second item", BookmarkName = "bmSecond" },
                new Item { Title = "Third item", BookmarkName = "bmThird" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a bullet list.
        List list = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;

        // Insert LINQ Reporting tags.
        // The foreach iterates over Items, and each list item contains a bookmark.
        builder.Writeln("<<foreach [item in Items]>>");
        // Open bookmark tag, the expression returns the bookmark name.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        // The actual text of the list item.
        builder.Writeln("<<[item.Title]>>");
        // Close bookmark tag.
        builder.Writeln("<</bookmark>>");
        // End of foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(report, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
