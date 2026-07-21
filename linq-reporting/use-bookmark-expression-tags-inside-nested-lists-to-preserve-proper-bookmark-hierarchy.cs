using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Categories = new()
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new()
                    {
                        new Item { Title = "Apple",  BookmarkName = "bm_Apple" },
                        new Item { Title = "Banana", BookmarkName = "bm_Banana" }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new()
                    {
                        new Item { Title = "Carrot",   BookmarkName = "bm_Carrot" },
                        new Item { Title = "Tomato",   BookmarkName = "bm_Tomato" }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Define outer (numbered) and inner (bulleted) list styles.
        List outerList = doc.Lists.Add(ListTemplate.NumberArabicDot);
        List innerList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Begin outer foreach over Categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.ListFormat.List = outerList;
        builder.Writeln("<<[cat.Name]>>"); // Category name as outer list item.

        // Begin inner foreach over Items of the current Category.
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.ListFormat.List = innerList;
        // Bookmark tag preserving hierarchy.
        builder.Writeln("<<bookmark [itm.BookmarkName]>><<[itm.Title]>><</bookmark>>");
        builder.Writeln("<</foreach>>"); // End inner foreach.

        // Reset to outer list for the next category.
        builder.ListFormat.List = outerList;
        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options.

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
