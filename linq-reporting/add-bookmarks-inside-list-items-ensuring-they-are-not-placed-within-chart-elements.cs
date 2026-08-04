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
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a numbered list that will be used for the items.
        builder.ListFormat.List = templateDoc.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberDefault);

        // Write the foreach tag that iterates over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Each iteration is a list paragraph. Insert a bookmark that wraps the item's title.
        // The bookmark name is taken from the data source (item.BookmarkName).
        builder.Writeln("<<bookmark [item.BookmarkName]>><<[item.Title]>><</bookmark>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Title = "First item", BookmarkName = "bmFirst" },
                new Item { Title = "Second item", BookmarkName = "bmSecond" },
                new Item { Title = "Third item", BookmarkName = "bmThird" }
            }
        };

        // -------------------------------------------------
        // 3. Load the template and build the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}

// -------------------------------------------------
// Data model classes.
// -------------------------------------------------
public class ReportModel
{
    // Collection of items to be listed.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Text displayed for the list item.
    public string Title { get; set; } = string.Empty;

    // Name of the bookmark that will be placed around the title.
    public string BookmarkName { get; set; } = string.Empty;
}
