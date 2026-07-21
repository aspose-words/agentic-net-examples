using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data.
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
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Create a numbered list.
        builder.ListFormat.List = template.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberDefault);

        // LINQ Reporting tags:
        //   <<foreach [item in Items]>> ... <</foreach>>
        //   <<bookmark [item.BookmarkName]>> ... <</bookmark>>
        //   <<[item.Title]>>
        builder.Writeln("<<foreach [item in Items]>>");
        // Each paragraph becomes a list item because ListFormat is active.
        builder.Writeln("<<bookmark [item.BookmarkName]>><<[item.Title]>> <</bookmark>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and generate the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
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
