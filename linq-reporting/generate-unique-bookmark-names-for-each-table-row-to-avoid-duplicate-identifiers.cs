using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

public class Program
{
    public static void Main()
    {
        // Prepare sample data with unique bookmark names.
        var model = new ReportModel
        {
            Items = new List<RowItem>()
        };

        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new RowItem
            {
                Index = i,
                Name = $"Item {i}",
                // Generate a unique bookmark name for each row.
                BookmarkName = $"RowBookmark_{i}"
            });
        }

        // Create the LINQ Reporting template programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template and build the report.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the foreach loop over the collection named Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with a header row.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Bookmark");
        builder.EndRow();

        // Row template: each cell will be filled from the current item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");

        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");

        builder.InsertCell();
        // Insert a bookmark whose name comes from the data source.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        // Content inside the bookmark (can be any text, here we repeat the name).
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");

        builder.EndRow();

        // Finish the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(filePath);
    }
}

// Root model passed to the reporting engine.
public class ReportModel
{
    public List<RowItem> Items { get; set; } = new();
}

// Individual row data.
public class RowItem
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
