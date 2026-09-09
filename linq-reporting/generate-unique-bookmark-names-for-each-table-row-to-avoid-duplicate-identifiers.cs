using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table class

public class RowData
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public string BookmarkName { get; set; } = "";

    public RowData() { }

    public RowData(int index, string name)
    {
        Index = index;
        Name = name;
        // Generate a unique bookmark name for each row.
        BookmarkName = $"Row_{Guid.NewGuid():N}";
    }
}

public class ReportModel
{
    public List<RowData> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -------------------------
        // Create the template document
        // -------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Table with unique bookmarks per row:");
        // Start the foreach block.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build a table inside the foreach.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Bookmark");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        // Bookmark tag with a unique name per row.
        builder.Writeln("<<bookmark [item.BookmarkName]>>Bookmark Content<</bookmark>>");
        builder.EndRow();

        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -------------------------
        // Load the template for reporting
        // -------------------------
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel();
        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new RowData(i, $"Item {i}"));
        }

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
