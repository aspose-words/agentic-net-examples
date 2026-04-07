using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Added for Table type

public class Program
{
    public static void Main()
    {
        // Prepare sample data with unique bookmark names.
        var model = new ReportModel
        {
            Rows = new List<RowData>()
        };

        for (int i = 1; i <= 5; i++)
        {
            model.Rows.Add(new RowData
            {
                Name = $"Item {i}",
                Value = (i * 10).ToString(),
                // Generate a unique bookmark name for each row.
                Bookmark = $"RowBookmark_{i}"
            });
        }

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Report with unique bookmarks per row:");
        // Begin the foreach block that iterates over the Rows collection.
        builder.Writeln("<<foreach [row in Rows]>>");

        // Build a table inside the foreach block.
        Table table = builder.StartTable();

        // First column – contains a bookmark that wraps the row's Name.
        builder.InsertCell();
        builder.Writeln("<<bookmark [row.Bookmark]>>"); // Open bookmark.
        builder.Writeln("<<[row.Name]>>");              // Row content.
        builder.Writeln("<</bookmark>>");               // Close bookmark.

        // Second column – displays the row's Value.
        builder.InsertCell();
        builder.Writeln("<<[row.Value]>>");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of rows to be displayed in the report.
    public List<RowData> Rows { get; set; } = new();
}

public class RowData
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    // Unique bookmark identifier for this row.
    public string Bookmark { get; set; } = string.Empty;
}
