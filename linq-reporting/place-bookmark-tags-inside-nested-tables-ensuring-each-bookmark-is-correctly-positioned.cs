using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // ----- Outer table ------------------------------------------------
        Table outerTable = builder.StartTable();

        // First cell of the outer table.
        builder.InsertCell();
        builder.Writeln("Outer cell 1");

        // Second cell will contain the inner table.
        builder.InsertCell();
        builder.Writeln("Outer cell 2 with inner table:");

        // ----- Inner table (nested) --------------------------------------
        Table innerTable = builder.StartTable();

        // Single cell in the inner table that holds a bookmark.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");

        // End the inner table.
        builder.EndRow();
        builder.EndTable();

        // End the outer table row and the outer table itself.
        builder.EndRow();
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "BookmarkNestedTableTemplate.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<SectionItem>
            {
                new SectionItem { BookmarkName = "BM_First", Title = "First Item Title" },
                new SectionItem { BookmarkName = "BM_Second", Title = "Second Item Title" },
                new SectionItem { BookmarkName = "BM_Third", Title = "Third Item Title" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        const string outputPath = "BookmarkNestedTableReport.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (must be public with public properties).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<SectionItem> Items { get; set; } = new();
}

public class SectionItem
{
    public string BookmarkName { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
}
