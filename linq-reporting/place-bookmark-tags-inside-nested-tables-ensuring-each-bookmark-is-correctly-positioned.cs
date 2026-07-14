using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class BookmarkInNestedTables
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document with nested tables.
        // -------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer foreach – iterate over Parents.
        builder.Writeln("<<foreach [parent in Parents]>>");

        // Start outer table.
        var outerTable = builder.StartTable();

        // First cell of the outer row.
        builder.InsertCell();

        // Inner foreach – iterate over Children of the current Parent.
        builder.Writeln("<<foreach [child in parent.Children]>>");

        // Start inner table.
        var innerTable = builder.StartTable();

        // Cell that will contain the bookmark.
        builder.InsertCell();

        // Bookmark tag: the expression returns the bookmark name.
        builder.Writeln("<<bookmark [child.BookmarkName]>>");
        // Content inside the bookmark.
        builder.Writeln("<<[child.Text]>>");
        // Closing bookmark tag.
        builder.Writeln("<</bookmark>>");

        // End inner table row and table.
        builder.EndRow();
        builder.EndTable();

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer table row and table.
        builder.EndRow();
        builder.EndTable();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare the data model.
        // -------------------------------------------------
        var model = new ReportModel
        {
            Parents = new List<Parent>
            {
                new Parent
                {
                    Children = new List<Child>
                    {
                        new Child { BookmarkName = "BM1", Text = "First child of first parent" },
                        new Child { BookmarkName = "BM2", Text = "Second child of first parent" }
                    }
                },
                new Parent
                {
                    Children = new List<Child>
                    {
                        new Child { BookmarkName = "BM3", Text = "First child of second parent" },
                        new Child { BookmarkName = "BM4", Text = "Second child of second parent" }
                    }
                }
            }
        };

        // -------------------------------------------------
        // 3. Load the template and build the report.
        // -------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the final document.
        report.Save(outputPath);
    }
}

// -----------------------------------------------------------------
// Data model classes – all members are public and initialized.
// -----------------------------------------------------------------
public class ReportModel
{
    public List<Parent> Parents { get; set; } = new();
}

public class Parent
{
    public List<Child> Children { get; set; } = new();
}

public class Child
{
    public string BookmarkName { get; set; } = "";
    public string Text { get; set; } = "";
}
