using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Conditionally create a bookmark only when the name is not empty.
        builder.Writeln("<<if [item.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");
        // Save the template.
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<ReportItem>
            {
                new ReportItem { Title = "First Item", BookmarkName = "FirstBookmark" },
                new ReportItem { Title = "Second Item", BookmarkName = "" } // Empty name – bookmark will be skipped.
            }
        };

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

// Item model used in the foreach loop.
public class ReportItem
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
