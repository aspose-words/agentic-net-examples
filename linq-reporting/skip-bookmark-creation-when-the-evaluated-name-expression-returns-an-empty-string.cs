using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure code page provider is available (required for some data sources).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "template.docx";
        string reportPath = "report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write the LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        // If the bookmark name is not empty, create a bookmark around the title.
        builder.Writeln("<<if [item.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");
        // If the bookmark name is empty, just write the title without a bookmark.
        builder.Writeln("<<if [item.BookmarkName == \"\"]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</if>>");
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
                new Item { Title = "First Section", BookmarkName = "FirstBookmark" },
                new Item { Title = "Second Section", BookmarkName = "" } // Empty bookmark name – should be skipped.
            }
        };

        // -------------------------------------------------
        // 3. Load the template and build the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);
    }
}

// -------------------------------------------------
// Data model classes.
// -------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Title { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
