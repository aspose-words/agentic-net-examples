using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // 1. Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Table with Bookmarks and Hyperlinks");

        // Begin foreach over the collection Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table for each iteration.
        Table table = builder.StartTable();

        // Header row (only once, but placed inside foreach for simplicity).
        builder.InsertCell();
        builder.Writeln("ID");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Link");
        builder.EndRow();

        // Data row.
        // Cell 1 – bookmark around the ID.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.Bookmark]>>");
        builder.Writeln("<<[item.Id]>>");
        builder.Writeln("<</bookmark>>");

        // Cell 2 – plain name.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");

        // Cell 3 – hyperlink that navigates to the bookmark defined above.
        builder.InsertCell();
        builder.Writeln("<<link [item.Bookmark] [item.Name]>>");

        // End the data row.
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // 2. Prepare the data model.
        ReportModel model = new ReportModel
        {
            Items = new List<RowItem>
            {
                new RowItem { Id = 1, Name = "Alpha",   Bookmark = "bm1" },
                new RowItem { Id = 2, Name = "Beta",    Bookmark = "bm2" },
                new RowItem { Id = 3, Name = "Gamma",   Bookmark = "bm3" }
            }
        };

        // 3. Load the template and build the report.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // 4. Save the final document.
        report.Save(reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<RowItem> Items { get; set; } = new();
}

// Individual row data.
public class RowItem
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Bookmark { get; set; } = "";
}
