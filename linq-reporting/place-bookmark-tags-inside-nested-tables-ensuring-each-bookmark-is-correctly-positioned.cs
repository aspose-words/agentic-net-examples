using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<OuterItem>
            {
                new OuterItem
                {
                    Header = "Section 1",
                    Details = new List<InnerItem>
                    {
                        new InnerItem { BookmarkName = "Bkm1", Text = "First inner item" },
                        new InnerItem { BookmarkName = "Bkm2", Text = "Second inner item" }
                    }
                },
                new OuterItem
                {
                    Header = "Section 2",
                    Details = new List<InnerItem>
                    {
                        new InnerItem { BookmarkName = "Bkm3", Text = "Third inner item" },
                        new InnerItem { BookmarkName = "Bkm4", Text = "Fourth inner item" }
                    }
                }
            }
        };

        // Create the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin outer foreach.
        builder.Writeln("<<foreach [outer in Model.Items]>>");

        // Start outer table.
        Table outerTable = builder.StartTable();

        // First cell: outer header.
        builder.InsertCell();
        builder.Writeln("<<[outer.Header]>>");

        // Second cell: will contain inner table.
        builder.InsertCell();

        // Begin inner foreach inside the second cell.
        builder.Writeln("<<foreach [inner in outer.Details]>>");

        // Start inner table.
        Table innerTable = builder.StartTable();

        // Cell with bookmark.
        builder.InsertCell();
        builder.Writeln("<<bookmark [inner.BookmarkName]>>");
        builder.Writeln("<<[inner.Text]>>");
        builder.Writeln("<</bookmark>>");

        // End inner row and table.
        builder.EndRow();
        builder.EndTable();

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer row.
        builder.EndRow();

        // End outer table.
        builder.EndTable();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Build the report using the LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "Model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<OuterItem> Items { get; set; } = new();
}

public class OuterItem
{
    public string Header { get; set; } = "";
    public List<InnerItem> Details { get; set; } = new();
}

public class InnerItem
{
    public string BookmarkName { get; set; } = "";
    public string Text { get; set; } = "";
}
