using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model
        var model = new ReportModel
        {
            OuterItems = new List<OuterItem>
            {
                new OuterItem
                {
                    Name = "Outer 1",
                    InnerItems = new List<InnerItem>
                    {
                        new InnerItem { Title = "Inner 1A", BookmarkName = "BM_1A" },
                        new InnerItem { Title = "Inner 1B", BookmarkName = "BM_1B" }
                    }
                },
                new OuterItem
                {
                    Name = "Outer 2",
                    InnerItems = new List<InnerItem>
                    {
                        new InnerItem { Title = "Inner 2A", BookmarkName = "BM_2A" },
                        new InnerItem { Title = "Inner 2B", BookmarkName = "BM_2B" }
                    }
                }
            }
        };

        // Create template document programmatically
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin outer foreach
        builder.Writeln("<<foreach [outer in Model.OuterItems]>>");

        // Outer table start
        builder.StartTable();

        // First cell: outer name
        builder.InsertCell();
        builder.Writeln("<<[outer.Name]>>");

        // Second cell: will contain inner table
        builder.InsertCell();

        // Begin inner foreach
        builder.Writeln("<<foreach [inner in outer.InnerItems]>>");

        // Inner table start
        builder.StartTable();

        // Inner row first cell with bookmark
        builder.InsertCell();
        builder.Writeln("<<bookmark [inner.BookmarkName]>>");
        builder.Writeln("<<[inner.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Inner row second cell (optional additional content)
        builder.InsertCell();
        builder.Writeln("Details for <<[inner.Title]>>");

        // End inner row
        builder.EndRow();

        // End inner table
        builder.EndTable();

        // End inner foreach
        builder.Writeln("<</foreach>>");

        // End outer row
        builder.EndRow();

        // End outer table
        builder.EndTable();

        // End outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // Load the template for reporting
        var templateDoc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "Model");

        // Save the generated report
        var outputPath = "output.docx";
        templateDoc.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<OuterItem> OuterItems { get; set; } = new();
}

public class OuterItem
{
    public string Name { get; set; } = string.Empty;
    public List<InnerItem> InnerItems { get; set; } = new();
}

public class InnerItem
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
