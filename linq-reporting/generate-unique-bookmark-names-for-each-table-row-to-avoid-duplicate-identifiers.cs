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
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new RowItem { Index = 1, Name = "Alpha" },
                new RowItem { Index = 2, Name = "Beta" },
                new RowItem { Index = 3, Name = "Gamma" }
            }
        };

        // Create the LINQ Reporting template.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template and build the report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row with a unique bookmark per row.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Index]>>");
        builder.Writeln("<</bookmark>>");

        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Root data model.
public class ReportModel
{
    public List<RowItem> Items { get; set; } = new();
}

// Individual row item.
public class RowItem
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    // Unique bookmark name based on the row index.
    public string BookmarkName => $"Row_{Index}";
}
