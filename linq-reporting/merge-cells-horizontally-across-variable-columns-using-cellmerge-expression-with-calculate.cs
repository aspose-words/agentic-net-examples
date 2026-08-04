using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data model
        var model = new ReportModel
        {
            Categories = new()
            {
                new Category { Name = "Group A", Span = 2 },
                new Category { Name = "Group B", Span = 3 }
            },
            Rows = new()
            {
                new DataRow { Values = new() { "A1", "A2", "B1", "B2", "B3" } },
                new DataRow { Values = new() { "A3", "A4", "B4", "B5", "B6" } }
            }
        };

        // Create template document
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Table with merged header cells:");

        // Start table
        builder.StartTable();

        // First header row: merge cells horizontally using <<cellMerge>> tag
        foreach (var cat in model.Categories)
        {
            for (int i = 0; i < cat.Span; i++)
            {
                builder.InsertCell();
                builder.Writeln($"<<cellMerge>>{cat.Name}");
            }
        }
        builder.EndRow();

        // Second header row: sub‑column titles
        int subIndex = 1;
        foreach (var cat in model.Categories)
        {
            for (int i = 0; i < cat.Span; i++)
            {
                builder.InsertCell();
                builder.Writeln($"Sub {subIndex++}");
            }
        }
        builder.EndRow();

        // Data rows (filled directly, no LINQ tags needed)
        foreach (var row in model.Rows)
        {
            foreach (var val in row.Values)
            {
                builder.InsertCell();
                builder.Writeln(val);
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Process the <<cellMerge>> tags
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "Model");

        // Save the generated report
        doc.Save("Report.docx");
    }
}

// Data model classes
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
    public List<DataRow> Rows { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = "";
    public int Span { get; set; }
}

public class DataRow
{
    public List<string> Values { get; set; } = new();
}
