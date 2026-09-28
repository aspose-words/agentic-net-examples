using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class ReportItem
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel
        {
            Items = new List<ReportItem>
            {
                new ReportItem { Id = 1, Name = "Task A", Status = "Completed" },
                new ReportItem { Id = 2, Name = "Task B", Status = "Pending" },
                new ReportItem { Id = 3, Name = "Task C", Status = "Failed" }
            }
        };

        // Create template.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table for each iteration.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Status");
        builder.EndRow();

        // Data row with conditional background color.
        string backColorExpr = "<<backColor [item.Status == \"Completed\" ? \"LightGreen\" : item.Status == \"Pending\" ? \"LightYellow\" : \"LightCoral\"]>>";

        builder.InsertCell();
        builder.Writeln($"{backColorExpr}<<[item.Id]>> <</backColor>>");
        builder.InsertCell();
        builder.Writeln($"{backColorExpr}<<[item.Name]>> <</backColor>>");
        builder.InsertCell();
        builder.Writeln($"{backColorExpr}<<[item.Status]>> <</backColor>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save template.
        doc.Save(templatePath);

        // Load template and build report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save final report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
