using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Rows = new List<RowData>
            {
                // Row where first three cells have the same text – they will merge horizontally.
                new RowData
                {
                    Cell1 = "Group A",
                    Cell2 = "Group A",
                    Cell3 = "Group A",
                    Cell4 = "Item 1",
                    Cell5 = "100"
                },
                // Row where only first two cells are the same – they will merge.
                new RowData
                {
                    Cell1 = "Group B",
                    Cell2 = "Group B",
                    Cell3 = "Detail",
                    Cell4 = "Item 2",
                    Cell5 = "200"
                },
                // Row with no merging – all cells have distinct values.
                new RowData
                {
                    Cell1 = "Solo",
                    Cell2 = "Info",
                    Cell3 = "Detail",
                    Cell4 = "Item 3",
                    Cell5 = "300"
                }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("LINQ Reporting – Horizontal Cell Merge Example");
        builder.Writeln();

        // Begin the foreach block that iterates over Model.Rows.
        builder.Writeln("<<foreach [row in Model.Rows]>>");

        // Start a table for each row.
        Table table = builder.StartTable();

        // Header row (optional – shown for clarity).
        builder.InsertCell();
        builder.Writeln("Category");
        builder.InsertCell();
        builder.Writeln("Subcategory");
        builder.InsertCell();
        builder.Writeln("Detail");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data row – each cell contains the <<cellMerge>> tag.
        // Cells that have identical text will be merged horizontally.
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[row.Cell1]>>");
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[row.Cell2]>>");
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[row.Cell3]>>");
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[row.Cell4]>>");
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[row.Cell5]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model object. The root name in the template is "Model".
        engine.BuildReport(reportDoc, model, "Model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<RowData> Rows { get; set; } = new();
}

public class RowData
{
    public string Cell1 { get; set; } = string.Empty;
    public string Cell2 { get; set; } = string.Empty;
    public string Cell3 { get; set; } = string.Empty;
    public string Cell4 { get; set; } = string.Empty;
    public string Cell5 { get; set; } = string.Empty;
}
