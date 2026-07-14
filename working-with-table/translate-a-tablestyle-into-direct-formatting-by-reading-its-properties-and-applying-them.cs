using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style and set formatting properties.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
        customStyle.CellSpacing = 5; // space between cells
        customStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite; // table background
        customStyle.Borders.Color = Color.Blue; // border color
        customStyle.Borders.LineStyle = LineStyle.DotDash; // border style
        customStyle.RowStripe = 2; // optional row banding
        customStyle.ColumnStripe = 2; // optional column banding

        // Apply the style to the table.
        table.Style = customStyle;

        // Convert the style's formatting to direct formatting on the table, rows, and cells.
        doc.ExpandTableStylesToDirectFormatting();

        // Save the resulting document.
        doc.Save("TableStyleToDirectFormatting.docx");
    }
}
