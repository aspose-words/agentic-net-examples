using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Insert the first cell where we will apply custom borders.
        builder.InsertCell();

        // Set different line widths, styles, and colors for each side of the cell.
        builder.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Left.LineWidth = 1.0; // 1 point
        builder.CellFormat.Borders.Left.Color = Color.Red;

        builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Right.LineWidth = 2.0; // 2 points
        builder.CellFormat.Borders.Right.Color = Color.Green;

        builder.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Top.LineWidth = 3.0; // 3 points
        builder.CellFormat.Borders.Top.Color = Color.Blue;

        builder.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Bottom.LineWidth = 4.0; // 4 points
        builder.CellFormat.Borders.Bottom.Color = Color.Orange;

        // Add some text to the custom‑bordered cell.
        builder.Writeln("Custom bordered cell");

        // Insert a second cell with default formatting for comparison.
        builder.InsertCell();
        builder.Writeln("Normal cell");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        doc.Save("CustomCellBorders.docx");
    }
}
