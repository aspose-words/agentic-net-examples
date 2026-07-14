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

        // Start a table. The builder will apply the current CellFormat to each new cell.
        Table table = builder.StartTable();

        // Reset any previous cell formatting.
        builder.CellFormat.ClearFormatting();

        // Define custom borders for the upcoming cell.
        // Left border: 2 points, red.
        builder.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Left.LineWidth = 2.0;
        builder.CellFormat.Borders.Left.Color = Color.Red;

        // Right border: 4 points, green.
        builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Right.LineWidth = 4.0;
        builder.CellFormat.Borders.Right.Color = Color.Green;

        // Top border: 6 points, blue.
        builder.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Top.LineWidth = 6.0;
        builder.CellFormat.Borders.Top.Color = Color.Blue;

        // Bottom border: 8 points, purple.
        builder.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Bottom.LineWidth = 8.0;
        builder.CellFormat.Borders.Bottom.Color = Color.Purple;

        // Insert a single cell that will receive the custom borders.
        builder.InsertCell();
        builder.Write("Cell with custom borders");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local file system.
        doc.Save("CustomCellBorders.docx");
    }
}
