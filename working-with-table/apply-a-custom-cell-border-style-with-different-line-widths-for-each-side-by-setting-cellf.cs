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

        // Begin a table.
        builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Retrieve the cell that was just created.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a cell.");

        // Apply custom borders with different line widths.
        // Top border – 1 point.
        cell.CellFormat.Borders.Top.LineWidth = 1.0;
        cell.CellFormat.Borders.Top.Color = Color.Black;
        cell.CellFormat.Borders.Top.LineStyle = LineStyle.Single;

        // Bottom border – 2 points.
        cell.CellFormat.Borders.Bottom.LineWidth = 2.0;
        cell.CellFormat.Borders.Bottom.Color = Color.Black;
        cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;

        // Left border – 3 points.
        cell.CellFormat.Borders.Left.LineWidth = 3.0;
        cell.CellFormat.Borders.Left.Color = Color.Black;
        cell.CellFormat.Borders.Left.LineStyle = LineStyle.Single;

        // Right border – 4 points.
        cell.CellFormat.Borders.Right.LineWidth = 4.0;
        cell.CellFormat.Borders.Right.Color = Color.Black;
        cell.CellFormat.Borders.Right.LineStyle = LineStyle.Single;

        // Add some text to the cell.
        builder.Writeln("Custom borders");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "CustomCellBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
