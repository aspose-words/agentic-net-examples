using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        builder.StartTable();

        // Insert a single cell and add some text.
        builder.InsertCell();
        builder.Writeln("Custom Borders");

        // Get the current cell that was just created.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a table cell.");

        // Apply different line widths, colors and styles to each side of the cell.
        cell.CellFormat.Borders.Left.LineWidth = 1.0;   // 1 point left border
        cell.CellFormat.Borders.Left.Color = Color.Blue;
        cell.CellFormat.Borders.Left.LineStyle = LineStyle.Single;

        cell.CellFormat.Borders.Top.LineWidth = 2.0;    // 2 points top border
        cell.CellFormat.Borders.Top.Color = Color.Green;
        cell.CellFormat.Borders.Top.LineStyle = LineStyle.Single;

        cell.CellFormat.Borders.Right.LineWidth = 3.0;  // 3 points right border
        cell.CellFormat.Borders.Right.Color = Color.Red;
        cell.CellFormat.Borders.Right.LineStyle = LineStyle.Single;

        cell.CellFormat.Borders.Bottom.LineWidth = 4.0; // 4 points bottom border
        cell.CellFormat.Borders.Bottom.Color = Color.Purple;
        cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputPath = "CustomCellBorders.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
