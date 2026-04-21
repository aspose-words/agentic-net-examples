using System;
using System.IO;
using System.Drawing;               // Needed for Color
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

        // Insert the first cell.
        Cell cell = builder.InsertCell();

        // Apply custom borders with different line widths and colors.
        cell.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
        cell.CellFormat.Borders.Left.LineWidth = 2.0;
        cell.CellFormat.Borders.Left.Color = Color.Red;

        cell.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
        cell.CellFormat.Borders.Right.LineWidth = 4.0;
        cell.CellFormat.Borders.Right.Color = Color.Green;

        cell.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
        cell.CellFormat.Borders.Top.LineWidth = 6.0;
        cell.CellFormat.Borders.Top.Color = Color.Blue;

        cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
        cell.CellFormat.Borders.Bottom.LineWidth = 8.0;
        cell.CellFormat.Borders.Bottom.Color = Color.Purple;

        // Add some text to the cell.
        builder.Writeln("Cell with custom borders");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "CustomCellBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
