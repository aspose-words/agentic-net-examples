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

        // Start a table and insert a single cell.
        builder.StartTable();
        builder.InsertCell();

        // Apply a diagonal line from top‑left to bottom‑right.
        builder.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
        builder.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Red;
        builder.CellFormat.Borders[BorderType.DiagonalDown].LineWidth = 2.0;

        // Apply a diagonal line from bottom‑left to top‑right.
        builder.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
        builder.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Red;
        builder.CellFormat.Borders[BorderType.DiagonalUp].LineWidth = 2.0;

        // Add some text to the cell.
        builder.Write("Diagonal");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        doc.Save("DiagonalCell.docx");
    }
}
