using System;
using System.IO;
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
        builder.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Black;
        builder.CellFormat.Borders[BorderType.DiagonalDown].LineWidth = 1.0;

        // Apply a diagonal line from bottom‑left to top‑right.
        builder.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
        builder.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Black;
        builder.CellFormat.Borders[BorderType.DiagonalUp].LineWidth = 1.0;

        // Add some text to the cell.
        builder.Write("Diagonal lines");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DiagonalCell.docx");
        doc.Save(outputPath);
    }
}
