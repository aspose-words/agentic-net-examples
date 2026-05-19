using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

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

        // Define the appearance of the diagonal borders.
        // Enable the diagonal borders by setting their style, width and color.
        builder.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
        builder.CellFormat.Borders[BorderType.DiagonalDown].LineWidth = 1.5;
        builder.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Red;

        builder.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
        builder.CellFormat.Borders[BorderType.DiagonalUp].LineWidth = 1.5;
        builder.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Red;

        // Add some text to the cell.
        builder.Write("Cell with diagonal lines");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local folder.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DiagonalCell.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
