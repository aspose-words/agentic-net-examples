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

        // Start a table and add a single cell with some text.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Diagonal Cell");

        // Retrieve the cell that was just created.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Unable to obtain the created cell.");

        // Apply diagonal lines across the cell using border types.
        cell.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
        cell.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Black;
        cell.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
        cell.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Black;

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "DiagonalCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
