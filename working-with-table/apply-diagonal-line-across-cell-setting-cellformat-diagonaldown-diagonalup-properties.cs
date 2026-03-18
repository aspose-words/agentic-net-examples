using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        Cell cell = builder.InsertCell();
        builder.Write("Cell with diagonal borders");

        // Configure the diagonal-down border.
        cell.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
        cell.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Blue;
        cell.CellFormat.Borders[BorderType.DiagonalDown].LineWidth = 2.0;

        // Configure the diagonal-up border.
        cell.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
        cell.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Red;
        cell.CellFormat.Borders[BorderType.DiagonalUp].LineWidth = 2.0;

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("DiagonalCell.docx");
    }
}
