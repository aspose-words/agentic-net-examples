using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDiagonal
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and insert the first cell.
            builder.StartTable();
            builder.InsertCell();

            // Define the appearance of the diagonal borders.
            // DiagonalDown border.
            builder.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
            builder.CellFormat.Borders[BorderType.DiagonalDown].LineWidth = 1.5;
            builder.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Red;

            // DiagonalUp border.
            builder.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
            builder.CellFormat.Borders[BorderType.DiagonalUp].LineWidth = 1.5;
            builder.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Red;

            // Add some text to the cell.
            builder.Write("Cell with diagonal lines");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Ensure the output directory exists and save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DiagonalCell.docx");
            doc.Save(outputPath);
        }
    }
}
