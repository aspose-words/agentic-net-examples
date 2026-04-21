using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace AsposeWordsTableDiagonal
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Insert the first cell where we will apply diagonal lines.
            builder.InsertCell();

            // Configure the diagonal borders of the current cell.
            // Set a single line style for both diagonal directions.
            builder.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
            builder.CellFormat.Borders[BorderType.DiagonalDown].Color = System.Drawing.Color.Red;
            builder.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
            builder.CellFormat.Borders[BorderType.DiagonalUp].Color = System.Drawing.Color.Blue;

            // Add some text to the cell.
            builder.Write("Cell with diagonal lines");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DiagonalCell.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
