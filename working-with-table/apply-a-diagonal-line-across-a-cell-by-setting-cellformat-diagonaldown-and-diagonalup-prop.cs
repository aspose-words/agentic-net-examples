using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace AsposeWordsDiagonalCellExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            builder.StartTable();

            // Insert the first cell and add some text.
            builder.InsertCell();
            builder.Writeln("Diagonal Cell");

            // Retrieve the cell that was just created.
            Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

            // Apply diagonal lines across the cell using the Borders collection.
            // Diagonal from top‑left to bottom‑right.
            cell.CellFormat.Borders[BorderType.DiagonalDown].LineStyle = LineStyle.Single;
            cell.CellFormat.Borders[BorderType.DiagonalDown].Color = Color.Black;

            // Diagonal from bottom‑left to top‑right.
            cell.CellFormat.Borders[BorderType.DiagonalUp].LineStyle = LineStyle.Single;
            cell.CellFormat.Borders[BorderType.DiagonalUp].Color = Color.Black;

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            string outputPath = "DiagonalCell.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved successfully.");
        }
    }
}
