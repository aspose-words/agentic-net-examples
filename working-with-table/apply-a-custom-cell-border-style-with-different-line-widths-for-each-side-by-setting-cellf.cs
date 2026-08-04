using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableBordersExample
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

            // Insert the first cell.
            builder.InsertCell();

            // Retrieve the cell that was just created.
            Cell firstCell = builder.CurrentParagraph.ParentNode as Cell;
            if (firstCell == null)
                throw new InvalidOperationException("Unable to obtain the first cell.");

            // Apply custom borders with different line widths to each side of the cell.
            // Left border: 2 points, solid black.
            firstCell.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
            firstCell.CellFormat.Borders.Left.LineWidth = 2.0;
            firstCell.CellFormat.Borders.Left.Color = Color.Black;

            // Right border: 4 points, solid black.
            firstCell.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
            firstCell.CellFormat.Borders.Right.LineWidth = 4.0;
            firstCell.CellFormat.Borders.Right.Color = Color.Black;

            // Top border: 1 point, solid black.
            firstCell.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
            firstCell.CellFormat.Borders.Top.LineWidth = 1.0;
            firstCell.CellFormat.Borders.Top.Color = Color.Black;

            // Bottom border: 3 points, solid black.
            firstCell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
            firstCell.CellFormat.Borders.Bottom.LineWidth = 3.0;
            firstCell.CellFormat.Borders.Bottom.Color = Color.Black;

            // Add some text to the first cell.
            builder.Writeln("Cell with custom borders");

            // Insert a second cell with default borders for comparison.
            builder.InsertCell();
            builder.Writeln("Cell with default borders");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "CustomCellBorders.docx";
            doc.Save(outputPath);
        }
    }
}
