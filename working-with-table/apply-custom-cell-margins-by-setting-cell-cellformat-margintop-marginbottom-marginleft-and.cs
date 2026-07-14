using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMarginsExample
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

            // First row, first cell.
            Cell cell1 = builder.InsertCell();
            // Apply custom paddings to the cell (these act as margins inside the cell).
            cell1.CellFormat.TopPadding = 5;      // 5 points top padding
            cell1.CellFormat.BottomPadding = 5;   // 5 points bottom padding
            cell1.CellFormat.LeftPadding = 10;    // 10 points left padding
            cell1.CellFormat.RightPadding = 10;   // 10 points right padding
            builder.Write("Cell 1");

            // First row, second cell.
            Cell cell2 = builder.InsertCell();
            cell2.CellFormat.TopPadding = 8;
            cell2.CellFormat.BottomPadding = 8;
            cell2.CellFormat.LeftPadding = 12;
            cell2.CellFormat.RightPadding = 12;
            builder.Write("Cell 2");

            // End the first row.
            builder.EndRow();

            // Second row, first cell.
            Cell cell3 = builder.InsertCell();
            cell3.CellFormat.TopPadding = 3;
            cell3.CellFormat.BottomPadding = 3;
            cell3.CellFormat.LeftPadding = 6;
            cell3.CellFormat.RightPadding = 6;
            builder.Write("Cell 3");

            // Second row, second cell.
            Cell cell4 = builder.InsertCell();
            cell4.CellFormat.TopPadding = 4;
            cell4.CellFormat.BottomPadding = 4;
            cell4.CellFormat.LeftPadding = 8;
            cell4.CellFormat.RightPadding = 8;
            builder.Write("Cell 4");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = "CustomCellMargins.docx";
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
