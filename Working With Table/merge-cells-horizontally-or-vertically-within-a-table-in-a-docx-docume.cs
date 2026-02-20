using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableCellMergingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Build a table and merge cells horizontally.
            // -------------------------------------------------
            Table table = builder.StartTable();

            // First row – merge first two cells horizontally.
            builder.InsertCell();
            // Mark this cell as the first cell in a horizontal merge range.
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Horizontally merged cells (1-2)");

            builder.InsertCell();
            // This cell will be merged with the previous cell.
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for the merged cell.
            builder.InsertCell(); // Insert third cell (not merged).
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Normal cell");
            builder.EndRow();

            // Second row – normal cells (no merging).
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.InsertCell();
            builder.Write("Row 2, Cell 3");
            builder.EndRow();

            // -------------------------------------------------
            // Merge cells vertically.
            // -------------------------------------------------
            // Add a third row where the first cell will be merged vertically with the cell above.
            builder.InsertCell();
            // Mark this cell as the first cell in a vertical merge range.
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Vertically merged cells (Row 3-4, Col 1)");

            builder.InsertCell();
            builder.Write("Row 3, Cell 2");
            builder.InsertCell();
            builder.Write("Row 3, Cell 3");
            builder.EndRow();

            // Fourth row – continue the vertical merge.
            builder.InsertCell();
            // This cell merges with the previous cell vertically.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            // No text needed for the merged cell.
            builder.InsertCell();
            builder.Write("Row 4, Cell 2");
            builder.InsertCell();
            builder.Write("Row 4, Cell 3");
            builder.EndRow();

            builder.EndTable();

            // Save the document with merged cells.
            doc.Save("MergedCellsTable.docx");

            // -------------------------------------------------
            // Example of loading an existing document and merging cells.
            // -------------------------------------------------
            Document loadedDoc = new Document("MergedCellsTable.docx");
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];

            // Merge the second and third cells of the first row horizontally.
            Cell firstRowSecondCell = loadedTable.Rows[0].Cells[1];
            Cell firstRowThirdCell = loadedTable.Rows[0].Cells[2];
            firstRowSecondCell.CellFormat.HorizontalMerge = CellMerge.First;
            firstRowThirdCell.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Merge the second column cells vertically (rows 1 and 2).
            Cell secondRowSecondCell = loadedTable.Rows[1].Cells[1];
            secondRowSecondCell.CellFormat.VerticalMerge = CellMerge.First;
            Cell thirdRowSecondCell = loadedTable.Rows[2].Cells[1];
            thirdRowSecondCell.CellFormat.VerticalMerge = CellMerge.Previous;

            // Save the modified document.
            loadedDoc.Save("MergedCellsTable_Modified.docx");
        }
    }
}
