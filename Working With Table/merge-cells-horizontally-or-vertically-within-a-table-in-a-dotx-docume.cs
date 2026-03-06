using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableCellMergingExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTX template that already contains a table.
            Document doc = new Document("Template.dotx");

            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // -------------------------------------------------
            // Horizontal merge: merge the first two cells of the first row.
            // -------------------------------------------------
            // The leftmost cell becomes the first cell in the merged range.
            Cell firstCell = table.Rows[0].Cells[0];
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;

            // The cell to the right is merged with the previous cell.
            Cell secondCell = table.Rows[0].Cells[1];
            secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

            // -------------------------------------------------
            // Vertical merge: merge the first two cells of the first column.
            // -------------------------------------------------
            // The top cell becomes the first cell in the vertical merge range.
            firstCell.CellFormat.VerticalMerge = CellMerge.First;

            // The cell directly below merges with the cell above.
            Cell cellBelow = table.Rows[1].Cells[0];
            cellBelow.CellFormat.VerticalMerge = CellMerge.Previous;

            // Save the modified document.
            doc.Save("MergedTable.docx");
        }
    }
}
