using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            builder.StartTable();

            // ---- First row, first cell ----
            // This cell will be the first cell in a 2x2 merged region.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;   // Start horizontal merge.
            builder.CellFormat.VerticalMerge = CellMerge.First;     // Start vertical merge.
            builder.Write("Merged 2x2 cell");

            // ---- First row, second cell ----
            // Merge this cell horizontally with the cell to its left.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge.
            builder.CellFormat.VerticalMerge = CellMerge.None;       // No vertical merge for this cell.
            // No text needed; merged cells display content from the first cell.

            // End the first row.
            builder.EndRow();

            // ---- Second row, first cell ----
            // Merge this cell vertically with the cell above.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;    // No horizontal merge for this cell.
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue vertical merge.
            // No text needed; merged cells display content from the first cell.

            // ---- Second row, second cell ----
            // This cell remains independent.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Normal cell");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document in DOCX format.
            doc.Save("MergedCellsBothDirections.docx");
        }
    }
}
