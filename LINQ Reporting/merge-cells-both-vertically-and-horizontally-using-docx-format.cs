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
            Table table = builder.StartTable();

            // ---- First Row ----
            // Cell (0,0) – first cell of a 2x2 merged block.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;   // Start horizontal merge.
            builder.CellFormat.VerticalMerge = CellMerge.First;     // Start vertical merge.
            builder.Write("Merged 2x2 cell");

            // Cell (0,1) – continues horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with left cell.
            builder.CellFormat.VerticalMerge = CellMerge.None;       // No vertical merge for this cell.
            // No text needed; merged cells must be empty.

            // End first row.
            builder.EndRow();

            // ---- Second Row ----
            // Cell (1,0) – continues vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;     // No horizontal merge for this cell.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;   // Merge with cell above.
            // No text needed; merged cells must be empty.

            // Cell (1,1) – continues both horizontal and vertical merges.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with left cell.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;   // Merge with cell above.
            // No text needed; merged cells must be empty.

            // End second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document in DOCX format.
            doc.Save("MergedCells.docx");
        }
    }
}
