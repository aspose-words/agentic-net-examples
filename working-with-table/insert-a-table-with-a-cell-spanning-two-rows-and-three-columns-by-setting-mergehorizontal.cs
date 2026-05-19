using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // -------------------------------------------------
            // First row – the cell that will span 2 rows and 3 columns.
            // -------------------------------------------------
            // Insert the top‑left cell and mark it as the first cell in both
            // a horizontal and a vertical merge range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.CellFormat.VerticalMerge   = CellMerge.First;
            builder.Write("Spanned cell");

            // Insert the next two cells in the first row.
            // They are merged horizontally with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge   = CellMerge.None; // No vertical merge for these cells.

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge   = CellMerge.None;

            // End the first row.
            builder.EndRow();

            // -------------------------------------------------
            // Second row – continuation of the merged cell.
            // -------------------------------------------------
            // Insert the first cell of the second row.
            // It continues both the horizontal and vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge   = CellMerge.Previous;

            // Insert the remaining two cells of the second row.
            // They continue the horizontal merge but have no vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge   = CellMerge.None;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge   = CellMerge.None;

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to a file.
            doc.Save("MergedTable.docx");
        }
    }
}
