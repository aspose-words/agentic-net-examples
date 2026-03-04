using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeDemo
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

            // -------------------------------------------------
            // Row 1 – Demonstrate horizontal merging of cells.
            // -------------------------------------------------
            // First cell – mark as the first cell in a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged horizontally across 3 cells");

            // Second cell – merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell – also merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Fourth cell – normal, not merged.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Normal cell");

            // End the first row.
            builder.EndRow();

            // -------------------------------------------------
            // Row 2 – Demonstrate vertical merging of cells.
            // -------------------------------------------------
            // First column – start a vertically merged range.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Merged vertically across 2 rows");

            // Second column – normal cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row2, Col2");

            // Third column – normal cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row2, Col3");

            // Fourth column – normal cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row2, Col4");

            // End the second row.
            builder.EndRow();

            // -------------------------------------------------
            // Row 3 – Continue vertical merge started in Row 2.
            // -------------------------------------------------
            // First column – merge with the cell above.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            // The rest of the cells are normal.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row3, Col2");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row3, Col3");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row3, Col4");

            // End the third row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document in the legacy DOC format.
            doc.Save("MergedTableDemo.doc", SaveFormat.Doc);
        }
    }
}
