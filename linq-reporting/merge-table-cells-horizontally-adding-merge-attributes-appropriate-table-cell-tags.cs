using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMerge
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with three columns.
            builder.StartTable();

            // ---- First row: merge first two cells horizontally ----
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell 1-2");

            // Insert the second cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for the merged cell.

            // Insert the third cell (unmerged).
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Cell 3");

            // End the first row.
            builder.EndRow();

            // ---- Second row: normal unmerged cells ----
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Row2 Cell1");
            builder.InsertCell();
            builder.Write("Row2 Cell2");
            builder.InsertCell();
            builder.Write("Row2 Cell3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            doc.Save("TableWithHorizontalMerge.docx");
        }
    }
}
