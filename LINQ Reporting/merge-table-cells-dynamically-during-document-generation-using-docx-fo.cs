using System;
using System.Collections.Generic;
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

            // Use DocumentBuilder to construct the document content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // Example data that determines how cells should be merged.
            // Each tuple: (cell text, merge horizontally?, merge vertically?)
            var cellData = new List<(string Text, bool MergeHorizontally, bool MergeVertically)>
            {
                ("Header 1", true, false),
                ("Header 2", true, false),
                ("Data A1", false, true),
                ("Data A2", false, false),
                ("Data B1", false, true),
                ("Data B2", false, false)
            };

            // Build the first row (header) with horizontal merge across the first two cells.
            // First cell: mark as the first cell in a horizontally merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write(cellData[0].Text);

            // Second cell: mark as merged to the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for merged cell.
            builder.EndRow();

            // Build the second row with vertical merges.
            // First column: first cell of a vertically merged range.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write(cellData[2].Text);
            // Second column: independent cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write(cellData[3].Text);
            builder.EndRow();

            // Third row: continue vertical merge for the first column.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            // No text needed; it merges with the cell above.
            // Second column: independent cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write(cellData[5].Text);
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document in DOCX format.
            doc.Save("MergedTable.docx");
        }
    }
}
